Attribute VB_Name = "modStock"
' ============================================================================
' MÓDULO: modStock
' PROPÓSITO: Consolidación de funciones de actualización de stock en el
' inventario. Reemplaza Restar_Cantidad (duplicado 3 veces), setear_funcion,
' actualizar_almacen, ActualizarStockProducto, ReabastecerProducto.
' ============================================================================

Option Explicit

' --- RESTAR CANTIDAD DEL INVENTARIO ----------------------------------------
' Busca el producto por código y resta la cantidad vendida.
' Actualiza también fondo, ingreso esperado y ganancia esperada.

Public Sub RestarCantidadInventario(codigo As String, cantidad As Double)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim fila As ListRow
    
    Set ws = ThisWorkbook.Worksheets(HOJA_ALMACEN)
    Set tbl = ws.ListObjects(TBL_INVENTARIO)
    
    For Each fila In tbl.ListRows
        If codigo = fila.Range(COL_INV_CODIGO) Then
            ' Restar cantidad actual
            fila.Range(COL_INV_CANT_ACT) = fila.Range(COL_INV_CANT_ACT) - cantidad
            
            ' Recalcular columnas financieras
            fila.Range(COL_INV_FONDO) = fila.Range(COL_INV_CANT_ACT) * fila.Range(COL_INV_PRECIO_C)
            fila.Range(COL_INV_INGRESO) = fila.Range(COL_INV_CANT_ACT) * fila.Range(COL_INV_PRECIO_V)
            fila.Range(COL_INV_GANANCIA) = fila.Range(COL_INV_INGRESO) - fila.Range(COL_INV_FONDO)
            Exit Sub
        End If
    Next fila
End Sub

' --- ACTUALIZAR STOCK DE PRODUCTO (desde importaciï¿½n) ----------------------

Public Function ActualizarStockProducto(codigoProducto As String, cantidad As Double) As Boolean
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim i As Long
    Dim stockActual As Double
    
    Set ws = ThisWorkbook.Sheets(HOJA_ALMACEN)
    
    For i = 2 To ws.Cells(ws.Rows.count, 1).End(xlUp).Row
        If ws.Cells(i, 2).Value = codigoProducto Then
            stockActual = ws.Cells(i, 6).Value
            ws.Cells(i, 6).Value = stockActual - cantidad
            ActualizarStockProducto = True
            Exit Function
        End If
    Next i
    
    ActualizarStockProducto = False
    Exit Function
    
ErrorHandler:
    MsgBox "Error al actualizar el stock del producto." & vbCrLf & _
           "Producto: " & codigoProducto & vbCrLf & _
           "Cantidad: " & cantidad & vbCrLf & _
           "Detalle: " & Err.Description, vbExclamation
    ActualizarStockProducto = False
End Function

' --- FUNCIï¿½N DE HOJA: ACTUALIZAR ALMACÉN DESDE VENTAS ---------------------

Public Function ObtenerStockRestante(codigo As String, stockInicial As Double) As Double
    Dim celda As Range
    Dim cantidadVendida As Double
    Dim mermas As Double
    
    cantidadVendida = 0
    
    ' Sumar todas las ventas de este producto
    For Each celda In Hoja8.Range("B5:B" & Hoja8.Range("B" & Rows.count).End(xlUp).Row)
        If CStr(celda) = CStr(codigo) Then
            cantidadVendida = cantidadVendida + Hoja8.Range("E" & celda.Row).Value
        End If
    Next celda
    
    ' Obtener mermas registradas
    mermas = modMermas.ObtenerMermasProducto(codigo)
    
    ' Stock inicial - vendido - mermas = stock actual
    ObtenerStockRestante = stockInicial - cantidadVendida - mermas
End Function

' --- REABASTECER PRODUCTO --------------------------------------------------

Public Sub ReabastecerProducto()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim rng As Range
    Dim filaSeleccionada As Range
    Dim cantidadStr As String
    Dim cantidad As Double
    Dim stockActual As Double
    Dim stockInicial As Double
    Dim filaTabla As Long
    
    Set ws = ThisWorkbook.Worksheets(HOJA_ALMACEN)
    Set tbl = ws.ListObjects(TBL_INVENTARIO)
    
    ' Validar selección
    If TypeName(Selection) <> "Range" Then
        MsgBox "Por favor seleccione una fila de la tabla INVENTARIO.", vbExclamation
        Exit Sub
    End If
    
    If tbl.DataBodyRange Is Nothing Then
        MsgBox "No hay productos en el almacén.", vbExclamation
        Exit Sub
    End If
    
    Set rng = Intersect(Selection, tbl.DataBodyRange)
    If rng Is Nothing Then
        MsgBox "Debe seleccionar una fila de la tabla INVENTARIO.", vbExclamation
        Exit Sub
    End If
    
    Set filaSeleccionada = rng.Rows(1)
    filaTabla = filaSeleccionada.Row - tbl.HeaderRowRange.Row
    
    stockActual = tbl.DataBodyRange.Cells(filaTabla, COL_INV_CANT_ACT).Value
    stockInicial = tbl.DataBodyRange.Cells(filaTabla, COL_INV_CANT_INI).Value
    
    cantidadStr = InputBox("Ingrese la cantidad a añadir para reabastecer:" & vbCrLf & _
                           "Producto: " & tbl.DataBodyRange.Cells(filaTabla, COL_INV_NOMBRE).Value & vbCrLf & _
                           "Stock actual: " & stockActual, _
                           "Reabastecer Producto", "0")
    
    If cantidadStr = "" Then Exit Sub
    If Not IsNumeric(cantidadStr) Or Val(cantidadStr) <= 0 Then
        MsgBox "Debe ingresar una cantidad válida mayor a cero.", vbExclamation
        Exit Sub
    End If
    
    cantidad = Val(cantidadStr)
    
    ' Sumar a cantidad inicial
    tbl.DataBodyRange.Cells(filaTabla, COL_INV_CANT_INI).Value = _
        Format(Round(stockInicial + cantidad, 3), "General Number")
    tbl.DataBodyRange.Cells(filaTabla, COL_INV_FECHA).Value = Date
    
    Call ForzarActualizacionStock
    
    MsgBox "Reabastecimiento exitoso." & vbCrLf & _
           "Stock anterior: " & stockActual & vbCrLf & _
           "Cantidad añadida: " & cantidad & vbCrLf & _
                       "Nuevo stock: " & (stockActual + cantidad), vbInformation
End Sub

' --- BUSCAR PRODUCTOS EN ALMACEN POR NOMBRE ---------------------------------
' Busca TODOS los productos cuyo nombre coincida exactamente con nombreProducto.
' Devuelve un array 2D (1-based): {Codigo, Precio_V, CANT_INI, CANT_ACT, Precio_C}.
' Si no hay coincidencias, devuelve un array vacio (UBound falla).
' Util para detectar si un producto elaborado ya existe y para mostrar opciones.

Public Function BuscarProductosEnAlmacenPorNombre(ByVal nombreProducto As String) As Variant
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim fila As ListRow
    Dim resultados() As Variant
    Dim count As Long
    
    Set ws = ThisWorkbook.Sheets(HOJA_ALMACEN)
    Set tbl = ws.ListObjects(TBL_INVENTARIO)
    
    ' Primer pase: contar coincidencias
    count = 0
    For Each fila In tbl.ListRows
        If UCase(Trim(fila.Range.Cells(1, COL_INV_NOMBRE).Value)) = UCase(Trim(nombreProducto)) Then
            count = count + 1
        End If
    Next fila
    
    If count = 0 Then
        BuscarProductosEnAlmacenPorNombre = Array()
        Exit Function
    End If
    
    ' Segundo pase: llenar array
    ReDim resultados(1 To count, 1 To 5)
    count = 0
    For Each fila In tbl.ListRows
        If UCase(Trim(fila.Range.Cells(1, COL_INV_NOMBRE).Value)) = UCase(Trim(nombreProducto)) Then
            count = count + 1
            resultados(count, 1) = fila.Range.Cells(1, COL_INV_CODIGO).Value
            resultados(count, 2) = fila.Range.Cells(1, COL_INV_PRECIO_V).Value
            resultados(count, 3) = fila.Range.Cells(1, COL_INV_CANT_INI).Value
            resultados(count, 4) = fila.Range.Cells(1, COL_INV_CANT_ACT).Value
            resultados(count, 5) = fila.Range.Cells(1, COL_INV_PRECIO_C).Value
        End If
    Next fila
    
    BuscarProductosEnAlmacenPorNombre = resultados
End Function

' --- FORZAR ACTUALIZACIÓN DE FÓRMULAS DE STOCK ----------------------------
' Vuelve a calcular y escribir el stock actual de cada producto

Public Sub ForzarActualizacionStock()
    Dim ultima_fila As Long
    Dim i As Long
    Dim codigo As String
    Dim stockInicial As Double
    Dim stockRestante As Double
    
    Application.ScreenUpdating = False
    
    ultima_fila = Hoja4.Range("B" & Rows.count).End(xlUp).Row
    
    For i = 5 To ultima_fila
        codigo = CStr(Hoja4.Range("B" & i).Value)
        stockInicial = CDbl(Hoja4.Range("E" & i).Value)
        
        ' Calcular y escribir VALOR, no fórmula
        stockRestante = ObtenerStockRestante(codigo, stockInicial)
        Hoja4.Range("F" & i).Value = stockRestante
        
        ' Columnas derivadas como valores directos
        Hoja4.Range("I" & i).Value = stockRestante * Hoja4.Range("H" & i).Value  ' Fondo
        Hoja4.Range("J" & i).Value = stockRestante * Hoja4.Range("G" & i).Value  ' Ingreso
        Hoja4.Range("K" & i).Value = Hoja4.Range("J" & i).Value - Hoja4.Range("I" & i).Value  ' Ganancia
    Next i
    
    Application.ScreenUpdating = True
End Sub

