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

' --- ACTUALIZAR STOCK DE PRODUCTO (desde importación) ----------------------

Public Function ActualizarStockProducto(codigoProducto As String, cantidad As Double) As Boolean
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim i As Long
    Dim stockActual As Double
    
    Set ws = ThisWorkbook.Sheets(HOJA_ALMACEN)
    
    For i = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
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

' --- FUNCIÓN DE HOJA: ACTUALIZAR ALMACÉN DESDE VENTAS ---------------------

Public Function ObtenerStockRestante(codigo As String) As Double
    ' Calcula el stock disponible restando lo vendido y las mermas
    Dim celda As Range
    Dim cantidadVendida As Double
    Dim mermas As Double
    
    cantidadVendida = 0
    
    ' Sumar todas las ventas de este producto
    For Each celda In Hoja8.Range("B5:B" & Hoja8.Range("B" & Rows.Count).End(xlUp).Row)
        If CStr(celda) = CStr(codigo) Then
            cantidadVendida = cantidadVendida + Hoja8.Range("E" & celda.Row).Value
        End If
    Next celda
    
    ' Obtener mermas registradas
    mermas = ObtenerMermasProducto(codigo)
    
    ' Stock inicial - vendido - mermas = stock actual
    ObtenerStockRestante = Hoja4.Range("E" & Application.Caller.Row) - cantidadVendida - mermas
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

' --- FORZAR ACTUALIZACIÓN DE FÓRMULAS DE STOCK ----------------------------
' Vuelve a escribir la fórmula en las celdas para forzar recálculo.

Public Sub ForzarActualizacionStock()
    Dim ultima_fila As Integer
    Dim i As Integer
    
    ultima_fila = Hoja4.Range("B" & Rows.Count).End(xlUp).Row
    
    For i = 5 To ultima_fila
        ' Forzar recálculo de la función ObtenerStockRestante
        Hoja4.Range("F" & i).Formula = "=ObtenerStockRestante([@Código])"
        
        ' Recalcular columnas derivadas
        Hoja4.Range("I" & i).Value = Hoja4.Range("F" & i).Value * Hoja4.Range("H" & i).Value
        Hoja4.Range("J" & i).Value = Hoja4.Range("F" & i).Value * Hoja4.Range("G" & i).Value
        Hoja4.Range("K" & i).Value = Hoja4.Range("J" & i).Value - Hoja4.Range("I" & i).Value
    Next i
End Sub
