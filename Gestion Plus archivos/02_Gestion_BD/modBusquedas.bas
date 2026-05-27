Attribute VB_Name = "modBusquedas"
' ============================================================================
' MÓDULO: modBusquedas
' PROPÓSITO: Funciones de búsqueda y verificación de repeticiones.
'            Versión desacoplada: recibe controles como parámetros.
' ============================================================================

Option Explicit

' --- COMPROBAR REPETICIÓN EN FACTURACIÓN (listbox) -------------------------
' Recibe el código a verificar y el listbox como parámetros.

Public Function ComprobarRepeticionFacturacion(ByVal codigo As String, _
                                                ByRef lstProductos As MSForms.ListBox) As Boolean
    Dim i As Integer
    
    ComprobarRepeticionFacturacion = False
    
    For i = 0 To lstProductos.ListCount - 1
        If codigo = lstProductos.List(i, 0) Then
            ComprobarRepeticionFacturacion = True
            Exit Function
        End If
    Next i
End Function

' --- COMPROBAR REPETICIÓN EN INGREDIENTE NUEVO ----------------------------

Public Sub ComprobarRepeticionIngredienteNuevo()
    Dim i As Integer
    Dim ultima_fila As Integer
    Dim HojaBD As Worksheet
    Dim nombre_producto As String
    
    Set HojaBD = ThisWorkbook.Sheets(HOJA_INGREDIENTES)
    ultima_fila = HojaBD.Range("B" & Rows.Count).End(xlUp).Row
    nombre_producto = ingresar_ingr_nuevo_form.Nombre_ingrediente_Nuevo_TextBox.Value
    
    repeticion2 = False
    
    For i = 5 To ultima_fila
        If nombre_producto = HojaBD.Range("B" & i).Value Then
            repeticion2 = True
            Exit For
        End If
    Next i
End Sub

' --- COMPROBAR REPETICIÓN EN INGREDIENTE MODIFICAR ------------------------
' Permite usar el mismo nombre del elemento que se está modificando.

Public Sub ComprobarRepeticionIngredienteModificar()
    Dim i As Integer
    Dim ultima_fila As Integer
    Dim fila_activa As Integer
    Dim nombre_nuevo As String
    Dim nombre_actual As String
    Dim HojaBD As Worksheet
    
    Set HojaBD = ThisWorkbook.Sheets(HOJA_INGREDIENTES)
    ultima_fila = HojaBD.Cells(HojaBD.Rows.Count, "B").End(xlUp).Row
    fila_activa = ActiveCell.Row
    nombre_nuevo = ingresar_ingr_nuevo_form.Nombre_ingrediente_Nuevo_TextBox.Value
    nombre_actual = HojaBD.Range("B" & fila_activa).Value
    
    repeticion3 = False
    
    For i = 5 To ultima_fila
        If HojaBD.Range("B" & i).Value = nombre_nuevo And i <> fila_activa Then
            repeticion3 = True
            Exit For
        End If
    Next i
End Sub

' --- CARGAR TODOS LOS PRODUCTOS EN LISTBOX ----------------------------------
' Llena un listbox con todos los productos del inventario (col: Código, Nombre, Disponibilidad, Precio).

Public Sub CargarTodosLosProductos(ByRef lstResultados As MSForms.ListBox)
    Dim wsAlmacen As Worksheet
    Dim tblInv As ListObject
    Dim fila As ListRow
    Dim i As Long
    
    Set wsAlmacen = ThisWorkbook.Sheets(HOJA_ALMACEN)
    Set tblInv = wsAlmacen.ListObjects(TBL_INVENTARIO)
    
    lstResultados.Clear
    i = 0
    
    For Each fila In tblInv.ListRows
        lstResultados.AddItem
        lstResultados.List(i, 0) = fila.Range.Cells(1, COL_INV_CODIGO).Value
        lstResultados.List(i, 1) = fila.Range.Cells(1, COL_INV_NOMBRE).Value
        lstResultados.List(i, 2) = modFormateo.FormatearNumero(fila.Range.Cells(1, COL_INV_CANT_ACT).Value)
        lstResultados.List(i, 3) = modFormateo.FormatearNumero(fila.Range.Cells(1, COL_INV_PRECIO_V).Value)
        i = i + 1
    Next fila
End Sub

' --- BUSCAR PRODUCTOS POR NOMBRE --------------------------------------------
' Filtra productos cuyo nombre contenga el texto buscado.

Public Sub BuscarProductos(ByVal textoBusqueda As String, _
                            ByRef lstResultados As MSForms.ListBox)
    Dim wsAlmacen As Worksheet
    Dim tblInv As ListObject
    Dim fila As ListRow
    Dim i As Long
    
    Set wsAlmacen = ThisWorkbook.Sheets(HOJA_ALMACEN)
    Set tblInv = wsAlmacen.ListObjects(TBL_INVENTARIO)
    
    lstResultados.Clear
    i = 0
    
    If textoBusqueda = "" Then
        ' Si no hay texto, cargar todos
        Call CargarTodosLosProductos(lstResultados)
        Exit Sub
    End If
    
    For Each fila In tblInv.ListRows
        Dim nombre As String
        nombre = fila.Range.Cells(1, COL_INV_NOMBRE).Value
        
        If UCase(nombre) Like "*" & UCase(textoBusqueda) & "*" Then
            lstResultados.AddItem
            lstResultados.List(i, 0) = fila.Range.Cells(1, COL_INV_CODIGO).Value
            lstResultados.List(i, 1) = nombre
            lstResultados.List(i, 2) = modFormateo.FormatearNumero(fila.Range.Cells(1, COL_INV_CANT_ACT).Value)
            lstResultados.List(i, 3) = modFormateo.FormatearNumero(fila.Range.Cells(1, COL_INV_PRECIO_V).Value)
            i = i + 1
        End If
    Next fila
End Sub

' --- BUSCAR FACTURA POR CÓDIGO ---------------------------------------------

Public Sub BuscarFacturaAvanzada()
    Dim codigoBuscar As String
    Dim celdaEncontrada As Range
    Dim HojaBD As Worksheet
    
    Set HojaBD = ThisWorkbook.Sheets(HOJA_FACTURAS)
    
    codigoBuscar = Trim(InputBox("Ingrese el código de factura:", "Buscar Factura"))
    If codigoBuscar = "" Then Exit Sub
    
    Application.ScreenUpdating = False
    Set celdaEncontrada = HojaBD.Columns("B").Find(codigoBuscar, LookIn:=xlValues, LookAt:=xlWhole)
    Application.ScreenUpdating = True
    
    If Not celdaEncontrada Is Nothing Then
        celdaEncontrada.Select
        Application.GoTo celdaEncontrada, True
    Else
        MsgBox "Factura no encontrada.", vbExclamation
    End If
End Sub

' --- VERIFICAR SI FACTURA YA EXISTE EN EL LIBRO ---------------------------

Public Function FacturaExisteEnLibro(facturaID As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim ultimaFila As Long
    Dim i As Long
    Dim idBuscado As String
    
    Set ws = ThisWorkbook.Sheets(HOJA_FACTURAS)
    idBuscado = Trim(CStr(facturaID))
    ultimaFila = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    
    For i = 5 To ultimaFila
        If Trim(CStr(ws.Cells(i, 2).Value)) = idBuscado Then
            FacturaExisteEnLibro = True
            Exit Function
        End If
    Next i
    
    FacturaExisteEnLibro = False
    Exit Function
    
ErrorHandler:
    FacturaExisteEnLibro = False
End Function

' ============================================================================
' FUNCIONES DE BÚSQUEDA POR RANGO DE FECHAS (para Busqueda*.frm)
' ============================================================================

' --- CARGAR DESCRIPCIONES ÚNICAS EN UN COMBOBOX ----------------------------

Public Sub CargarDescripcionesUnicas(ByRef tbl As ListObject, _
                                      ByVal columna As Long, _
                                      ByRef comboBox As MSForms.ComboBox)
    Dim dict As Object
    Dim fila As ListRow
    Dim valor As String
    
    If comboBox Is Nothing Then Exit Sub
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    comboBox.Clear
    comboBox.AddItem ""
    
    For Each fila In tbl.ListRows
        valor = Trim(CStr(fila.Range(columna).Value))
        If valor <> "" And Not dict.Exists(valor) Then
            dict.Add valor, valor
            comboBox.AddItem valor
        End If
    Next fila
    
    If comboBox.ListCount > 0 Then comboBox.ListIndex = 0
    Set dict = Nothing
End Sub

' --- BUSCAR EN TABLA DE GASTOS ----------------------------------------------
' Devuelve array 2D con: Fecha, Categoría, Descripción, Monto
' filasOriginales se llena con los números de fila reales.

Public Function BuscarGastos(ByVal fechaInicio As Date, _
                              ByVal fechaFin As Date, _
                              ByVal descripcion As String, _
                              ByRef filasOriginales As Collection) As Variant
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim fila As ListRow
    Dim resultados() As Variant
    Dim contador As Long
    Dim i As Long
    Dim cumple As Boolean
    
    Set ws = ThisWorkbook.Sheets(HOJA_GASTOS)
    Set tbl = ws.ListObjects(TBL_GASTOS)
    
    ' Primera pasada: contar
    contador = 0
    For Each fila In tbl.ListRows
        cumple = True
        If CDate(fila.Range(COL_GAS_FECHA).Value) < fechaInicio Then cumple = False
        If CDate(fila.Range(COL_GAS_FECHA).Value) > fechaFin Then cumple = False
        If descripcion <> "" Then
            If InStr(1, UCase(CStr(fila.Range(COL_GAS_DESCRIPCION).Value)), _
                     UCase(descripcion), vbTextCompare) = 0 Then cumple = False
        End If
        If cumple Then contador = contador + 1
    Next fila
    
    If contador = 0 Then
        BuscarGastos = Array()
        Exit Function
    End If
    
    ' Segunda pasada: llenar
    ReDim resultados(1 To contador, 1 To 4)
    i = 1
    For Each fila In tbl.ListRows
        cumple = True
        If CDate(fila.Range(COL_GAS_FECHA).Value) < fechaInicio Then cumple = False
        If CDate(fila.Range(COL_GAS_FECHA).Value) > fechaFin Then cumple = False
        If descripcion <> "" Then
            If InStr(1, UCase(CStr(fila.Range(COL_GAS_DESCRIPCION).Value)), _
                     UCase(descripcion), vbTextCompare) = 0 Then cumple = False
        End If
        If cumple Then
            resultados(i, 1) = Format(fila.Range(COL_GAS_FECHA).Value, "d-mmm-yyyy")
            resultados(i, 2) = fila.Range(COL_GAS_CATEGORIA).Value
            resultados(i, 3) = fila.Range(COL_GAS_DESCRIPCION).Value
            resultados(i, 4) = "CUP      " & modFormateo.FormatearNumero(fila.Range(COL_GAS_MONTO).Value)
            filasOriginales.Add fila.Range.Row, CStr(i)
            i = i + 1
        End If
    Next fila
    
    BuscarGastos = resultados
End Function

' --- BUSCAR EN TABLA DE VENTAS ----------------------------------------------
' Devuelve array 2D con: Descripción, Fecha, Col4, Monto, Col6, FacturaID

Public Function BuscarVentas(ByVal fechaInicio As Date, _
                              ByVal fechaFin As Date, _
                              ByVal descripcion As String, _
                              ByRef filasOriginales As Collection) As Variant
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim fila As ListRow
    Dim resultados() As Variant
    Dim contador As Long
    Dim i As Long
    Dim cumple As Boolean
    
    Set ws = ThisWorkbook.Sheets(HOJA_VENTAS)
    Set tbl = ws.ListObjects(TBL_VENTAS)
    
    contador = 0
    For Each fila In tbl.ListRows
        cumple = True
        If CDate(fila.Range(3).Value) < fechaInicio Then cumple = False
        If CDate(fila.Range(3).Value) > fechaFin Then cumple = False
        If descripcion <> "" Then
            If InStr(1, UCase(CStr(fila.Range(2).Value)), _
                     UCase(descripcion), vbTextCompare) = 0 Then cumple = False
        End If
        If cumple Then contador = contador + 1
    Next fila
    
    If contador = 0 Then
        BuscarVentas = Array()
        Exit Function
    End If
    
    ReDim resultados(1 To contador, 1 To 6)
    i = 1
    For Each fila In tbl.ListRows
        cumple = True
        If CDate(fila.Range(3).Value) < fechaInicio Then cumple = False
        If CDate(fila.Range(3).Value) > fechaFin Then cumple = False
        If descripcion <> "" Then
            If InStr(1, UCase(CStr(fila.Range(2).Value)), _
                     UCase(descripcion), vbTextCompare) = 0 Then cumple = False
        End If
        If cumple Then
            resultados(i, 1) = fila.Range(2).Value       ' Descripción
            resultados(i, 2) = Format(fila.Range(3).Value, "d-mmm-yyyy") ' Fecha
            resultados(i, 3) = fila.Range(4).Value       ' Col4
            resultados(i, 4) = "CUP      " & modFormateo.FormatearNumero(fila.Range(5).Value) ' Monto
            resultados(i, 5) = "CUP      " & modFormateo.FormatearNumero(fila.Range(6).Value) ' Col6
            resultados(i, 6) = fila.Range(7).Value       ' FacturaID
            filasOriginales.Add fila.Range.Row, CStr(i)
            i = i + 1
        End If
    Next fila
    
    BuscarVentas = resultados
End Function

' --- BUSCAR EN TABLA DE FACTURAS --------------------------------------------
' Devuelve array 2D con: FacturaID, Fecha, Total, Efectivo, Transferencia

Public Function BuscarFacturas(ByVal fechaInicio As Date, _
                                ByVal fechaFin As Date, _
                                ByVal descripcion As String, _
                                ByRef filasOriginales As Collection) As Variant
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim fila As ListRow
    Dim resultados() As Variant
    Dim contador As Long
    Dim i As Long
    Dim cumple As Boolean
    
    Set ws = ThisWorkbook.Sheets(HOJA_FACTURAS)
    Set tbl = ws.ListObjects(TBL_FACTURAS)
    
    contador = 0
    For Each fila In tbl.ListRows
        cumple = True
        If CDate(fila.Range(2).Value) < fechaInicio Then cumple = False
        If CDate(fila.Range(2).Value) > fechaFin Then cumple = False
        If descripcion <> "" Then
            If InStr(1, UCase(CStr(fila.Range(1).Value)), _
                     UCase(descripcion), vbTextCompare) = 0 Then cumple = False
        End If
        If cumple Then contador = contador + 1
    Next fila
    
    If contador = 0 Then
        BuscarFacturas = Array()
        Exit Function
    End If
    
    ReDim resultados(1 To contador, 1 To 5)
    i = 1
    For Each fila In tbl.ListRows
        cumple = True
        If CDate(fila.Range(2).Value) < fechaInicio Then cumple = False
        If CDate(fila.Range(2).Value) > fechaFin Then cumple = False
        If descripcion <> "" Then
            If InStr(1, UCase(CStr(fila.Range(1).Value)), _
                     UCase(descripcion), vbTextCompare) = 0 Then cumple = False
        End If
        If cumple Then
            resultados(i, 1) = fila.Range(1).Value       ' FacturaID
            resultados(i, 2) = Format(fila.Range(2).Value, "d-mmm-yyyy") ' Fecha
            resultados(i, 3) = "CUP      " & modFormateo.FormatearNumero(fila.Range(3).Value) ' Total
            resultados(i, 4) = "CUP      " & modFormateo.FormatearNumero(fila.Range(4).Value) ' Efectivo
            resultados(i, 5) = "CUP      " & modFormateo.FormatearNumero(fila.Range(5).Value) ' Transferencia
            filasOriginales.Add fila.Range.Row, CStr(i)
            i = i + 1
        End If
    Next fila
    
    BuscarFacturas = resultados
End Function

' --- NAVEGAR A FILA ORIGINAL DESDE UN RESULTADO -----------------------------

Public Sub NavegarAFilaOriginal(ByVal numFilaOriginal As Long, _
                                 ByVal nombreHoja As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(nombreHoja)
    ws.Activate
    Application.GoTo Reference:=ws.Cells(numFilaOriginal, 1), Scroll:=True
    ws.Rows(numFilaOriginal).Select
End Sub
