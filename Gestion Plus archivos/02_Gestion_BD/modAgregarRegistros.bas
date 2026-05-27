Attribute VB_Name = "modAgregarRegistros"
' ============================================================================
' MÓDULO: modAgregarRegistros
' PROPÓSITO: Consolidación de todas las funciones para agregar registros
' a las tablas. Reemplaza las versiones duplicadas en Módulo1, serverControl,
' facturacion, ingresar_form, Gasto_Form, Elaborar, etc.
' ============================================================================
' FUNCIONES INCLUIDAS:
'   - AgregarGasto          (antes Agregar_gasto, Agregar_gasto_inv)
'   - AgregarFecha          (antes Agregar_Fecha - 6 duplicados)
'   - AgregarProductoUnico  (antes Agregar_Producto - 4 duplicados)
'   - AgregarSalario        (antes Agregar_sal)
'   - InsertarFactura       (antes InsertarFacturaEnLibro)
'   - InsertarDetalleVenta  (antes InsertarDetalleVenta)
' ============================================================================

Option Explicit

' --- AGREGAR GASTO ---------------------------------------------------------
' Crea un nuevo registro en la tabla de gastos.
' Categorías típicas: "Fijo", "Variable", "Inversión", "Merma"

Public Sub AgregarGasto( _
    fecha As Date, _
    categoria As String, _
    descripcion As String, _
    monto As Double)

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim newrow As ListRow
    
    Set ws = ThisWorkbook.Sheets(HOJA_GASTOS)
    Set tbl = ws.ListObjects(TBL_GASTOS)
    
    ws.Unprotect password:=PASS_HOJA
    
    Set newrow = tbl.ListRows.Add
    With newrow
        .Range(COL_GAS_FECHA) = CDate(fecha)
        .Range(COL_GAS_CATEGORIA) = categoria
        .Range(COL_GAS_DESCRIPCION) = descripcion
        .Range(COL_GAS_MONTO) = CDbl(monto)
    End With
    
    tbl.Sort.Apply
    ws.Protect password:=PASS_HOJA, UserInterfaceOnly:=True
End Sub

' --- AGREGAR GASTO DE INVERSIÓN (atalajo) ---------------------------------

Public Sub AgregarGastoInversion(fecha As Date, descripcion As String, monto As Double)
    Call AgregarGasto(fecha, CAT_INVERSION, descripcion, monto)
End Sub

' --- AGREGAR FECHA ÚNICA ---------------------------------------------------
' Agrega una fecha a la tabla FechasUnicas si no existe ya.

Public Sub AgregarFecha(Optional fecha As Variant)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim fila As ListRow
    Dim fechaValor As Date
    
    If IsMissing(fecha) Then
        fechaValor = Date
    Else
        fechaValor = CDate(fecha)
    End If
    
    Set ws = ThisWorkbook.Sheets(HOJA_FECHAS)
    Set tbl = ws.ListObjects(TBL_FECHAS_UNICAS)
    
    ' Verificar si ya existe
    For Each fila In tbl.ListRows
        If CDate(fila.Range(1)) = fechaValor Then
            Exit Sub
        End If
    Next fila
    
    ws.Unprotect password:=PASS_HOJA
    
    Set fila = tbl.ListRows.Add
    fila.Range(1) = fechaValor
    tbl.Sort.Apply
    
    ws.Protect password:=PASS_HOJA, UserInterfaceOnly:=True
End Sub

' --- AGREGAR PRODUCTO ÚNICO ------------------------------------------------
' Agrega un producto a la tabla de productos únicos si no existe.

Public Sub AgregarProductoUnico(nombreProducto As String)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim fila As ListRow
    
    Set ws = ThisWorkbook.Sheets(HOJA_FECHAS)
    Set tbl = ws.ListObjects(TBL_PRODUCTOS_UNICOS)
    
    ' Verificar si ya existe (comparación sin distinguir mayúsculas)
    For Each fila In tbl.ListRows
        If UCase(Trim(fila.Range(1))) = UCase(Trim(nombreProducto)) Then
            Exit Sub
        End If
    Next fila
    
    ws.Unprotect password:=PASS_HOJA
    
    Set fila = tbl.ListRows.Add
    fila.Range(1) = nombreProducto
    tbl.Sort.Apply
    
    ws.Protect password:=PASS_HOJA, UserInterfaceOnly:=True
End Sub

' --- AGREGAR SALARIO A GASTOS ----------------------------------------------

Public Sub AgregarSalario(cargoTrabajador As String, montoPago As Double)
    Dim ws As Worksheet
    Dim tbl As ListObject
    
    Set ws = ThisWorkbook.Sheets(HOJA_GASTOS)
    Set tbl = ws.ListObjects(TBL_GASTOS)
    
    ws.Unprotect password:=PASS_HOJA
    
    With tbl.ListRows.Add
        .Range(COL_GAS_FECHA) = Date
        .Range(COL_GAS_CATEGORIA) = CAT_FIJO
        .Range(COL_GAS_DESCRIPCION) = "Salario " & cargoTrabajador
        .Range(COL_GAS_MONTO) = CDbl(montoPago)
    End With
    
    tbl.Sort.Apply
    ws.Protect password:=PASS_HOJA, UserInterfaceOnly:=True
End Sub

' --- INSERTAR FACTURA EN LIBRO ---------------------------------------------

Public Function InsertarFacturaEnLibro( _
    facturaID As String, _
    fecha As Date, _
    total As Double, _
    efectivo As Double, _
    transferencia As Double) As Boolean
    
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim ultimaFila As Long
    Dim filaInsercion As Long
    
    Set ws = ThisWorkbook.Sheets(HOJA_FACTURAS)
    
    ultimaFila = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    If ultimaFila < 5 Then
        filaInsercion = 5
    ElseIf ultimaFila = 5 And Trim(CStr(ws.Cells(5, 2).Value)) = "" Then
        filaInsercion = 5
    Else
        filaInsercion = ultimaFila + 1
    End If
    
    ws.Cells(filaInsercion, 2).Value = facturaID
    ws.Cells(filaInsercion, 3).Value = fecha
    ws.Cells(filaInsercion, 4).Value = total
    ws.Cells(filaInsercion, 5).Value = efectivo
    ws.Cells(filaInsercion, 6).Value = transferencia
    
    Call AgregarFecha(fecha)
    
    ws.ListObjects(TBL_FACTURAS).Sort.Apply
    InsertarFacturaEnLibro = True
    
    Exit Function
    
ErrorHandler:
    MsgBox "Error al insertar la factura en el libro." & vbCrLf & _
           "FacturaID: " & facturaID & vbCrLf & _
           "Detalle: " & Err.Description, vbExclamation
    InsertarFacturaEnLibro = False
End Function

' --- INSERTAR DETALLE DE VENTA ---------------------------------------------

Public Function InsertarDetalleVenta( _
    facturaID As String, _
    fecha As Date, _
    codigoProducto As String, _
    nombre As String, _
    cantidad As Double, _
    precio As Double, _
    subtotal As Double) As Boolean
    
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim ultimaFila As Long
    Dim filaInsercion As Long
    
    Set ws = ThisWorkbook.Sheets(HOJA_VENTAS)
    
    ultimaFila = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    If ultimaFila < 5 Then
        filaInsercion = 5
    ElseIf ultimaFila = 5 And Trim(CStr(ws.Cells(5, 2).Value)) = "" Then
        filaInsercion = 5
    Else
        filaInsercion = ultimaFila + 1
    End If
    
    ' Escribir datos en columnas B:H
    ws.Cells(filaInsercion, 2).Value = codigoProducto
    ws.Cells(filaInsercion, 3).Value = nombre
    ws.Cells(filaInsercion, 4).Value = fecha
    ws.Cells(filaInsercion, 5).Value = cantidad
    ws.Cells(filaInsercion, 6).Value = precio
    ws.Cells(filaInsercion, 7).Value = subtotal
    ws.Cells(filaInsercion, 8).Value = facturaID
    
    ' Actualizar stock y registrar inversión
    Call RestarCantidadInventario(codigoProducto, cantidad)
    
    Dim montoInv As Double
    montoInv = CalcularMontoInversion(codigoProducto, cantidad)
    Call AgregarGasto(fecha, CAT_INVERSION, nombre, montoInv)
    
    ws.ListObjects(TBL_VENTAS).Sort.Apply
    InsertarDetalleVenta = True
    
    Exit Function
    
ErrorHandler:
    MsgBox "Error al insertar el detalle de venta." & vbCrLf & _
           "FacturaID: " & facturaID & vbCrLf & _
           "Producto: " & codigoProducto & vbCrLf & _
           "Detalle: " & Err.Description, vbExclamation
    InsertarDetalleVenta = False
End Function
