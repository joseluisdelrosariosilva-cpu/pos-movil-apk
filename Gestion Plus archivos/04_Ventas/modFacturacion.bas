Attribute VB_Name = "modFacturacion"
' ============================================================================
' MÓDULO: modFacturacion
' PROPÓSITO: Cálculos de facturación y generación de ID de factura.
'            Versión desacoplada: recibe los controles como parámetros.
' ============================================================================

Option Explicit

' --- CALCULAR TOTAL DEL LISTBOX DE FACTURACIÓN -----------------------------
' Recibe el listbox de productos como parámetro (desacoplado del form).

Public Sub CalcularTotalFacturacion(ByRef lstProductos As MSForms.ListBox, _
                                     ByRef txtTotal As MSForms.TextBox)
    Dim total_de_linea As Double
    Dim suma_total As Double
    Dim i As Integer
    
    On Error Resume Next
    For i = 0 To lstProductos.ListCount - 1
        total_de_linea = CDbl(lstProductos.List(i, 2)) * _
                        CDbl(lstProductos.List(i, 3))
        suma_total = suma_total + total_de_linea
    Next i
    On Error GoTo 0
    
    txtTotal.Value = modFormateo.FormatearNumero(suma_total)
End Sub

' --- GENERAR ID DE FACTURA ------------------------------------------------
' Formato: [5 dígitos de fecha serial] + [5 dígitos consecutivos del día]
' El prefijo se reinicia diariamente.
' Recibe los TextBox como parámetros (desacoplado del form).

Public Sub GenerarIdFactura(ByRef txtFecha As MSForms.TextBox, _
                             ByRef txtFacturaId As MSForms.TextBox)
    Dim numero As Long
    Dim fecha As Date
    Dim identificador_fecha As Long
    Dim maximo As Long
    Dim theWs As Worksheet
    Dim theTbl As ListObject
    Dim fila As ListRow
    Dim idActual As String
    
    Set theWs = ThisWorkbook.Sheets(HOJA_VENTAS)
    Set theTbl = theWs.ListObjects(TBL_VENTAS)
    
    If txtFecha.Value = "" Or Not IsDate(txtFecha.Value) Then Exit Sub
    fecha = CDate(txtFecha.Value)
    identificador_fecha = CLng(fecha)
    
    ' Buscar el número consecutivo más alto del día
    maximo = 0
    For Each fila In theTbl.ListRows
        idActual = CStr(fila.Range(7).Value)
        If Left(idActual, 5) = CStr(identificador_fecha) Then
            numero = CLng(Right(idActual, 5))
            If numero > maximo Then maximo = numero
        End If
    Next fila
    
    maximo = maximo + 1
    
    txtFacturaId.Value = _
        CStr(identificador_fecha) & CStr(Format(maximo, "00000"))
End Sub

' --- CALCULAR PRECIO DE LOTE ----------------------------------------------

Public Function CalcularPrecioLote(ByVal precioUnitario As Double, _
                                    ByVal Cantidad As Double) As Double
    CalcularPrecioLote = precioUnitario * Cantidad
End Function

' --- REGISTRAR VENTA COMPLETA (productos + factura + actualizar stock) -------
' Recibe los datos de la venta y escribe en las tablas HISTORIALVENTAS y
' HISTORIALVENTAS4. Actualiza stock y registra gasto de inversión.
' Parámetros:
'   arrProductos - Array 2D (i, 0-3): codigo, nombre, precio, cantidad
'   facturaID, fecha     - Datos de la factura
'   totalAPagar           - Total a pagar
'   ingresoCliente        - Dinero recibido del cliente
'   devolucion            - Vuelto a devolver
'   cantidadTransferencia - Monto de la transferencia
'   esEfectivo            - True si pagó en efectivo
'   esTransferencia       - True si pagó con transferencia

Public Sub RegistrarVentaCompleta(ByRef arrProductos() As Variant, _
                                   ByVal facturaID As String, _
                                   ByVal fecha As Date, _
                                   ByVal totalAPagar As Double, _
                                   ByVal ingresoCliente As Double, _
                                   ByVal devolucion As Double, _
                                   ByVal cantidadTransferencia As Double, _
                                   ByVal esEfectivo As Boolean, _
                                   ByVal esTransferencia As Boolean)
    On Error GoTo ErrorHandler
    
    Dim wsVentas As Worksheet
    Dim wsFacturas As Worksheet
    Dim tblventas As ListObject
    Dim tblFacturas As ListObject
    Dim ultimaFilaVentas As Long
    Dim ultimaFilaFacturas As Long
    Dim filaDestino As Long
    Dim i As Long
    Dim totalElementos As Long
    
    Set wsVentas = ThisWorkbook.Sheets(HOJA_VENTAS)
    Set wsFacturas = ThisWorkbook.Sheets(HOJA_FACTURAS)
    Set tblventas = wsVentas.ListObjects(TBL_VENTAS)
    Set tblFacturas = wsFacturas.ListObjects(TBL_FACTURAS)
    
    totalElementos = UBound(arrProductos, 1) - LBound(arrProductos, 1) + 1
    
    ' Determinar fila de inicio en Hoja8
    ultimaFilaVentas = wsVentas.Range("B" & wsVentas.Rows.Count).End(xlUp).Row
    If ultimaFilaVentas = 5 And Trim(CStr(wsVentas.Cells(5, 2).Value)) = "" Then
        filaDestino = 5
    Else
        filaDestino = ultimaFilaVentas + 1
    End If
    
    ' --- ESCRIBIR PRODUCTOS EN HISTORIALVENTAS ---
    For i = LBound(arrProductos, 1) To UBound(arrProductos, 1)
        Dim codigo As String
        Dim nombre As String
        Dim precio As Double
        Dim Cantidad As Double
        Dim subtotal As Double
        
        codigo = CStr(arrProductos(i, 0))
        nombre = CStr(arrProductos(i, 1))
        precio = CDbl(arrProductos(i, 2))
        Cantidad = CDbl(arrProductos(i, 3))
        subtotal = precio * Cantidad
        
        wsVentas.Cells(filaDestino, 2).Value = codigo
        wsVentas.Cells(filaDestino, 3).Value = nombre
        wsVentas.Cells(filaDestino, 4).Value = CDate(fecha)
        wsVentas.Cells(filaDestino, 5).Value = Cantidad
        wsVentas.Cells(filaDestino, 6).Value = precio
        wsVentas.Cells(filaDestino, 7).Value = subtotal
        wsVentas.Cells(filaDestino, 8).Value = facturaID
        
        ' Actualizar stock y registrar inversión
        Call modStock.RestarCantidadInventario(codigo, Cantidad)
        Call modAgregarRegistros.AgregarGastoInversion(fecha, nombre, _
            modCalculosInventario.CalcularMontoInversion(codigo, Cantidad))
        
        filaDestino = filaDestino + 1
    Next i
    
    ' --- ESCRIBIR FACTURA EN HISTORIALVENTAS4 ---
    ultimaFilaFacturas = wsFacturas.Range("B" & wsFacturas.Rows.Count).End(xlUp).Row
    Dim filaFactura As Long
    
    If ultimaFilaFacturas = 5 And Trim(CStr(wsFacturas.Cells(5, 2).Value)) = "" Then
        filaFactura = 5
    Else
        filaFactura = ultimaFilaFacturas + 1
    End If
    
    wsFacturas.Cells(filaFactura, 2).Value = facturaID
    wsFacturas.Cells(filaFactura, 3).Value = CDate(fecha)
    wsFacturas.Cells(filaFactura, 4).Value = totalAPagar
    
    ' Calcular efectivo neto (ingreso - devolución)
    If esEfectivo And esTransferencia Then
        ' Mixto
        wsFacturas.Cells(filaFactura, 5).Value = ingresoCliente
        wsFacturas.Cells(filaFactura, 6).Value = cantidadTransferencia
    ElseIf esEfectivo Then
        ' Solo efectivo
        wsFacturas.Cells(filaFactura, 5).Value = ingresoCliente - devolucion
        wsFacturas.Cells(filaFactura, 6).Value = 0
    ElseIf esTransferencia Then
        ' Solo transferencia
        wsFacturas.Cells(filaFactura, 5).Value = 0
        wsFacturas.Cells(filaFactura, 6).Value = cantidadTransferencia
    End If
    
    ' Ordenar y finalizar
    tblventas.Sort.Apply
    tblFacturas.Sort.Apply
    Call modAgregarRegistros.AgregarFecha(fecha)
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error al registrar la venta." & vbCrLf & _
           "Factura: " & facturaID & vbCrLf & _
           "Detalle: " & Err.Description, vbExclamation
End Sub

