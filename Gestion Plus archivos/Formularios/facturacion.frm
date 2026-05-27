VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} facturacion 
   Caption         =   "Orden de venta"
   ClientHeight    =   9580.001
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   13620
   OleObjectBlob   =   "facturacion.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "facturacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

' ============================================================================
' FORMULARIO: facturacion
' PROPÓSITO:  Interfaz de ventas con búsqueda de productos, carrito,
'              procesamiento de pago (efectivo/transferencia). Toda la
'              lógica de negocio delega en módulos reorganizados.
' ============================================================================

' --- AGREGAR PRODUCTO AL CARRITO --------------------------------------------
Private Sub boton_add_covered_Click()
    On Error Resume Next
    
    ' Validar campos obligatorios
    If Me.codigo_facturacion_TextBox.Value = "" Then
        MsgBox "Añada algún producto desde la lista de búsqueda para continuar", vbExclamation
        Exit Sub
    ElseIf Me.cantidad_facturacion_TextBox.Value = "" Then
        MsgBox "Indique la cantidad de unidades que desea agregar", vbExclamation
        Exit Sub
    ElseIf Me.Disponibilidad_facturacio_TextBox.Value <= 0 Then
        MsgBox "No hay disponibilidad en almacén", vbCritical
        Exit Sub
    ElseIf CDbl(Me.cantidad_facturacion_TextBox.Value) <= 0 Then
        MsgBox "Debe añadir una cantidad superior a 0", vbExclamation
        Exit Sub
    End If
    
    ' Validar repetición
    If modBusquedas.ComprobarRepeticionFacturacion( _
        Me.codigo_facturacion_TextBox.Value, Me.productos_facturacion_ListBox) Then
        MsgBox "No puede añadir un producto añadido anteriormente.", vbExclamation
        Exit Sub
    End If
    
    ' Validar cantidad vs disponibilidad
    If CDbl(Me.cantidad_facturacion_TextBox.Value) > CDbl(Me.Disponibilidad_facturacio_TextBox.Value) Then
        MsgBox "La disponibilidad en almacén es: " & Me.Disponibilidad_facturacio_TextBox.Value, vbExclamation
        Me.cantidad_facturacion_TextBox.Value = Me.Disponibilidad_facturacio_TextBox.Value
        Me.SpinButton1.Value = CDbl(Me.Disponibilidad_facturacio_TextBox.Value)
        Exit Sub
    End If
    
    ' Agregar al listbox
    Me.productos_facturacion_ListBox.AddItem
    With Me.productos_facturacion_ListBox
        .List(.ListCount - 1, 0) = Me.codigo_facturacion_TextBox.Value
        .List(.ListCount - 1, 1) = Me.nombre_de_producto_facturacionTextBox.Value
        .List(.ListCount - 1, 2) = Me.Precio_facturación_TextBox.Value
        .List(.ListCount - 1, 3) = Me.cantidad_facturacion_TextBox.Value
    End With
    
    ' Limpiar campos y recalcular
    Me.codigo_facturacion_TextBox.Value = ""
    Me.nombre_de_producto_facturacionTextBox.Value = ""
    Me.Precio_facturación_TextBox.Value = ""
    Me.Disponibilidad_facturacio_TextBox.Value = ""
    Me.cantidad_facturacion_TextBox.Value = ""
    
    Call modFacturacion.CalcularTotalFacturacion( _
        Me.productos_facturacion_ListBox, Me.total_facturacion_TextBox)
End Sub

' --- ELIMINAR PRODUCTO DEL CARRITO ------------------------------------------
Private Sub boton_delete_covered_Click()
    If Me.productos_facturacion_ListBox.ListIndex = -1 Then
        MsgBox "Seleccione el producto que desea quitar de la facturación", vbExclamation
    Else
        Dim respuesta
        respuesta = MsgBox("¿Está seguro que desea quitar el producto seleccionado?", _
                           vbYesNo, "Eliminar producto")
        If respuesta = 6 Then
            Me.productos_facturacion_ListBox.RemoveItem (Me.productos_facturacion_ListBox.ListIndex)
            Call modFacturacion.CalcularTotalFacturacion( _
                Me.productos_facturacion_ListBox, Me.total_facturacion_TextBox)
        End If
    End If
End Sub

' --- EDITAR CANTIDAD DE PRODUCTO --------------------------------------------
Private Sub boton_edit_covered_Click()
    On Error GoTo Salir
    If Me.productos_facturacion_ListBox.ListIndex = -1 Then
        MsgBox "Seleccione el producto al que desea modificar su cantidad", vbExclamation
    Else
        Dim nueva_cantidad As Double
        nueva_cantidad = CDbl(InputBox("Inserte la nueva cantidad", "Editar Cantidad"))
        Me.productos_facturacion_ListBox.List(Me.productos_facturacion_ListBox.ListIndex, 3) = _
            modFormateo.FormatearNumero(nueva_cantidad)
        Call modFacturacion.CalcularTotalFacturacion( _
            Me.productos_facturacion_ListBox, Me.total_facturacion_TextBox)
    End If
Salir:
    Exit Sub
End Sub

' --- PROCESAR PAGO (pasar a pantalla de pago) --------------------------------
Private Sub boton_procesar_pago_covered_Click()
    If Me.productos_facturacion_ListBox.ListCount = 0 Then
        MsgBox "Añada productos a la factura para continuar", vbExclamation
        Exit Sub
    End If
    
    Dim total_fact As Double
    total_fact = CDbl(Me.total_facturacion_TextBox.Value)
    
    Me.total_a_pagar_TextBox.Value = modFormateo.FormatearNumero(total_fact)
    Me.restan_por_pagar_TextBox.Value = modFormateo.FormatearNumero(-1 * total_fact)
    
    ' Deshabilitar controles de selección de productos
    Me.factura_id_TextBox.Enabled = False
    Me.fecha_de_orden_TextBox.Enabled = False
    Me.boton_calendario_reposo.Visible = False
    Me.nombre_producto_busqueda_TextBox.Enabled = False
    Me.productos_busqueda_ListBox.Enabled = False
    Me.codigo_facturacion_TextBox.Enabled = False
    Me.nombre_de_producto_facturacionTextBox.Enabled = False
    Me.Precio_facturación_TextBox.Enabled = False
    Me.Disponibilidad_facturacio_TextBox.Enabled = False
    Me.cantidad_facturacion_TextBox.Enabled = False
    Me.SpinButton1.Enabled = False
    Me.productos_facturacion_ListBox.Enabled = False
    Me.total_facturacion_TextBox.Enabled = False
    
    ' Habilitar controles de pago
    Me.total_a_pagar_TextBox.Enabled = True
    Me.restan_por_pagar_TextBox.Enabled = True
    Me.transferecia_CheckBox.Enabled = True
    Me.efectivo_CheckBox.Enabled = True
    Me.MultiPage1.Value = 1
End Sub

' --- REGISTRAR VENTA (guardar en tablas) -------------------------------------
Private Sub boton_registrar_ventas_covered_Click()
    On Error GoTo ErrorHandler
    
    If CDbl(Me.restan_por_pagar_TextBox.Value) <> 0 Then
        MsgBox "Por favor complete el proceso de pago para finalizar", vbExclamation
        Exit Sub
    End If
    
    ' Extraer productos del listbox a un array
    Dim totalElementos As Long
    totalElementos = Me.productos_facturacion_ListBox.ListCount
    If totalElementos = 0 Then Exit Sub
    
    Dim arrProductos() As Variant
    ReDim arrProductos(0 To totalElementos - 1, 0 To 3)
    
    Dim i As Long
    For i = 0 To totalElementos - 1
        arrProductos(i, 0) = Me.productos_facturacion_ListBox.List(i, 0)
        arrProductos(i, 1) = Me.productos_facturacion_ListBox.List(i, 1)
        arrProductos(i, 2) = CDbl(Me.productos_facturacion_ListBox.List(i, 2))
        arrProductos(i, 3) = CDbl(Me.productos_facturacion_ListBox.List(i, 3))
    Next i
    
    ' Preparar datos de pago
    Dim devolucionValor As Double
    If Me.devolucion_TextBox.Value = "" Then devolucionValor = 0 Else devolucionValor = CDbl(Me.devolucion_TextBox.Value)
    
    Dim ingresoCliente As Double
    If Me.ingreso_cliente_pago_TextBox.Value = "" Then ingresoCliente = 0 Else ingresoCliente = CDbl(Me.ingreso_cliente_pago_TextBox.Value)
    
    Dim cantidadTransf As Double
    If Me.cantidad_a_pagar_transeferencia_TextBox.Value = "" Then cantidadTransf = 0 Else cantidadTransf = CDbl(Me.cantidad_a_pagar_transeferencia_TextBox.Value)
    
    ' Registrar venta completa
    Call modFacturacion.RegistrarVentaCompleta( _
        arrProductos, _
        Me.factura_id_TextBox.Value, _
        CDate(Me.fecha_de_orden_TextBox.Value), _
        CDbl(Me.total_a_pagar_TextBox.Value), _
        ingresoCliente, _
        devolucionValor, _
        cantidadTransf, _
        CBool(Me.efectivo_CheckBox.Value), _
        CBool(Me.transferecia_CheckBox.Value))
    
    Call reinicio_POS
    Exit Sub
    
ErrorHandler:
    MsgBox "Error al registrar la venta: " & Err.Description, vbCritical
End Sub

' --- REINICIO DE POS (volver a estado inicial tras venta exitosa) ------------
Private Sub reinicio_POS()
    MsgBox "Venta realizada con éxito", vbInformation, "Orden de venta"
    
    Me.MultiPage1.Value = 0
    Me.productos_facturacion_ListBox.Clear
    Me.total_facturacion_TextBox.Value = ""
    Me.total_a_pagar_TextBox.Value = ""
    Me.restan_por_pagar_TextBox.Value = ""
    Me.devolucion_TextBox.Value = ""
    Me.ingreso_cliente_pago_TextBox.Value = ""
    Me.cantidad_a_pagar_transeferencia_TextBox.Value = ""
    Me.efectivo_CheckBox.Value = False
    Me.transferecia_CheckBox.Value = False
    Me.devoluacion_efectuada_CheckBox.Value = False
    Me.fecha_de_orden_TextBox.Value = Format(Date, "d-mmm-yyyy")
    
    ' Re-habilitar controles de selección de productos
    Me.factura_id_TextBox.Enabled = True
    Me.fecha_de_orden_TextBox.Enabled = True
    Me.nombre_producto_busqueda_TextBox.Enabled = True
    Me.productos_busqueda_ListBox.Enabled = True
    Me.codigo_facturacion_TextBox.Enabled = True
    Me.nombre_de_producto_facturacionTextBox.Enabled = True
    Me.Precio_facturación_TextBox.Enabled = True
    Me.Disponibilidad_facturacio_TextBox.Enabled = True
    Me.cantidad_facturacion_TextBox.Enabled = True
    Me.SpinButton1.Enabled = True
    Me.productos_facturacion_ListBox.Enabled = True
    Me.total_facturacion_TextBox.Enabled = True
    
    ' Deshabilitar controles de pago
    Me.total_a_pagar_TextBox.Enabled = False
    Me.restan_por_pagar_TextBox.Enabled = False
    Me.transferecia_CheckBox.Enabled = False
    Me.efectivo_CheckBox.Enabled = False
    Me.ingreso_cliente_pago_TextBox.Enabled = True
    Me.boton_calendario_reposo.Visible = True
    
    ' Recargar lista de productos (por si cambió disponibilidad)
    Call modBusquedas.CargarTodosLosProductos(Me.productos_busqueda_ListBox)
    Call modFacturacion.GenerarIdFactura( _
        Me.fecha_de_orden_TextBox, Me.factura_id_TextBox)
End Sub

' --- BOTÓN ATRÁS (volver a selección de productos) ---------------------------
Private Sub boton_atras_covered_Click()
    Me.MultiPage1.Value = 0
    
    ' Re-habilitar controles de selección
    Me.factura_id_TextBox.Enabled = True
    Me.fecha_de_orden_TextBox.Enabled = True
    Me.nombre_producto_busqueda_TextBox.Enabled = True
    Me.productos_busqueda_ListBox.Enabled = True
    Me.codigo_facturacion_TextBox.Enabled = True
    Me.nombre_de_producto_facturacionTextBox.Enabled = True
    Me.Precio_facturación_TextBox.Enabled = True
    Me.Disponibilidad_facturacio_TextBox.Enabled = True
    Me.cantidad_facturacion_TextBox.Enabled = True
    Me.SpinButton1.Enabled = True
    Me.productos_facturacion_ListBox.Enabled = True
    Me.total_facturacion_TextBox.Enabled = True
    
    ' Restablecer controles de pago
    Me.total_a_pagar_TextBox.Enabled = False
    Me.restan_por_pagar_TextBox.Enabled = False
    Me.total_a_pagar_TextBox.Value = ""
    Me.restan_por_pagar_TextBox.Value = ""
    Me.devolucion_TextBox.Value = ""
    Me.ingreso_cliente_pago_TextBox.Value = ""
    Me.efectivo_CheckBox.Value = False
    Me.efectivo_CheckBox.Enabled = False
    Me.transferecia_CheckBox.Value = False
    Me.transferecia_CheckBox.Enabled = False
    Me.ingreso_cliente_pago_TextBox.Enabled = True
    Me.devoluacion_efectuada_CheckBox.Enabled = True
    Me.devoluacion_efectuada_CheckBox.Value = False
    Me.transferencia_recibida_CheckBox.Enabled = True
    Me.transferencia_recibida_CheckBox.Value = False
    Me.cantidad_a_pagar_transeferencia_TextBox.Enabled = True
    Me.boton_calendario_reposo.Visible = True
End Sub

' --- BUSCAR PRODUCTOS (filtro en tiempo real) --------------------------------
Private Sub nombre_producto_busqueda_TextBox_Change()
    Call modBusquedas.BuscarProductos( _
        Me.nombre_producto_busqueda_TextBox.Value, _
        Me.productos_busqueda_ListBox)
End Sub

' --- SELECCIONAR PRODUCTO DE LA BÚSQUEDA (doble click) -----------------------
Private Sub productos_busqueda_ListBox_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    On Error Resume Next
    With Me.productos_busqueda_ListBox
        Me.codigo_facturacion_TextBox.Value = .List(.ListIndex, 0)
        Me.nombre_de_producto_facturacionTextBox.Value = .List(.ListIndex, 1)
        Me.Disponibilidad_facturacio_TextBox.Value = .List(.ListIndex, 2)
        Me.Precio_facturación_TextBox.Value = .List(.ListIndex, 3)
    End With
End Sub

' --- EVENTOS DE PAGO: EFECTIVO -----------------------------------------------
Private Sub efectivo_CheckBox_Click()
    Dim visible As Boolean
    visible = CBool(Me.efectivo_CheckBox.Value)
    
    Me.ingreso_cliente_pago_TextBox.Visible = visible
    Me.devolucion_TextBox.Visible = visible
    Me.devoluacion_efectuada_CheckBox.Visible = visible
    Me.Label11.Visible = visible
    Me.Label12.Visible = visible
End Sub

' --- EVENTOS DE PAGO: TRANSFERENCIA ------------------------------------------
Private Sub transferecia_CheckBox_Click()
    Dim visible As Boolean
    visible = CBool(Me.transferecia_CheckBox.Value)
    
    Me.Label24.Visible = visible
    Me.cantidad_a_pagar_transeferencia_TextBox.Visible = visible
    Me.transferencia_recibida_CheckBox.Visible = visible
End Sub

' --- CÁLCULO DE DEVOLUCIÓN E INGRESO DEL CLIENTE -----------------------------
Private Sub ingreso_cliente_pago_TextBox_Change()
    Dim devolucionCalc As Double
    
    If Me.total_a_pagar_TextBox <> "" And Me.ingreso_cliente_pago_TextBox.Value <> "" Then
        Me.restan_por_pagar_TextBox.Value = modFormateo.FormatearNumero( _
            -1 * CDbl(Me.total_a_pagar_TextBox.Value) + CDbl(Me.ingreso_cliente_pago_TextBox.Value))
        devolucionCalc = CDbl(Me.ingreso_cliente_pago_TextBox) - CDbl(Me.total_a_pagar_TextBox)
    Else
        Me.ingreso_cliente_pago_TextBox.Value = 0#
    End If
    
    If devolucionCalc > 0 Then
        Me.devolucion_TextBox.Value = modFormateo.FormatearNumero(devolucionCalc)
    Else
        Me.devolucion_TextBox.Value = ""
    End If
End Sub

' --- CONFIRMAR DEVOLUCIÓN EFECTUADA ------------------------------------------
Private Sub devoluacion_efectuada_CheckBox_Click()
    If Me.devoluacion_efectuada_CheckBox.Value = True And Me.devolucion_TextBox <> "" Then
        Me.restan_por_pagar_TextBox.Value = modFormateo.FormatearNumero( _
            CDbl(Me.restan_por_pagar_TextBox.Value) - CDbl(Me.devolucion_TextBox.Value))
        Me.devoluacion_efectuada_CheckBox.Enabled = False
        Me.transferecia_CheckBox.Value = False
        Me.transferecia_CheckBox.Enabled = False
    ElseIf Me.devoluacion_efectuada_CheckBox.Value = True And Me.devolucion_TextBox = "" Then
        MsgBox "No se ha ingresado suficiente dinero para ejecutar una devolución", vbCritical
        Me.devoluacion_efectuada_CheckBox.Value = False
    End If
End Sub

' --- CONFIRMAR TRANSFERENCIA RECIBIDA ----------------------------------------
Private Sub transferencia_recibida_CheckBox_Click()
    If Me.cantidad_a_pagar_transeferencia_TextBox = "" Then
        MsgBox "No hay cantidad asignada a transferencia.", vbExclamation
        Me.transferencia_recibida_CheckBox.Value = False
        Exit Sub
    End If
    
    If Me.transferencia_recibida_CheckBox.Value = True Then
        Me.transferencia_recibida_CheckBox.Enabled = False
        Dim temporal As Double
        temporal = CDbl(Me.cantidad_a_pagar_transeferencia_TextBox.Value)
        Me.restan_por_pagar_TextBox.Value = modFormateo.FormatearNumero( _
            CDbl(Me.restan_por_pagar_TextBox) + CDbl(Me.cantidad_a_pagar_transeferencia_TextBox))
        Me.cantidad_a_pagar_transeferencia_TextBox.Value = modFormateo.FormatearNumero(temporal)
    End If
    
    If Me.ingreso_cliente_pago_TextBox <> 0 And Me.MultiPage1.Value = 1 Then
        Me.ingreso_cliente_pago_TextBox.Enabled = False
        Me.devoluacion_efectuada_CheckBox.Enabled = False
        Me.cantidad_a_pagar_transeferencia_TextBox.Enabled = False
    Else
        Me.efectivo_CheckBox.Value = False
        Me.efectivo_CheckBox.Enabled = False
    End If
End Sub

' --- RESTA POR PAGAR: estilo visual y cálculo de transferencia ---------------
Private Sub restan_por_pagar_TextBox_Change()
    If IsNumeric(Me.restan_por_pagar_TextBox.Value) Then
        If CDbl(Me.restan_por_pagar_TextBox.Value) >= 0 Then
            Me.restan_por_pagar_TextBox.ForeColor = vbGreen
            Me.restan_por_pagar_TextBox.BorderColor = vbGreen
        Else
            Me.restan_por_pagar_TextBox.ForeColor = vbRed
            Me.restan_por_pagar_TextBox.BorderColor = vbRed
        End If
    End If
    
    ' Calcular monto pendiente para transferencia
    If IsNumeric(Me.restan_por_pagar_TextBox.Value) Then
        If CDbl(Me.restan_por_pagar_TextBox.Value) < 0 And Me.ingreso_cliente_pago_TextBox.Value <> "" Then
            Me.cantidad_a_pagar_transeferencia_TextBox.Value = modFormateo.FormatearNumero( _
                CDbl(Me.total_a_pagar_TextBox) - CDbl(Me.ingreso_cliente_pago_TextBox))
        ElseIf CDbl(Me.restan_por_pagar_TextBox.Value) < 0 And Me.ingreso_cliente_pago_TextBox.Value = "" Then
            Me.cantidad_a_pagar_transeferencia_TextBox.Value = modFormateo.FormatearNumero( _
                CDbl(Me.total_a_pagar_TextBox))
        Else
            Me.cantidad_a_pagar_transeferencia_TextBox.Value = ""
        End If
    End If
End Sub

' --- SPINBUTTON PARA CANTIDAD ------------------------------------------------
Private Sub SpinButton1_Change()
    Me.cantidad_facturacion_TextBox = Me.SpinButton1.Value
End Sub

' --- CALENDARIO ---------------------------------------------------------------
Private Sub boton_calendario_covered_Click()
    seleccion_destino_calendario = 2
    frmCalendario.Show
End Sub

' --- BOTÓN ADMIN -------------------------------------------------------------
Private Sub boton_admin_covered_Click()
    Login.Show
End Sub

' --- GENERAR ID DE FACTURA AL CAMBIAR FECHA ----------------------------------
Private Sub fecha_de_orden_TextBox_Change()
    Call modFacturacion.GenerarIdFactura( _
        Me.fecha_de_orden_TextBox, Me.factura_id_TextBox)
End Sub

' --- INICIALIZACIÓN DEL FORMULARIO -------------------------------------------
Private Sub UserForm_Initialize()
    ' Cargar lista de productos
    With Me.productos_busqueda_ListBox
        .ColumnCount = 4
        .ColumnWidths = "60;85;55;30"
    End With
    Call modBusquedas.CargarTodosLosProductos(Me.productos_busqueda_ListBox)
    
    ' Configurar columnas del carrito
    Me.productos_facturacion_ListBox.ColumnWidths = "55;75;70;25"
    
    ' Estado inicial de controles de pago
    Me.total_a_pagar_TextBox.Enabled = False
    Me.restan_por_pagar_TextBox.Enabled = False
    Me.transferecia_CheckBox.Enabled = False
    Me.efectivo_CheckBox.Enabled = False
End Sub

Private Sub UserForm_Activate()
    Me.MultiPage1.Value = 0
    Me.fecha_de_orden_TextBox.Value = Format(Date, "d-mmm-yyyy")
    
    If Application.Visible = True Then
        Me.boton_admin_reposo.Visible = False
    Else
        Me.boton_admin_reposo.Visible = True
    End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu And Me.boton_admin_reposo.Visible = True Then
        ThisWorkbook.Save
        Application.Quit
    End If
End Sub

' --- VALIDACIÓN NUMÉRICA (KeyPress) ------------------------------------------
Private Sub cantidad_facturacion_TextBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Call modValidacionNumerica.SoloNumeros(KeyAscii)
End Sub
Private Sub ingreso_cliente_pago_TextBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Call modValidacionNumerica.SoloNumeros(KeyAscii)
End Sub

' --- EVENTOS MOUSE MOVE (manejo de ZOrder de botones) ------------------------
Private Sub boton_add_reposo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.boton_add_covered.Visible = True
    Me.boton_add_covered.ZOrder msoBringToFront
End Sub

Private Sub boton_admin_reposo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.boton_admin_covered.Visible = True
    Me.boton_admin_covered.ZOrder msoBringToFront
End Sub

Private Sub boton_atras_reposo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.boton_atras_covered.Visible = True
    Me.boton_atras_covered.ZOrder msoBringToFront
End Sub

Private Sub boton_calendario_reposo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.boton_calendario_covered.Visible = True
    Me.boton_calendario_covered.ZOrder msoBringToFront
End Sub

Private Sub boton_delete_reposo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.boton_delete_covered.Visible = True
    Me.boton_delete_covered.ZOrder msoBringToFront
End Sub

Private Sub boton_edit_reposo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.boton_edit_covered.Visible = True
    Me.boton_edit_covered.ZOrder msoBringToFront
End Sub

Private Sub boton_procesar_pago_reposo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.boton_procesar_pago_covered.Visible = True
    Me.boton_procesar_pago_covered.ZOrder msoBringToFront
End Sub

Private Sub boton_registrar_ventas_reposo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.boton_registrar_ventas_covered.ZOrder msoBringToFront
    Me.Label26.ZOrder msoBringToFront
End Sub

Private Sub Label26_Click()
    boton_registrar_ventas_covered_Click
End Sub

Private Sub MultiPage1_MouseMove(ByVal Index As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.boton_add_covered.Visible = False
    Me.boton_add_reposo.ZOrder msoBringToFront
    Me.boton_edit_covered.Visible = False
    Me.boton_edit_reposo.ZOrder msoBringToFront
    Me.boton_delete_covered.Visible = False
    Me.boton_delete_reposo.ZOrder msoBringToFront
    Me.boton_procesar_pago_covered.Visible = False
    Me.boton_procesar_pago_reposo.ZOrder msoBringToFront
    Me.boton_atras_covered.Visible = False
    Me.boton_atras_reposo.ZOrder msoBringToFront
    Me.boton_registrar_ventas_reposo.ZOrder msoBringToFront
    Me.Label26.ZOrder msoBringToFront
End Sub

Private Sub productos_facturacion_ListBox_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.boton_add_covered.Visible = False
    Me.boton_add_reposo.ZOrder msoBringToFront
    Me.boton_edit_covered.Visible = False
    Me.boton_edit_reposo.ZOrder msoBringToFront
    Me.boton_delete_covered.Visible = False
    Me.boton_delete_reposo.ZOrder msoBringToFront
End Sub

Private Sub total_facturacion_TextBox_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.boton_procesar_pago_covered.Visible = False
    Me.boton_procesar_pago_reposo.ZOrder msoBringToFront
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.boton_admin_reposo.ZOrder msoBringToFront
    Me.boton_admin_covered.Visible = False
    Me.boton_calendario_reposo.ZOrder msoBringToFront
    Me.boton_calendario_covered.Visible = False
End Sub

Private Sub fecha_de_orden_TextBox_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.boton_calendario_reposo.ZOrder msoBringToFront
    Me.boton_calendario_covered.Visible = False
End Sub
