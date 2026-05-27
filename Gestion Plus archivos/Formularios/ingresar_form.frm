VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ingresar_form 
   Caption         =   "Ingresar nuevo producto"
   ClientHeight    =   5930
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   14940
   OleObjectBlob   =   "ingresar_form.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "ingresar_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

' ============================================================================
' FORMULARIO: ingresar_form
' PROPÓSITO:  Interfaz para insertar o modificar productos en inventario.
'              Toda la lógica de negocio delega en módulos reorganizados.
' ============================================================================

Private Sub boton_calendario_covered_Click()
    seleccion_destino_calendario = 1
    frmCalendario.Show
End Sub

Private Sub boton_calendario_reposo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.boton_calendario_covered.ZOrder msoBringToFront
End Sub

' --- INSERTAR NUEVO PRODUCTO ------------------------------------------------
Private Sub boton_ingresar_covered_Click()
    Call modValidacionCampos.ValidarVaciosProducto(Me)
    
    If contador > 0 Then
        MsgBox "Rellene los campos vacíos", vbExclamation
        Exit Sub
    End If
    
    Dim ultima_fila As Integer
    Dim HojaBD As Worksheet
    Set HojaBD = ThisWorkbook.Sheets(HOJA_ALMACEN)
    
    ultima_fila = HojaBD.Range("B" & Rows.Count).End(xlUp).Row
    
    If ultima_fila = 5 And HojaBD.Range("B5").Value = "" Then
        HojaBD.Range("B5").Value = Me.codigo_TextBox.Value
        HojaBD.Range("C5").Value = Me.nombre_de_producto_TextBox.Value
        HojaBD.Range("D5").Value = CDate(Me.fecha_de_entrada_TextBox.Value)
        HojaBD.Range("E5").Value = Format(Round(CDbl(Me.cantidad_TextBox.Value), 3), "General Number")
        HojaBD.Range("F5").Value = Format(Round(CDbl(Me.cantidad_TextBox.Value), 3), "General Number")
        HojaBD.Range("G5").Value = Me.precio_venta_TextBox.Value
        HojaBD.Range("H5").Value = Me.precio_costo_TextBox.Value
        HojaBD.Range("I5").Value = CDbl(Me.precio_costo_TextBox.Value) * CDbl(Me.cantidad_TextBox.Value)
        HojaBD.Range("J5").Value = Me.Ingreso_TextBox.Value
        HojaBD.Range("K5").Value = Me.Ganancia_TextBox.Value
    Else
        Dim f As Integer
        f = ultima_fila + 1
        HojaBD.Range("B" & f).Value = Me.codigo_TextBox.Value
        HojaBD.Range("C" & f).Value = Me.nombre_de_producto_TextBox.Value
        HojaBD.Range("D" & f).Value = CDate(Me.fecha_de_entrada_TextBox.Value)
        HojaBD.Range("E" & f).Value = Format(Round(CDbl(Me.cantidad_TextBox.Value), 3), "General Number")
        HojaBD.Range("F" & f).Value = Format(Round(CDbl(Me.cantidad_TextBox.Value), 3), "General Number")
        HojaBD.Range("G" & f).Value = Me.precio_venta_TextBox.Value
        HojaBD.Range("H" & f).Value = Me.precio_costo_TextBox.Value
        HojaBD.Range("I" & f).Value = CDbl(Me.precio_costo_TextBox.Value) * CDbl(Me.cantidad_TextBox.Value)
        HojaBD.Range("J" & f).Value = Me.Ingreso_TextBox.Value
        HojaBD.Range("K" & f).Value = Me.Ganancia_TextBox.Value
    End If
    
    ' Registrar en tablas auxiliares
    Call modAgregarRegistros.AgregarProductoUnico(Me.nombre_de_producto_TextBox.Value)
    Call modAgregarRegistros.AgregarFecha(CDate(Me.fecha_de_entrada_TextBox.Value))
    
    ' Ordenar tabla por nombre
    HojaBD.Range("C5").Sort key1:=HojaBD.Range("C5"), Header:=xlYes, key2:=HojaBD.Range("D5"), Header:=xlYes
    
    Unload Me
End Sub

Private Sub boton_ingresar_reposo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.boton_ingresar_covered.ZOrder msoBringToFront
    Me.Label8.ZOrder msoBringToFront
End Sub

' --- MODIFICAR PRODUCTO EXISTENTE -------------------------------------------
Private Sub boton_modificar_covered_Click()
    Dim fila_activa As Integer
    fila_activa = Application.ActiveCell.Row
    
    Call modValidacionCampos.ValidarVaciosProducto(Me)
    
    If contador > 0 Then
        MsgBox "Rellene los campos vacíos", vbExclamation
        Exit Sub
    End If
    
    ' Escribir en la hoja activa (fila seleccionada por el usuario)
    Range("B" & fila_activa).Value = Me.codigo_TextBox.Value
    Range("C" & fila_activa).Value = Me.nombre_de_producto_TextBox.Value
    Range("D" & fila_activa).Value = CDate(Me.fecha_de_entrada_TextBox.Value)
    Range("E" & fila_activa).Value = Me.cantidad_TextBox.Value
    Range("G" & fila_activa).Value = Me.precio_venta_TextBox.Value
    Range("H" & fila_activa).Value = Me.precio_costo_TextBox.Value
    Range("I" & fila_activa).Value = CDbl(Me.precio_costo_TextBox.Value) * CDbl(Me.cantidad_TextBox.Value)
    Range("J" & fila_activa).Value = Me.Ingreso_TextBox.Value
    Range("K" & fila_activa).Value = Me.Ganancia_TextBox.Value
    
	Call ForzarActualizacionStock
    Call modAgregarRegistros.AgregarProductoUnico(Me.nombre_de_producto_TextBox.Value)
    
    Unload Me
End Sub

Private Sub boton_modificar_reposo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.boton_modificar_covered.Visible = True
    Me.Label9.Visible = True
    Me.boton_modificar_covered.ZOrder msoBringToFront
    Me.Label9.ZOrder msoBringToFront
End Sub

Private Sub fondo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.boton_ingresar_reposo.ZOrder msoBringToFront
    Me.boton_modificar_reposo.ZOrder msoBringToFront
    Me.boton_calendario_reposo.ZOrder msoBringToFront
End Sub

Private Sub Label8_Click()
    boton_ingresar_covered_Click
End Sub

Private Sub Label9_Click()
    boton_modificar_covered_Click
End Sub

' --- CÁLCULOS UI ---------------------------------------------------------

Private Sub cantidad_TextBox_Change()
    CalcularIngresoBeneficio
End Sub

Private Sub precio_venta_TextBox_Change()
    CalcularIngresoBeneficio
End Sub

Private Sub precio_costo_TextBox_Change()
    If Me.cantidad_TextBox.Value <> "" And Me.precio_venta_TextBox.Value <> "" And Me.precio_costo_TextBox.Value <> "" Then
        Dim beneficioEsperado As Double
        beneficioEsperado = CDbl(Me.Ingreso_TextBox.Value) - (CDbl(Me.precio_costo_TextBox.Value) * CDbl(Me.cantidad_TextBox.Value))
        Me.Ganancia_TextBox.Value = Format(Round(beneficioEsperado, 3), "General Number")
    Else
        Me.Ganancia_TextBox.Value = ""
    End If
End Sub

Private Sub CalcularIngresoBeneficio()
    If Me.cantidad_TextBox.Value <> "" And Me.precio_venta_TextBox.Value <> "" Then
        Dim ingreso As Double
        ingreso = CDbl(Me.precio_venta_TextBox.Value) * CDbl(Me.cantidad_TextBox.Value)
        Me.Ingreso_TextBox.Value = Format(Round(ingreso, 3), "General Number")
        
        If Me.precio_costo_TextBox.Value <> "" Then
            Dim beneficio As Double
            beneficio = ingreso - (CDbl(Me.precio_costo_TextBox.Value) * CDbl(Me.cantidad_TextBox.Value))
            Me.Ganancia_TextBox.Value = Format(Round(beneficio, 3), "General Number")
        End If
    Else
        Me.Ingreso_TextBox.Value = ""
        Me.Ganancia_TextBox.Value = ""
    End If
End Sub

' --- VALIDACIÓN KeyPress ---------------------------------------------------

Private Sub cantidad_TextBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Call modValidacionNumerica.SoloNumeros(KeyAscii)
End Sub

Private Sub precio_venta_TextBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Call modValidacionNumerica.SoloNumeros(KeyAscii)
End Sub

Private Sub precio_costo_TextBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Call modValidacionNumerica.SoloNumeros(KeyAscii)
End Sub

' --- INICIALIZACIÓN -------------------------------------------------------

Private Sub UserForm_Activate()
    If Me.modificar_registrar_textbox.Value = "R" Then
        Me.boton_modificar_reposo.Visible = False
        Me.boton_ingresar_reposo.Visible = True
        
        Call modGeneradorCodigo.GenerarCodigoProducto
        Me.codigo_TextBox.Value = Codigo_Producto
        Me.fecha_de_entrada_TextBox.Value = Format(Date, "d-mmm-yyyy")
    End If
    
    If Me.modificar_registrar_textbox.Value = "M" Then
        Me.boton_modificar_reposo.Visible = True
        Me.boton_ingresar_reposo.Visible = False
        Me.Label9.Visible = True
        
        With Me
            .codigo_TextBox.Value = Range("B" & fila_a_modificar).Value
            .nombre_de_producto_TextBox.Value = Range("C" & fila_a_modificar).Value
            .fecha_de_entrada_TextBox.Value = Format(Range("D" & fila_a_modificar).Value, "d-mmm-yyyy")
            .cantidad_TextBox.Value = Format(Round(Range("E" & fila_a_modificar).Value, 3), "General Number")
            .precio_venta_TextBox.Value = Format(Round(Range("G" & fila_a_modificar).Value, 3), "General Number")
            .precio_costo_TextBox.Value = Format(Round(Range("H" & fila_a_modificar).Value, 3), "General Number")
        End With
    End If
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.boton_ingresar_reposo.ZOrder msoBringToFront
    Me.Label8.ZOrder msoBringToFront
    Me.boton_modificar_reposo.ZOrder msoBringToFront
    Me.Label9.ZOrder msoBringToFront
    Me.boton_calendario_reposo.ZOrder msoBringToFront
End Sub
