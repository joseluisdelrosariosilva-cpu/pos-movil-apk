VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ingresar_ingr_existente_form 
   Caption         =   "Ingresar Ingrediente Existente"
   ClientHeight    =   4600
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   10010
   OleObjectBlob   =   "ingresar_ingr_existente_form.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "ingresar_ingr_existente_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

' ============================================================================
' FORMULARIO: ingresar_ingr_existente_form
' PROPÓSITO:  Interfaz para agregar cantidad a un ingrediente existente.
'              Toda la lógica de negocio delega en módulos reorganizados.
' ============================================================================

Private Sub boton_ingresar_covered_Click()
    Call modValidacionCampos.ValidarVaciosIngredienteExistente(Me)
    
    If contador3 > 0 Then
        MsgBox "Rellene los campos vacíos", vbExclamation
        Me.boton_ingresar_reposo.ZOrder msoBringToFront
        Me.Label27.ZOrder msoBringToFront
        Exit Sub
    End If
    
    Dim celdaEncontrada As Range
    Dim fila As Integer
    Dim HojaBD As Worksheet
    Dim ingrediente As String
    Dim Unidad As String
    Dim cantidadSinConvertir As Double
    Dim cantidadConvertidaUBase As Double
    Dim cantidadActual As Double
    Dim fondoActual As Double
    Dim costo As Double
    
    Set HojaBD = ThisWorkbook.Sheets(HOJA_INGREDIENTES)
    
    ingrediente = Me.Nombre_ingr_ComboBox.Value
    cantidadSinConvertir = CDbl(Me.Cantidad_ingresar_producto_existente_TextBox.Value)
    Unidad = Me.Unidad_ComboBox.Value
    costo = CDbl(Me.Costo_TextBox.Value)
    
    ' Buscar ingrediente
    Application.ScreenUpdating = False
    Set celdaEncontrada = HojaBD.Columns("B").Find(ingrediente, LookIn:=xlValues, LookAt:=xlWhole)
    Application.ScreenUpdating = True
    
    If celdaEncontrada Is Nothing Then
        MsgBox "Ingrediente no encontrado", vbExclamation
        Exit Sub
    End If
    
    fila = celdaEncontrada.Row
    
    ' Convertir cantidad a unidad base usando módulo centralizado
    cantidadConvertidaUBase = modConversor.ConvertirUnidad(ingrediente, cantidadSinConvertir, Unidad)
    cantidadActual = HojaBD.Range("D" & fila).Value
    fondoActual = HojaBD.Range("G" & fila).Value
    
    HojaBD.Range("D" & fila).Value = cantidadConvertidaUBase + cantidadActual
    HojaBD.Range("G" & fila).Value = fondoActual + costo
    
    ' Recalcular precio por unidad base
    Call modCalculosInventario.CalcularPrecioUBase
    
    Unload Me
    MsgBox "Cantidad agregada correctamente", vbInformation
End Sub

Private Sub boton_ingresar_reposo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.boton_ingresar_covered.ZOrder msoBringToFront
    Me.Label27.ZOrder msoBringToFront
End Sub

Private Sub Cantidad_ingresar_producto_existente_TextBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Call modValidacionNumerica.SoloNumeros(KeyAscii)
End Sub

Private Sub Costo_TextBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Call modValidacionNumerica.SoloNumeros(KeyAscii)
End Sub

Private Sub Label27_Click()
    boton_ingresar_covered_Click
End Sub

' --- CARGAR UNIDADES SEGÚN INGREDIENTE SELECCIONADO -------------------------
Private Sub Nombre_ingr_ComboBox_Change()
    Dim valorSeleccionado As String
    Dim Unidad_base_valor As String
    Dim celdaEncontrada As Range
    Dim HojaBD As Worksheet
    Dim filaEncontrada As Integer
    
    valorSeleccionado = Me.Nombre_ingr_ComboBox.Value
    If valorSeleccionado = "" Then Exit Sub
    
    Set HojaBD = ThisWorkbook.Sheets(HOJA_INGREDIENTES)
    Set celdaEncontrada = HojaBD.Columns("B").Find(valorSeleccionado, LookIn:=xlValues, LookAt:=xlWhole)
    
    If celdaEncontrada Is Nothing Then Exit Sub
    
    filaEncontrada = celdaEncontrada.Row
    Unidad_base_valor = HojaBD.Range("E" & filaEncontrada).Value
    
    ' Delegar carga de unidades al módulo centralizado
    Call modRecetas.CargarUnidadesPorIngrediente(Me.Unidad_ComboBox, valorSeleccionado)
End Sub

Private Sub Nombre_ingr_ComboBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = 0
End Sub

Private Sub UserForm_Activate()
    Dim HojaBD As Worksheet
    Dim i As Integer
    Dim ultima_fila As Integer
    
    Set HojaBD = ThisWorkbook.Sheets(HOJA_INGREDIENTES)
    ultima_fila = HojaBD.Range("B" & Rows.Count).End(xlUp).Row
    
    Me.Nombre_ingr_ComboBox.Clear
    For i = 5 To ultima_fila
        If HojaBD.Range("B" & i).Value <> "" Then
            Me.Nombre_ingr_ComboBox.AddItem HojaBD.Range("B" & i).Value
        End If
    Next i
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.boton_ingresar_reposo.ZOrder msoBringToFront
    Me.Label27.ZOrder msoBringToFront
End Sub
