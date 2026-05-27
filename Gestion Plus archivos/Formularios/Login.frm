VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Login 
   Caption         =   "Acceso de administrador"
   ClientHeight    =   3450
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   9210.001
   OleObjectBlob   =   "Login.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

' ============================================================================
' FORMULARIO: Login
' PROPÓSITO:  Acceso de administrador con verificación de contraseña.
'              Toda la lógica delega en modSeguridad.
' ============================================================================

Private Sub boton_aceptar_covered_Click()
    If Me.pass_TextBox.Value = "" Then
        MsgBox "Por favor ingrese la contraseña para continuar", vbExclamation
        Exit Sub
    End If
    
    If modSeguridad.VerificarPasswordAdmin(Me.pass_TextBox.Value) Then
        Application.Visible = True
        Call modSeguridad.DesprotegerLibro
        Call modSeguridad.MostrarHojasAdministrativas(True)
        Call modSeguridad.ProtegerLibro
        
        Application.DisplayFullScreen = True
        ActiveWindow.DisplayHeadings = False
        
        Unload Me
        Unload facturacion
    Else
        MsgBox "Contraseña incorrecta", vbCritical
    End If
End Sub

Private Sub boton_cambiar_covered_Click()
    Dim inputCheckPassword As Variant
    Dim newCheckPassword As Variant
    
    ' Primer InputBox - contraseña actual
    inputCheckPassword = Application.InputBox("Ingrese la contraseña actual", "Cambio de contraseña", Type:=2)
    
    If inputCheckPassword = False Then
        Me.boton_cambiar_reposo.ZOrder msoBringToFront
        Me.Label5.ZOrder msoBringToFront
        Exit Sub
    End If
    
    ' Validar contraseña actual
    If modSeguridad.VerificarPasswordAdmin(CStr(inputCheckPassword)) Then
        ' Segundo InputBox - nueva contraseña
        newCheckPassword = Application.InputBox("Ingrese la nueva contraseña", "Cambio de contraseña", Type:=2)
        
        If newCheckPassword = False Then Exit Sub
        
        Call modSeguridad.CambiarPasswordAdmin(CStr(newCheckPassword))
        MsgBox "La contraseña se cambió correctamente", vbInformation
    Else
        MsgBox "Contraseña incorrecta", vbCritical
    End If
End Sub

Private Sub Label5_Click()
    boton_cambiar_covered_Click
End Sub

Private Sub Label6_Click()
    boton_aceptar_covered_Click
End Sub

Private Sub pass_TextBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then Call Label6_Click
End Sub

' --- EVENTOS MOUSE MOVE -------------------------------------------------------
Private Sub boton_aceptar_reposo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.boton_aceptar_covered.ZOrder msoBringToFront
    Me.Label6.ZOrder msoBringToFront
End Sub

Private Sub boton_cambiar_reposo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.boton_cambiar_covered.ZOrder msoBringToFront
    Me.Label5.ZOrder msoBringToFront
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.boton_aceptar_reposo.ZOrder msoBringToFront
    Me.Label6.ZOrder msoBringToFront
    Me.boton_cambiar_reposo.ZOrder msoBringToFront
    Me.Label5.ZOrder msoBringToFront
End Sub
