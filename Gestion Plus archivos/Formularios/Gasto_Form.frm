VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Gasto_Form 
   Caption         =   "Nuevo Gasto Variable"
   ClientHeight    =   4670
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   9800.001
   OleObjectBlob   =   "Gasto_Form.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Gasto_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

' ============================================================================
' FORMULARIO: Gasto_Form
' PROPÓSITO:  Interfaz para registrar un nuevo gasto variable.
'              Toda la lógica de negocio delega en módulos reorganizados.
' ============================================================================

Private Sub Ingresar_Click()
    Btn_Active_Click
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.Btn_Reposo.ZOrder msoBringToFront
    Me.Ingresar.ZOrder msoBringToFront
End Sub

Private Sub Btn_Reposo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.Btn_Active.ZOrder msoBringToFront
    Me.Ingresar.ZOrder msoBringToFront
End Sub

' --- VALIDAR CAMPOS Y REGISTRAR GASTO ---------------------------------------
Private Sub Btn_Active_Click()

    If Trim(Me.Txt_Fecha.Value) = "" Then
        MsgBox "Por favor ingrese una fecha", vbExclamation, "Campo requerido"
        Me.Txt_Fecha.SetFocus
        Exit Sub
    End If
    
    If Trim(Me.Txt_Monto.Value) = "" Then
        MsgBox "Por favor ingrese un monto", vbExclamation, "Campo requerido"
        Me.Txt_Monto.SetFocus
        Exit Sub
    End If
    
    If Trim(Me.Txt_Descripcion.Value) = "" Then
        MsgBox "Por favor ingrese una descripción", vbExclamation, "Campo requerido"
        Me.Txt_Descripcion.SetFocus
        Exit Sub
    End If
    
    If Not IsDate(Me.Txt_Fecha.Value) Then
        MsgBox "Por favor ingrese una fecha válida", vbExclamation, "Fecha inválida"
        Me.Txt_Fecha.SetFocus
        Exit Sub
    End If
    
    If Not IsNumeric(Me.Txt_Monto.Value) Or Me.Txt_Monto.Value < 0 Then
        MsgBox "Por favor ingrese un valor numérico positivo en el monto", vbExclamation, "Monto inválido"
        Me.Txt_Monto.SetFocus
        Exit Sub
    End If
    
    ' Delegar inserción a módulos reorganizados
    Call modAgregarRegistros.AgregarGasto( _
        CDate(Me.Txt_Fecha.Value), _
        CAT_VARIABLE, _
        Me.Txt_Descripcion.Value, _
        CDbl(Me.Txt_Monto.Value))
    
    Call modAgregarRegistros.AgregarFecha(CDate(Me.Txt_Fecha.Value))
    
    MsgBox "Gasto agregado correctamente", vbInformation, "Listo!"
    Unload Gasto_Form

End Sub

Private Sub Txt_Monto_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Call modValidacionNumerica.SoloNumeros(KeyAscii)
End Sub

Private Sub UserForm_Initialize()
    Txt_Fecha.Value = Format(Date, "d-mmm-yyyy")
End Sub
