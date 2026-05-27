Attribute VB_Name = "modValidacionNumerica"
' ============================================================================
' MÓDULO: modValidacionNumerica
' PROPÓSITO: Validación universal de entrada numérica en TextBox.
' Reemplaza los ~30 eventos KeyPress duplicados en los formularios.
' ============================================================================
' USO: En cada TextBox que solo acepte números, llama así:
'   Private Sub txtCantidad_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
'       Call SoloNumeros(KeyAscii)
'   End Sub
' ============================================================================

Option Explicit

' --- VALIDACIÓN: SOLO NÚMEROS Y PUNTO DECIMAL -----------------------------

Public Sub SoloNumeros(ByRef KeyAscii As MSForms.ReturnInteger)
    ' Permite dígitos (48-57) y punto (46)
    If (KeyAscii >= 46 And KeyAscii <= 57) Then
        ' Permitir: 0-9 y punto decimal
    Else
        KeyAscii = 0  ' Bloquear cualquier otra tecla
    End If
End Sub
