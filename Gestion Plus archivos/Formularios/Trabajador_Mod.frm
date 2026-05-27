VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Trabajador_Mod 
   Caption         =   "Modificar Trabajador"
   ClientHeight    =   4770
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   9160.001
   OleObjectBlob   =   "Trabajador_Mod.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Trabajador_Mod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

' ============================================================================
' FORMULARIO: Trabajador_Mod
' PROPÓSITO:  Interfaz para modificar un trabajador existente.
'              Toda la lógica de negocio delega en modTrabajadores.
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

Private Sub Btn_Active_Click()
    If Trim(Me.Txt_Nombre.Value) = "" Then
        MsgBox "Por favor ingrese un nombre", vbExclamation, "Campo requerido"
        Me.Txt_Nombre.SetFocus
        Exit Sub
    End If
    
    If Trim(Me.Txt_Cargo.Value) = "" Then
        MsgBox "Por favor ingrese un cargo", vbExclamation, "Campo requerido"
        Me.Txt_Cargo.SetFocus
        Exit Sub
    End If
    
    If Trim(Me.Txt_sal.Value) = "" Then
        MsgBox "Por favor ingrese un salario", vbExclamation, "Campo requerido"
        Me.Txt_sal.SetFocus
        Exit Sub
    End If
    
    If Trim(Me.Txt_sal_min.Value) = "" Then
        MsgBox "Por favor ingrese un salario mínimo", vbExclamation, "Campo requerido"
        Me.Txt_sal_min.SetFocus
        Exit Sub
    End If
    
    If Not IsNumeric(Me.Txt_sal.Value) Or Me.Txt_sal.Value < 0 Or Me.Txt_sal.Value > 100 Then
        MsgBox "Por favor ingrese un valor numérico entre 0 y 100 en el Salario %", vbExclamation, "Porcentaje inválido"
        Me.Txt_sal.SetFocus
        Exit Sub
    End If
    
    If Not IsNumeric(Me.Txt_sal_min.Value) Or Me.Txt_sal_min.Value < 0 Then
        MsgBox "Por favor ingrese un valor numérico positivo en el salario mínimo", vbExclamation, "Valor inválido"
        Me.Txt_sal_min.SetFocus
        Exit Sub
    End If
    
    Call modTrabajadores.ModificarTrabajadorEnTabla( _
        fila_sel, Me.Txt_Nombre.Value, Me.Txt_Cargo.Value, _
        CDbl(Me.Txt_sal.Value) / 100, _
        CDbl(Me.Txt_sal_min.Value))
    
    MsgBox "Trabajador modificado correctamente", vbInformation, "Listo!"
    Unload Trabajador_Mod
End Sub

Private Sub UserForm_Initialize()
    Dim tbl As ListObject
    Dim sel_row As ListRow
    
    Set tbl = ThisWorkbook.Sheets(HOJA_TRABAJADORES).ListObjects(TBL_TRABAJADORES)
    Set sel_row = tbl.ListRows(fila_sel)
    
    With sel_row
        Me.Txt_Nombre = .Range(COL_TRAB_NOMBRE)
        Me.Txt_Cargo = .Range(COL_TRAB_CARGO)
        Me.Txt_sal = Format(Round(.Range(COL_TRAB_PORCENTAJE) * 100, 3), "General Number")
        Me.Txt_sal_min = Format(Round(.Range(COL_TRAB_SAL_MIN), 3), "General Number")
    End With
End Sub
