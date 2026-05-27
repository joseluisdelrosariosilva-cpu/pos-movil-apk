Attribute VB_Name = "Calling2"
' ============================================================================
' MÓDULO: Módulo1 (MIGRADO - Fase 2)
' PROPÓSITO: Macros asignadas a botones en las hojas.
' Toda la lógica de negocio ahora delega en los módulos consolidados.
' Los form-callers (Sub que solo abren formularios) se mantienen igual.
' ============================================================================

Option Explicit

' --- FORM-CALLERS (solo abren formularios, se mantienen igual) -------------

Sub call_Nueva_Receta_form()
    Ingr_Receta.Show
End Sub

Sub call_busqueda_form()
    Busqueda.Show
End Sub

Sub call_busqueda_Ventas_form()
    Busqueda_Ventas.Show
End Sub

Sub call_busqueda_Facturacion_form()
    Busqueda_Facturacion.Show
End Sub

Sub call_Gasto_Form()
    Gasto_Form.Show
End Sub

Sub call_Trabajador_Form()
    Trabajador.Show
End Sub

' --- MODIFICAR TRABAJADOR --------------------------------------------------

Sub call_Modificar_Form()
    Dim ws As Worksheet
    Dim tbl_sel As ListObject
    Dim selected As Range
    Dim selected_row As ListRow
    Dim response As VbMsgBoxResult
    
    Set ws = ThisWorkbook.Sheets(HOJA_TRABAJADORES)
    Set tbl_sel = ws.ListObjects(TBL_TRABAJADORES)
    Set selected = Selection
    
    If Intersect(selected, tbl_sel.Range) Is Nothing Then
        MsgBox "Por favor seleccione un trabajador dentro de la tabla", vbExclamation, "Selección no válida"
        Exit Sub
    End If
    
    Set selected_row = tbl_sel.ListRows(selected.Row - tbl_sel.HeaderRowRange.Row)
    response = MsgBox("Desea modificar el trabajador: " & selected_row.Range(1).Value & "?", _
                                    vbYesNo + vbQuestion, "Confirmar modificación")
    
    If response = vbYes Then
        fila_sel = selected_row.Index
        Trabajador_Mod.Show
    End If
End Sub

' (refresh_Dashboard fue reasignado directamente a modDashboard.RefrescarDashboard)

' --- DEBUG: VER VALOR DE HIDDEN.A1 -----------------------------------------

Sub verA1()
    MsgBox ThisWorkbook.Sheets(HOJA_HIDDEN).Range("A1")
End Sub

