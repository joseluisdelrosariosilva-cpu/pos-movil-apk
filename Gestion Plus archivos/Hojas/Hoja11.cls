VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' ================================================
' HOJA11 (MIGRADO - Fase 2)
' PROPOSITO: Eventos de la hoja de Recetas.
' Al seleccionar una receta, guarda el nombre en L3 para referencia.
' ================================================

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim tbl As ListObject, selectedRow As ListRow
    Dim receta As String, selectedCell As Range
    Set tbl = ThisWorkbook.Sheets(HOJA_RECETAS).ListObjects("Recetas")
    If Not Application.Intersect(Target, tbl.Range) Is Nothing Then
        Set selectedCell = Target.Cells(1, 1)
        On Error Resume Next
        Set selectedRow = tbl.ListRows(selectedCell.Row - tbl.HeaderRowRange.Row)
        On Error GoTo 0
        If Not selectedRow Is Nothing Then
            receta = selectedRow.Range(1)
            Me.Range("L3") = receta
        End If
    End If
End Sub
