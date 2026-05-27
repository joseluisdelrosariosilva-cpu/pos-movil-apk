VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' ============================================================================
' HOJA9 (MIGRADO - Fase 2)
' PROPOSITO: Eventos de la hoja "Registro de Facturas".
' Al seleccionar una fila de la tabla, guarda el ID en I3 para referencia.
' ============================================================================

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim tbl As ListObject, selectedRow As ListRow
    Dim facturaID As String, selectedCell As Range
    Set tbl = ThisWorkbook.Sheets("Registro de Facturas").ListObjects("HISTORIALVENTAS4")
    If Not Application.Intersect(Target, tbl.Range) Is Nothing Then
        Set selectedCell = Target.Cells(1, 1)
        On Error Resume Next
        Set selectedRow = tbl.ListRows(selectedCell.Row - tbl.HeaderRowRange.Row)
        On Error GoTo 0
        If Not selectedRow Is Nothing Then
            facturaID = selectedRow.Range(1)
            Me.Range("I3") = facturaID
        End If
    End If
End Sub
