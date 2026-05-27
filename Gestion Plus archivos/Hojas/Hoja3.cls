VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' ============================================================================
' HOJA3 (MIGRADO - Fase 2)
' PROPOSITO: Eventos de la hoja oculta que redirige a Historial de Gastos.
' Delega proteccion en modSeguridad.
' ============================================================================

Private Sub Worksheet_Activate()
    Call modSeguridad.DesprotegerLibro
    ThisWorkbook.Worksheets("Historial de Gastos").Activate
    ThisWorkbook.Sheets("Hidden").Visible = xlSheetVeryHidden
    Call modSeguridad.ProtegerLibro
End Sub
