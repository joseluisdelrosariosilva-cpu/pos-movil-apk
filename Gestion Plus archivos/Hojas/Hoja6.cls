VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' ============================================================================
' HOJA6 (MIGRADO - Fase 2)
' PROPOSITO: Eventos de la hoja Dashboard.
' Actualiza las formas con valores de las tablas dinamicas al filtrar.
' Toda la logica de negocio delega en modDashboard.
' ============================================================================

Private Sub Worksheet_PivotTableUpdate(ByVal Target As PivotTable)
    Call modDashboard.ActualizarFormasDashboard
End Sub
