VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' ============================================================================
' HOJA10 (MIGRADO - Fase 2)
' PROPOSITO: Eventos de la hoja "Gestión de Ingredientes".
' - Color coding automatico en columna D (existencia vs stock minimo).
' - Recalculo de costos de recetas cuando cambia precio de ingrediente.
' Toda la logica de negocio delega en modulos reorganizados.
' ============================================================================
Option Explicit
Private Sub Worksheet_Change(ByVal Target As Range)
    On Error Resume Next
    Application.EnableEvents = False
    Call AplicarColorCodingColumnaD
    Call RecalcularRecetasPorPrecioIngrediente(Target)
    Application.EnableEvents = True
End Sub
' --- COLOR CODING EN COLUMNA D ----------------------------------------------
Private Sub AplicarColorCodingColumnaD()
    Dim celda As Range, rango As Range
    Dim valorD As Double, valorC As Double, porcentaje As Double
    Set rango = Me.Range("D5:D1000")
    For Each celda In rango
        If IsEmpty(celda) Or celda.Value = "" Then
            celda.Interior.Color = vbBlack
            GoTo SiguienteCelda
        End If
        valorD = celda.Value
        valorC = celda.Offset(0, -1).Value
        If valorC = 0 Then
            celda.Interior.Color = vbBlack
            GoTo SiguienteCelda
        End If
        porcentaje = valorD / valorC
        Select Case True
            Case porcentaje >= 0.8: celda.Interior.Color = RGB(0, 176, 80)
            Case porcentaje >= 0.5: celda.Interior.Color = RGB(255, 255, 0)
            Case porcentaje >= 0.2: celda.Interior.Color = RGB(255, 192, 0)
            Case Else: celda.Interior.Color = RGB(255, 0, 0)
        End Select
SiguienteCelda:
    Next celda
End Sub
' --- RECALCULAR RECETAS CUANDO CAMBIA PRECIO DE INGREDIENTE -----------------
Private Sub RecalcularRecetasPorPrecioIngrediente(ByVal Target As Range)
    Dim wsIng As Worksheet, tblIng As ListObject
    Dim rngChanged As Range, celda As Range
    Dim ingrediente As String, colPrecio As Long
    Set wsIng = ThisWorkbook.Sheets(HOJA_INGREDIENTES)
    Set tblIng = wsIng.ListObjects(TBL_INGREDIENTES)
    Set rngChanged = Intersect(Target, tblIng.DataBodyRange)
    If rngChanged Is Nothing Then Exit Sub
    On Error Resume Next
    colPrecio = tblIng.ListColumns("Precio/U.Base").Index
    If colPrecio = 0 Then colPrecio = 5
    On Error GoTo 0
    Set rngChanged = Intersect(rngChanged, tblIng.ListColumns(colPrecio).Range)
    If rngChanged Is Nothing Then Exit Sub
    For Each celda In rngChanged
        ingrediente = tblIng.DataBodyRange.Cells(celda.Row - tblIng.HeaderRowRange.Row, 1).Value
        Call modElaboracion.ActualizarRecetasPorIngrediente(ingrediente)
    Next celda
End Sub
