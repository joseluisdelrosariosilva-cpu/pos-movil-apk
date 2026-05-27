VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' ============================================
' THISWORKBOOK (MIGRADO - Fase 2)
' PROPOSITO: Solo eventos del libro.
' Toda la logica delega en modulos reorganizados.
' ============================================

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Call modRespaldos.CrearRespaldoFinal
    Call modSeguridad.DesprotegerLibro
    ThisWorkbook.Sheets("Blanco").Visible = xlSheetVisible
    ThisWorkbook.Sheets("Historial de Gastos").Visible = xlSheetVeryHidden
    ThisWorkbook.Sheets("Registro de facturas").Visible = xlSheetVeryHidden
    ThisWorkbook.Sheets("Historial de ventas").Visible = xlSheetVeryHidden
    ThisWorkbook.Sheets(HOJA_ALMACEN).Visible = xlSheetVeryHidden
    ThisWorkbook.Sheets("Dashboard").Visible = xlSheetVeryHidden
    ThisWorkbook.Sheets(HOJA_TRABAJADORES).Visible = xlSheetVeryHidden
    ThisWorkbook.Sheets(HOJA_INGREDIENTES).Visible = xlSheetVeryHidden
    ThisWorkbook.Sheets(HOJA_RECETAS).Visible = xlSheetVeryHidden
    ThisWorkbook.Save
    Call modSeguridad.ProtegerLibro
    Application.Quit
End Sub
Private Sub Workbook_Open()
    Application.Visible = False
    If Not modLicencia.VerificarDispositivo Then
        MsgBox "Este libro no esta autorizado para su uso en este equipo", vbCritical
        ThisWorkbook.Close SaveChanges:=False
    End If
    Call modLicencia.VerificarIntegridadFecha
    ' If Not modLicencia.VerificarVigencia Then
    '     If Not modLicencia.SolicitarCodigoRenovacion Then
    '         MsgBox "El libro ha expirado y no se ha proporcionado un codigo valido.", vbCritical
    '         ThisWorkbook.Close SaveChanges:=False
    '     End If
    ' End If
    Call modSeguridad.ProtegerTodasLasHojas
    facturacion.Show
End Sub