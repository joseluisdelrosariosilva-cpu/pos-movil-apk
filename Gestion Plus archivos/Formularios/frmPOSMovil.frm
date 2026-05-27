VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPOSMovil 
   Caption         =   "Control POS Móvil"
   ClientHeight    =   6360
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   6960
   OleObjectBlob   =   "frmPOSMovil.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmPOSMovil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

' ============================================================================
' FORMULARIO: frmPOSMovil
' PROPÓSITO:  Control del servidor POS Móvil, timer de actualización y
'              desactivación. Delega la lógica en módulos de 07_POS_Movil/.
' ============================================================================

Private proximaActualizacion As Date
Private timerActivo As Boolean
Public desactivando As Boolean

Private Function ObtenerProcedimientoTimer() As String
    ObtenerProcedimientoTimer = "'" & ThisWorkbook.Name & "'!ActualizarDatosFormulario"
End Function

Private Sub boton_clipboard_covered_Click()
    On Error GoTo ControlarError
    Dim texto As String
    Dim url As String
    Dim p As Long
    Dim portapapeles As MSForms.DataObject
    
    texto = Trim$(lblIP.Caption)
    p = InStr(1, texto, "http", vbTextCompare)
    If p > 0 Then
        url = Mid$(texto, p)
    Else
        url = texto
    End If
    url = Trim$(url)
    
    If url = vbNullString Then
        MsgBox "No hay URL para copiar.", vbExclamation, "Aviso"
        Exit Sub
    End If
    
    Set portapapeles = New MSForms.DataObject
    portapapeles.SetText url
    portapapeles.PutInClipboard
    MsgBox "URL copiada: " & url, vbInformation, "Copiado"
    Exit Sub
    
ControlarError:
    MsgBox "No se pudo copiar la URL." & vbCrLf & _
           "Err.Number: " & Err.Number & vbCrLf & _
           "Err.Description: " & Err.Description, vbExclamation, "Error al copiar"
End Sub

Private Sub boton_desactivar_covered_Click()
    desactivando = True
    
    Dim respuesta As VbMsgBoxResult
    respuesta = MsgBox("¿Desactivar POS Móvil?" & vbCrLf & _
                      "Se cerrará el servidor y se importarán las ventas.", _
                      vbQuestion + vbYesNo, "Confirmar")
    
    If respuesta = vbYes Then
        DetenerTimer
        Call modImportacion.ImportarVentas
        Call modServerControl.DesactivarPOSMovil
    End If
End Sub

Private Sub CommandButton1_Click()
    Call modImportacion.ImportarVentas
End Sub

Private Sub Labeldesactivar_Click()
    boton_desactivar_covered_Click
End Sub

Private Sub btnImportar_Click()
    MsgBox "Función de importación en desarrollo", vbInformation
End Sub

' ============================================================================
' INICIALIZACIÓN
' ============================================================================
Private Sub UserForm_Initialize()
    Me.Caption = "Control POS Móvil"
    lblServidor.Caption = "Servidor: http://localhost:3000"
    
    If modServerControl.ServidorEstaActivo() Then
        lblEstado.Caption = "Estado: Conectado"
        lblEstado.ForeColor = &HFF00&    ' Verde
    Else
        lblEstado.Caption = "Estado: Desconectado"
        lblEstado.ForeColor = &HFF&      ' Rojo
    End If
    
    lblIP.Caption = "Móvil: " & modServerControl.ObtenerIPLocal()
    Call ActualizarVentasPendientes
    IniciarTimer
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If desactivando Then Exit Sub
    boton_desactivar_covered_Click
End Sub

Private Sub UserForm_Terminate()
    DetenerTimer
End Sub

' ============================================================================
' ACTUALIZAR VENTAS PENDIENTES (UI)
' ============================================================================
Sub ActualizarVentasPendientes()
    Dim ventasPendientes As Integer
    ventasPendientes = modServerControl.ContarVentasPendientes()
    
    If ventasPendientes > 0 Then
        lblVentas.Caption = "Ventas pendientes: " & ventasPendientes
        lblVentas.ForeColor = RGB(255, 100, 0)  ' Naranja
    Else
        lblVentas.Caption = "Ventas pendientes: 0"
        lblVentas.ForeColor = RGB(100, 100, 100)  ' Gris
    End If
End Sub

' ============================================================================
' TIMER DE ACTUALIZACIÓN (Application.OnTime)
' ============================================================================
Public Sub ActualizarDatosFormulario()
    On Error GoTo ControlarError
    
    If modServerControl.ServidorEstaActivo() Then
        lblEstado.Caption = "Estado: Conectado"
        lblEstado.ForeColor = &HFF00&
    Else
        lblEstado.Caption = "Estado: Desconectado"
        lblEstado.ForeColor = &HFF&
    End If
    
    ActualizarVentasPendientes
    
ProgramarSiguienteTick:
    If timerActivo Then
        On Error Resume Next
        proximaActualizacion = Now + TimeValue("00:00:05")
        Application.OnTime EarliestTime:=proximaActualizacion, _
            Procedure:=ObtenerProcedimientoTimer(), Schedule:=True
        On Error GoTo 0
    End If
    
    Exit Sub
    
ControlarError:
    Resume ProgramarSiguienteTick
End Sub

Sub IniciarTimer()
    Dim procedimientoTimer As String
    procedimientoTimer = ObtenerProcedimientoTimer()
    
    On Error Resume Next
    If timerActivo Then
        Application.OnTime EarliestTime:=proximaActualizacion, _
            Procedure:=procedimientoTimer, Schedule:=False
        timerActivo = False
    End If
    On Error GoTo 0
    
    proximaActualizacion = Now + TimeValue("00:00:05")
    Application.OnTime EarliestTime:=proximaActualizacion, _
        Procedure:=procedimientoTimer, Schedule:=True
    timerActivo = True
End Sub

Sub DetenerTimer()
    Dim procedimientoTimer As String
    procedimientoTimer = ObtenerProcedimientoTimer()
    
    On Error Resume Next
    If timerActivo Then
        Application.OnTime EarliestTime:=proximaActualizacion, _
            Procedure:=procedimientoTimer, Schedule:=False
    End If
    On Error GoTo 0
    
    timerActivo = False
    proximaActualizacion = 0
End Sub

' ============================================================================
' EVENTOS MOUSE MOVE
' ============================================================================
Private Sub boton_clipboard_reposo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.boton_clipboard_covered.ZOrder msoBringToFront
End Sub

Private Sub boton_desactivar_reposo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.boton_desactivar_covered.ZOrder msoBringToFront
    Me.Labeldesactivar.ZOrder msoBringToFront
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.boton_desactivar_reposo.ZOrder msoBringToFront
    Me.Labeldesactivar.ZOrder msoBringToFront
    Me.boton_clipboard_reposo.ZOrder msoBringToFront
End Sub
