Attribute VB_Name = "modServerControl"
' ============================================================================
' MÓDULO: modServerControl
' PROPÓSITO: Control del servidor Node.js del POS Móvil.
' Separado del monolito serverControl.bas original.
' ============================================================================
' SOLO contiene: iniciar/detener/verificar servidor.
' La importación de datos está en modImportacion.
' La sincronización está en modSincronizacion.
' ============================================================================

Option Explicit

' --- FUNCIONES DE RUTA -----------------------------------------------------

Private Function ObtenerRutaBase() As String
    ObtenerRutaBase = ThisWorkbook.Path & "\" & CARPETA_WEBAPP
End Function

Private Function ObtenerRutaFlags() As String
    ObtenerRutaFlags = ObtenerRutaBase() & CARPETA_FLAGS
End Function

Private Function ObtenerRutaServidor() As String
    ObtenerRutaServidor = ObtenerRutaBase() & "src\server.js"
End Function

' --- VERIFICAR SI EL SERVIDOR ESTÁ ACTIVO ----------------------------------

Public Function ServidorEstaActivo() As Boolean
    ServidorEstaActivo = (Dir(ObtenerRutaFlags() & ARCHIVO_FLAG_ACTIVO) <> "")
End Function

' --- INICIAR SERVIDOR ------------------------------------------------------

Public Function IniciarServidor() As Boolean
    On Error GoTo ErrorHandler
    
    Dim comando As String
    
    comando = "cmd /c START /B node """ & ObtenerRutaServidor() & """"
    Shell comando, vbHide
    
    ' Esperar 3 segundos para que el servidor inicie
    Application.Wait Now + TimeValue("00:00:03")
    
    IniciarServidor = ServidorEstaActivo()
    Exit Function
    
ErrorHandler:
    MsgBox "Error al iniciar servidor: " & Err.Description, vbCritical
    IniciarServidor = False
End Function

' --- ENVIAR SEÑAL DE DETENCIÓN --------------------------------------------

Public Sub EnviarSenalDetencion()
    On Error Resume Next
    
    Dim flagPath As String
    flagPath = ObtenerRutaFlags() & ARCHIVO_FLAG_DETENER
    
    Open flagPath For Output As #1
    Print #1, "detener"
    Close #1
End Sub

' --- DETENER SERVIDOR ------------------------------------------------------

Public Sub DesactivarPOSMovil()
    If Not ServidorEstaActivo() Then
        MsgBox "El servidor no está activo.", vbExclamation
        Call RestaurarInterfaz
        Unload frmPOSMovil
        Exit Sub
    End If
    
    Call EnviarSenalDetencion
    
    Dim espera As Integer
    espera = 0
    Do While ServidorEstaActivo() And espera < 10
        Application.Wait Now + TimeValue("00:00:01")
        espera = espera + 1
        DoEvents
    Loop
    
    If ServidorEstaActivo() Then
        MsgBox "El servidor no respondió. Puedes cerrarlo manualmente.", vbExclamation
    Else
        MsgBox "Servidor detenido correctamente", vbInformation
    End If
    
    Call RestaurarInterfaz
    Unload frmPOSMovil
End Sub

' --- OBTENER IP LOCAL ------------------------------------------------------

Public Function ObtenerIPLocal() As String
    On Error GoTo Fallback
    
    Dim wmi As Object
    Dim adaptador As Object
    Dim ips As Variant
    Dim ip As Variant
    
    Set wmi = GetObject("winmgmts:\\.\root\cimv2")
    
    For Each adaptador In wmi.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")
        ips = adaptador.IPAddress
        If Not IsNull(ips) Then
            For Each ip In ips
                If EsIPv4(CStr(ip)) And CStr(ip) <> "127.0.0.1" Then
                    ObtenerIPLocal = "http://" & CStr(ip) & ":3000"
                    Exit Function
                End If
            Next ip
        End If
    Next adaptador
    
Fallback:
    ObtenerIPLocal = "http://localhost:3000"
End Function

Private Function EsIPv4(ByVal valor As String) As Boolean
    Dim bloques As Variant
    
    If InStr(1, valor, ":") > 0 Then Exit Function ' IPv6
    bloques = Split(valor, ".")
    If UBound(bloques) <> 3 Then Exit Function
    
    EsIPv4 = True
End Function

' --- INTERFAZ DE USUARIO (mostrar/ocultar) ---------------------------------

Public Sub OcultarInterfaz()
    DesprotegerLibro
    
    ThisWorkbook.Sheets(HOJA_INGREDIENTES).visible = xlSheetVeryHidden
    ThisWorkbook.Sheets(HOJA_RECETAS).visible = xlSheetVeryHidden
    ThisWorkbook.Sheets(HOJA_GASTOS).visible = xlSheetVeryHidden
    ThisWorkbook.Sheets(HOJA_FACTURAS).visible = xlSheetVeryHidden
    ThisWorkbook.Sheets(HOJA_VENTAS).visible = xlSheetVeryHidden
    ThisWorkbook.Sheets(HOJA_ALMACEN).visible = xlSheetVeryHidden
    ThisWorkbook.Sheets(HOJA_DASHBOARD).visible = xlSheetVeryHidden
    ThisWorkbook.Sheets(HOJA_BLANCO).visible = xlSheetVisible
    ThisWorkbook.Sheets(HOJA_TRABAJADORES).visible = xlSheetVeryHidden
    
    ProtegerLibro
    Call ProtegerTodasLasHojas
End Sub

Public Sub RestaurarInterfaz()
    DesprotegerLibro
    
    ThisWorkbook.Sheets(HOJA_INGREDIENTES).visible = xlSheetVisible
    ThisWorkbook.Sheets(HOJA_RECETAS).visible = xlSheetVisible
    ThisWorkbook.Sheets(HOJA_GASTOS).visible = xlSheetVisible
    ThisWorkbook.Sheets(HOJA_FACTURAS).visible = xlSheetVisible
    ThisWorkbook.Sheets(HOJA_VENTAS).visible = xlSheetVisible
    ThisWorkbook.Sheets(HOJA_ALMACEN).visible = xlSheetVisible
    ThisWorkbook.Sheets(HOJA_DASHBOARD).visible = xlSheetVisible
    ThisWorkbook.Sheets(HOJA_TRABAJADORES).visible = xlSheetVisible
    ThisWorkbook.Sheets(HOJA_BLANCO).visible = xlSheetVeryHidden
    
    ProtegerLibro
    
    Application.DisplayFullScreen = True
    ActiveWindow.DisplayHeadings = False
End Sub

' --- CONTAR VENTAS PENDIENTES DESDE LOG ------------------------------------

Public Function ContarVentasPendientes() As Integer
    On Error GoTo ErrorHandler
    
    Dim rutaLog As String
    Dim archivo As Integer
    Dim contenido As String
    Dim lineas() As String
    Dim i As Integer
    Dim contador As Integer
    
    rutaLog = ObtenerRutaBase() & ARCHIVO_VENTAS_LOG
    
    If Dir(rutaLog) = "" Then
        ContarVentasPendientes = 0
        Exit Function
    End If
    
    archivo = FreeFile
    Open rutaLog For Input As #archivo
    contenido = Input$(LOF(archivo), archivo)
    Close #archivo
    
    lineas = Split(contenido, vbLf)
    contador = 0
    
    For i = 0 To UBound(lineas)
        If Trim(lineas(i)) <> "" Then contador = contador + 1
    Next i
    
    ContarVentasPendientes = contador
    Exit Function
    
ErrorHandler:
    ContarVentasPendientes = 0
End Function

