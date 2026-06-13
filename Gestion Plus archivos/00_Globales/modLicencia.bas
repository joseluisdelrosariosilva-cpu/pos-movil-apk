Attribute VB_Name = "modLicencia"
' ============================================================================
' MODULO: modLicencia
' PROPOSITO: Logica de licencia, verificacion de dispositivo, integridad de
' fecha del sistema, expiracion y codigos de renovacion.
' Extraido de ThisWorkbook.cls.
' ============================================================================
Option Explicit
' --- VERIFICAR DISPOSITIVO AUTORIZADO ---------------------------------------
Public Function VerificarDispositivo() As Boolean
    Dim dispositivosPermitidos As Collection
    Set dispositivosPermitidos = New Collection
    dispositivosPermitidos.Add "DESKTOP-FFH10GP|Jose|279994388"
    dispositivosPermitidos.Add "DESKTOP-27I02D0|Aris|351847289"
    dispositivosPermitidos.Add "DESKTOP-6Q2K16M|Noel|1088704218"
    dispositivosPermitidos.Add "ROSÉ|Administrador|-561350944"
    Dim idActual As String
    idActual = ObtenerIDDispositivo()
    Dim i As Integer
    For i = 1 To dispositivosPermitidos.Count
        If dispositivosPermitidos(i) = idActual Then
            VerificarDispositivo = True
            Exit Function
        End If
    Next i
    VerificarDispositivo = False
End Function
Private Function ObtenerIDDispositivo() As String
    Dim id As String
    id = Environ("COMPUTERNAME") & "|" & Environ("USERNAME") & "|" & _
         CreateObject("Scripting.FileSystemObject").GetDrive("C:").SerialNumber
    ObtenerIDDispositivo = id
End Function
' --- VERIFICAR INTEGRIDAD DE FECHA DEL SISTEMA ------------------------------
Public Sub VerificarIntegridadFecha()
    Dim ultimaFecha As Date, FechaActual As Date
    Call modSeguridad.DesprotegerHoja(ThisWorkbook.Sheets(HOJA_HIDDEN))
    FechaActual = Date
    ultimaFecha = ThisWorkbook.Sheets(HOJA_HIDDEN).Range("A2")
    If ultimaFecha = 0 Then
        ThisWorkbook.Sheets(HOJA_HIDDEN).Range("A2") = FechaActual
        Call modSeguridad.ProtegerHoja(ThisWorkbook.Sheets(HOJA_HIDDEN))
        Exit Sub
    End If
    If FechaActual < ultimaFecha Then
        MsgBox "Se detecto un retroceso en la fecha del sistema. El libro se cerrara.", vbCritical
        ThisWorkbook.Close SaveChanges:=False
    End If
    If FechaActual > ultimaFecha Then
        If Not ThisWorkbook.Name Like "Respaldo_*" Then
            Call modRespaldos.CrearRespaldoDiario
        End If
        ThisWorkbook.Sheets(HOJA_HIDDEN).Range("A2") = FechaActual
    End If
    Call modSeguridad.ProtegerHoja(ThisWorkbook.Sheets(HOJA_HIDDEN))
End Sub
' --- VERIFICAR VIGENCIA -----------------------------------------------------
Public Function VerificarVigencia() As Boolean
    On Error GoTo ErrorHandler
    Dim ws As Worksheet, fechaExpiracion As Date
    Set ws = ThisWorkbook.Sheets(HOJA_HIDDEN)
    fechaExpiracion = ws.Range("A1")
    If Date >= fechaExpiracion - 3 And Date <= fechaExpiracion Then
        MsgBox "Este libro esta pronto a expirar. Contacte con su proveedor.", vbInformation
    End If
    VerificarVigencia = (Date <= fechaExpiracion)
    Exit Function
ErrorHandler:
    Call modSeguridad.DesprotegerHoja(ws)
    fechaExpiracion = DateAdd("d", 1, Date)
    ThisWorkbook.Sheets(HOJA_HIDDEN).Range("A1") = fechaExpiracion
    ThisWorkbook.Sheets(HOJA_HIDDEN).visible = xlSheetVeryHidden
    VerificarVigencia = True
    Call modSeguridad.ProtegerHoja(ws)
End Function
' --- SOLICITAR CODIGO DE RENOVACION -----------------------------------------
Public Function SolicitarCodigoRenovacion() As Boolean
    Dim ws As Worksheet, codigoIngresado As String
    Dim codigoValido_mes As String, codigoValido_anual As String
    Set ws = ThisWorkbook.Sheets(HOJA_HIDDEN)
    codigoValido_mes = GenerarCodigoMensual(Day(Date), Month(Date), Year(Date))
    codigoValido_anual = GenerarCodigoAnual(Day(Date), Month(Date), Year(Date))
    codigoIngresado = InputBox("El libro ha expirado. Por favor, ingrese el codigo de renovacion:", "Renovacion requerida")
    If codigoIngresado = codigoValido_mes Then
        Call modSeguridad.DesprotegerHoja(ws)
        ThisWorkbook.Sheets(HOJA_HIDDEN).Range("A1") = DateAdd("d", 30, Date)
        SolicitarCodigoRenovacion = True
        Call modSeguridad.ProtegerHoja(ws)
        MsgBox "Gracias. Se han aniadido 30 dias a la vigencia del libro", vbInformation
    ElseIf codigoIngresado = codigoValido_anual Then
        Call modSeguridad.DesprotegerHoja(ws)
        ThisWorkbook.Sheets(HOJA_HIDDEN).Range("A1") = DateAdd("d", 180, Date)
        SolicitarCodigoRenovacion = True
        Call modSeguridad.ProtegerHoja(ws)
        MsgBox "Gracias. Se han aniadido 180 dias a la vigencia del libro", vbInformation
    Else
        SolicitarCodigoRenovacion = False
    End If
End Function
Private Function GenerarCodigoMensual(dia As Integer, mes As Integer, Ano As Integer) As String
    Dim semilla As String
    semilla = "theMonthCodegenisthe100th" & CStr(dia * 149) & CStr(mes * 149) & CStr(Ano)
    GenerarCodigoMensual = HashString(semilla)
End Function
Private Function GenerarCodigoAnual(dia As Integer, mes As Integer, Ano As Integer) As String
    Dim semilla As String
    semilla = "theYearCodegenisthe101th" & CStr(dia * 149) & CStr(mes * 149) & CStr(Ano)
    GenerarCodigoAnual = HashString(semilla)
End Function
Private Function HashString(texto As String) As String
    Dim i As Integer, hash As Long
    hash = 1
    For i = 1 To Len(texto)
        hash = (hash * 31 + Asc(Mid(texto, i, 1))) Mod 1000000
    Next i
    HashString = CStr(hash)
End Function


