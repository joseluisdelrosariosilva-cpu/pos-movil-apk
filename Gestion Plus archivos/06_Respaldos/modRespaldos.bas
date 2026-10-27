Attribute VB_Name = "modRespaldos"
' ============================================================================
' MÓDULO: modRespaldos
' PROPÓSITO: Creación de respaldos y limpieza de base de datos.
' Consolidación de Módulo2 + la función crear_respaldo_diario de ThisWorkbook.
' ============================================================================

Option Explicit

' --- LIMPIAR TABLA (helper) ------------------------------------------------

Public Sub LimpiarTabla(tbl As ListObject)
    If Not tbl.DataBodyRange Is Nothing Then
        tbl.DataBodyRange.Delete
    End If
End Sub


' --- CREAR RESPALDO (versión completa con macros) --------------------------

Public Sub CrearRespaldoCompleto()
    Dim wbRespaldo As Workbook
    Dim wbOriginal As Workbook
    Dim ws As Worksheet
    Dim nombreRespaldo As String
    Dim fechaMin As Date, fechaMax As Date
    Dim rutaRespaldos As String
    Dim i As Long
    Dim response As VbMsgBoxResult
    
    response = MsgBox("¿Está seguro de que desea crear una copia de respaldo y limpiar la base de datos actual?", _
                      vbYesNo + vbQuestion)
    If Not response = vbYes Then Exit Sub
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    On Error GoTo ErrorHandler
    
    Set wbOriginal = ThisWorkbook
    
    fechaMin = WorksheetFunction.Min(wbOriginal.Sheets(HOJA_FECHAS).ListObjects(TBL_FECHAS_UNICAS).ListColumns("Fecha").DataBodyRange)
    fechaMax = WorksheetFunction.Max(wbOriginal.Sheets(HOJA_FECHAS).ListObjects(TBL_FECHAS_UNICAS).ListColumns("Fecha").DataBodyRange)
    
    nombreRespaldo = "Respaldo_" & Format(fechaMin, "d-mmm-yyyy") & "_" & Format(fechaMax, "d-mmm-yyyy") & ".xlsm"
    rutaRespaldos = wbOriginal.Path & CARPETA_RESPALDOS
    
    If Dir(rutaRespaldos, vbDirectory) = "" Then MkDir rutaRespaldos
    
    wbOriginal.SaveCopyAs rutaRespaldos & nombreRespaldo
    Set wbRespaldo = Workbooks.Open(rutaRespaldos & nombreRespaldo)
    
    ' Remover botones del respaldo
    For Each ws In wbRespaldo.Worksheets
        ws.Unprotect password:=PASS_HOJA
        For i = ws.Shapes.Count To 1 Step -1
            If ws.Shapes(i).name Like "Boton_*" Then
                ws.Shapes(i).Delete
            End If
        Next i
        ws.Protect password:=PASS_HOJA, UserInterfaceOnly:=True
    Next ws
    
    ' Proteger Dashboard correctamente
    wbRespaldo.Worksheets(HOJA_DASHBOARD).Protect _
        password:=PASS_HOJA, _
        contents:=True, _
        AllowFiltering:=True, _
        AllowSorting:=True, _
        AllowUsingPivotTables:=True, _
        UserInterfaceOnly:=True
    
    ' Sellar fecha de expiración
    wbRespaldo.Worksheets(HOJA_HIDDEN).Unprotect password:=PASS_HOJA
    wbRespaldo.Worksheets(HOJA_HIDDEN).Range("A1") = DateAdd("y", 10000, Date)
    wbRespaldo.Worksheets(HOJA_HIDDEN).Protect password:=PASS_HOJA, UserInterfaceOnly:=True
    
    wbRespaldo.Save
    wbRespaldo.Close
    
    ' Limpiar tablas originales
    LimpiarTabla wbOriginal.Sheets(HOJA_FECHAS).ListObjects(TBL_FECHAS_UNICAS)
    LimpiarTabla wbOriginal.Sheets(HOJA_GASTOS).ListObjects(TBL_GASTOS)
    LimpiarTabla wbOriginal.Sheets(HOJA_VENTAS).ListObjects(TBL_VENTAS)
    LimpiarTabla wbOriginal.Sheets(HOJA_FACTURAS).ListObjects(TBL_FACTURAS)
    
    MsgBox "Respaldo completado exitosamente." & vbCrLf & _
           "Archivo guardado en: " & rutaRespaldos & nombreRespaldo, vbInformation
    
Exitsub:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    Resume Exitsub
End Sub

' --- CREAR RESPALDO DIARIO -------------------------------------------------

Public Sub CrearRespaldoDiario()
    Dim wbOriginal As Workbook
    Dim nombreRespaldo As String
    Dim rutaRespaldos As String
    Dim respaldoAnterior As String
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    On Error GoTo ErrorHandler
    
    Set wbOriginal = ThisWorkbook
    nombreRespaldo = "Respaldo_diario_" & Format(Date, "d-mmm-yyyy") & ".xlsm"
    rutaRespaldos = wbOriginal.Path & CARPETA_RESPALDOS
    
    If Dir(rutaRespaldos, vbDirectory) = "" Then MkDir rutaRespaldos
    
    ' Reemplazar respaldo diario anterior
    respaldoAnterior = Dir(rutaRespaldos & "Respaldo_diario_*.xlsm")
    If respaldoAnterior <> "" Then Kill rutaRespaldos & respaldoAnterior
    
    wbOriginal.SaveCopyAs rutaRespaldos & nombreRespaldo
    
Exitsub:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    Resume Exitsub
End Sub

' --- CREAR RESPALDO FINAL (antes de cerrar) --------------------------------

Public Sub CrearRespaldoFinal()
    Dim wbOriginal As Workbook
    Dim nombreRespaldo As String
    Dim rutaRespaldos As String
    Dim respaldoAnterior As String
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    On Error GoTo ErrorHandler
    
    Set wbOriginal = ThisWorkbook
    nombreRespaldo = "Ultimo_respaldo_" & Format(Date, "d-mmm-yyyy") & "_" & Format(Time, "hh-mm AM/PM") & ".xlsm"
    rutaRespaldos = wbOriginal.Path & CARPETA_RESPALDOS
    
    If Dir(rutaRespaldos, vbDirectory) = "" Then MkDir rutaRespaldos
    
    respaldoAnterior = Dir(rutaRespaldos & "Ultimo_respaldo_*.xlsm")
    If respaldoAnterior <> "" Then Kill rutaRespaldos & respaldoAnterior
    
    wbOriginal.SaveCopyAs rutaRespaldos & nombreRespaldo
    
Exitsub:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    Resume Exitsub
End Sub

' ============================================================================
' Macro: ImportarDatosDesdeExcel
' Proposito: Reemplaza el contenido de cada tabla del libro actual con el de
'            la tabla homonima del libro que elija el usuario.
'            Reemplaza, no agrega. Las tablas locales que no existan en el
'            origen no se tocan.
' ============================================================================

Sub ImportarDatosDesdeExcel()
    Dim srcFile    As Variant
    Dim srcWB      As Workbook
    Dim ws         As Worksheet
    Dim lo         As ListObject
    Dim srcLO      As ListObject
    Dim importadas As Long
    Dim omitidas   As Long
    Dim incompat   As Long
    Dim mensaje    As String

    Const TITULO As String = "Importar datos desde otro libro"

    ' ---- 1. Elegir el libro origen ----
    srcFile = Application.GetOpenFilename( _
        FileFilter:="Libros de Excel (*.xls;*.xlsx;*.xlsm;*.xlsb), *.xls; *.xlsx; *.xlsm; *.xlsb", _
        Title:="Selecciona el libro del cual importar los datos", _
        MultiSelect:=False)

    If srcFile = False Then Exit Sub

    If StrComp(ThisWorkbook.fullName, CStr(srcFile), vbTextCompare) = 0 Then
        MsgBox "Seleccionaste el mismo libro que esta abierto. Elegi otro.", _
               vbExclamation, TITULO
        Exit Sub
    End If

    If MsgBox("Se van a REEMPLAZAR los datos de todas las tablas locales que " & _
              "existan tambien en el libro elegido." & vbCrLf & vbCrLf & _
              "Continuar?", vbYesNo + vbQuestion, TITULO) <> vbYes Then
        Exit Sub
    End If

    ' ---- 2. Abrir el origen sin ejecutar sus macros ----
    On Error GoTo ErrorAbrir

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.AutomationSecurity = msoAutomationSecurityForceDisable

    Set srcWB = Workbooks.Open(Filename:=CStr(srcFile), ReadOnly:=True, UpdateLinks:=0)

    On Error GoTo ErrorTrabajo

    ' ---- 3. Desproteger el destino ----
    modSeguridad.DesprotegerLibro
    modSeguridad.DesprotegerTodasLasHojas

    ' ---- 4. Copiar tabla por tabla ----
    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            Set srcLO = FindListObjectByName(srcWB, lo.name)

            If srcLO Is Nothing Then
                omitidas = omitidas + 1
            ElseIf CopyTableData(lo, srcLO) Then
                importadas = importadas + 1
            Else
                incompat = incompat + 1
            End If
        Next lo
    Next ws

    ' Si trajiste el inventario de otro libro, el contador tiene que alcanzar
    ' al codigo mas alto importado o el proximo producto lo repetiria.
    modGeneradorCodigo.SincronizarContador

    mensaje = "Importacion terminada." & vbCrLf & vbCrLf & _
              "Tablas importadas: " & importadas & vbCrLf & _
              "Tablas no encontradas en el origen: " & omitidas

    If incompat > 0 Then
        mensaje = mensaje & vbCrLf & _
                  "Tablas con columnas distintas (no se tocaron): " & incompat
    End If

Salir:
    On Error Resume Next
    If Not srcWB Is Nothing Then
        srcWB.Close SaveChanges:=False
        Set srcWB = Nothing
    End If
    modSeguridad.ProtegerTodasLasHojas
    modSeguridad.ProtegerLibro
    RestaurarAplicacion
    On Error GoTo 0

    If Len(mensaje) > 0 Then MsgBox mensaje, vbInformation, TITULO
    Exit Sub

ErrorAbrir:
    RestaurarAplicacion
    MsgBox "No se pudo abrir el archivo:" & vbCrLf & srcFile & vbCrLf & vbCrLf & _
           "Error: " & Err.Description, vbCritical, TITULO
    Exit Sub

ErrorTrabajo:
    mensaje = ""
    MsgBox "Error durante la importacion:" & vbCrLf & vbCrLf & _
           Err.Description & vbCrLf & vbCrLf & _
           "Revisa el libro antes de seguir usando.", vbCritical, TITULO
    Resume Salir
End Sub


' Reemplaza el contenido de destLO con el de srcLO.
' Devuelve False si no tienen las mismas columnas en el mismo orden, para no
' escribir datos en columnas equivocadas.
Private Function CopyTableData(ByVal destLO As ListObject, ByVal srcLO As ListObject) As Boolean
    Dim vDatos As Variant
    Dim i As Long
    Dim j As Long

    If srcLO.ListColumns.Count <> destLO.ListColumns.Count Then Exit Function

    For j = 1 To srcLO.ListColumns.Count
        If StrComp(Trim$(srcLO.ListColumns(j).name), _
                   Trim$(destLO.ListColumns(j).name), vbTextCompare) <> 0 Then
            Exit Function
        End If
    Next j

    LimpiarTabla destLO

    If srcLO.DataBodyRange Is Nothing Then
        CopyTableData = True
        Exit Function
    End If

    vDatos = srcLO.DataBodyRange.Value

    ' Con una sola fila y una sola columna Excel devuelve un escalar, no una matriz.
    If Not IsArray(vDatos) Then
        destLO.ListRows.Add
        destLO.DataBodyRange.Value = vDatos
        CopyTableData = True
        Exit Function
    End If

    For i = 1 To UBound(vDatos, 1)
        destLO.ListRows.Add
    Next i

    destLO.DataBodyRange.Value = vDatos
    CopyTableData = True
End Function


' Busca una tabla por nombre en todo el libro. Comparacion sin distinguir
' mayusculas: VBA compara texto como binario y "INVENTARIO" <> "Inventario".
Private Function FindListObjectByName(ByVal wb As Workbook, ByVal nombre As String) As ListObject
    Dim ws As Worksheet
    Dim lo As ListObject

    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.name, nombre, vbTextCompare) = 0 Then
                Set FindListObjectByName = lo
                Exit Function
            End If
        Next lo
    Next ws
End Function

' ============================================================================
' Macro: LimpiarTodasLasTablas
' Proposito: Elimina TODAS las filas de datos de TODAS las tablas
'            (ListObjects) de TODAS las hojas del libro actual.
'            Las tablas quedan vacias, solo con los encabezados.
' ============================================================================

Sub LimpiarTodasLasTablas()
    Dim ws        As Worksheet
    Dim lo        As ListObject
    Dim respuesta As VbMsgBoxResult
    Dim tablas    As Long
    Dim mensaje   As String

    Const TITULO As String = "Limpiar todas las tablas"

    respuesta = MsgBox( _
        "Esto borra TODAS las filas de TODAS las tablas del libro:" & vbCrLf & _
        "inventario, ingredientes, recetas, ventas, facturas, gastos," & vbCrLf & _
        "trabajadores, fechas y productos unicos." & vbCrLf & vbCrLf & _
        "Las tablas quedan solo con los encabezados." & vbCrLf & _
        "NO se puede deshacer. Continuar?", _
        vbYesNo + vbExclamation, TITULO)

    If respuesta <> vbYes Then Exit Sub

    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    modSeguridad.DesprotegerLibro
    modSeguridad.DesprotegerTodasLasHojas

    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            LimpiarTabla lo
            tablas = tablas + 1
        Next lo
    Next ws

    ' Descomenta esta linea si ademas queres que el proximo producto sea Pr_00001.
    modGeneradorCodigo.ResetearContadorCodigos

    mensaje = "Listo: " & tablas & " tablas quedaron vacias (solo encabezados)."

Salir:
    On Error Resume Next
    modSeguridad.ProtegerTodasLasHojas
    modSeguridad.ProtegerLibro
    RestaurarAplicacion
    On Error GoTo 0

    If Len(mensaje) > 0 Then MsgBox mensaje, vbInformation, TITULO
    Exit Sub

ErrorHandler:
    mensaje = "Se produjo un error y la limpieza quedo a medias:" & vbCrLf & vbCrLf & _
              Err.Description & vbCrLf & vbCrLf & _
              "Revisa el libro antes de seguir usando."
    Resume Salir
End Sub


Private Sub RestaurarAplicacion()
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.AutomationSecurity = msoAutomationSecurityByUI
    Application.CutCopyMode = False
End Sub
