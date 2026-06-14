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
' Proposito: Te permite seleccionar un Excel de la misma carpeta y reescribe
'            los datos de TODAS las tablas (ListObjects) de TODAS las hojas
'            del libro actual con los datos del archivo seleccionado.
' ============================================================================

Sub ImportarDatosDesdeExcel()
    Dim srcFile   As Variant
    Dim srcWB     As Workbook
    Dim ws        As Worksheet
    Dim lo        As ListObject
    Dim srcLO     As ListObject
    Dim importados As Long, omitidos As Long
    
    ' ---- 1. Seleccionar archivo origen ----
    srcFile = Application.GetOpenFilename( _
        FileFilter:="Archivos Excel (*.xls*),*.xls*", _
        Title:="Seleccioná el archivo del cual importar los datos", _
        MultiSelect:=False)
    
    If srcFile = False Then
        MsgBox "No seleccionaste ningún archivo. Operación cancelada.", _
               vbInformation, "Importación cancelada"
        Exit Sub
    End If
    
    ' ---- 2. Validar que no sea el mismo archivo ----
    If ThisWorkbook.fullName = srcFile Then
        MsgBox "Seleccionaste el mismo archivo que está en uso. Elegí otro.", _
               vbExclamation, "Archivo inválido"
        Exit Sub
    End If
    
    ' ---- 3. Abrir archivo origen (solo lectura) ----
    On Error GoTo ErrorAbrir
    Set srcWB = Workbooks.Open(srcFile, ReadOnly:=True, UpdateLinks:=False)
    On Error GoTo 0
    
    ' ---- 4. Desproteger libro destino ----
    On Error Resume Next
    ThisWorkbook.Unprotect PASS_LIBRO
    On Error GoTo 0
    
    ' ---- 5. Procesar cada tabla del libro DESTINO ----
    importados = 0
    omitidos = 0
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    For Each ws In ThisWorkbook.Worksheets
        ' Desproteger hoja antes de modificar
        On Error Resume Next
        ws.Unprotect PASS_HOJA
        On Error GoTo 0
        
        For Each lo In ws.ListObjects
            Set srcLO = FindListObjectByName(srcWB, lo.name)
            
            If srcLO Is Nothing Then
                omitidos = omitidos + 1
            Else
                CopyTableData lo, srcLO
                importados = importados + 1
            End If
        Next lo
        
        ' Reproteger hoja (UserInterfaceOnly deja que el VBA siga modificando)
        On Error Resume Next
        ws.Protect password:=PASS_HOJA, UserInterfaceOnly:=True
        On Error GoTo 0
    Next ws
    
    ' Reproteger libro
    On Error Resume Next
    ThisWorkbook.Protect password:=PASS_LIBRO
    On Error GoTo 0
    
    srcWB.Close SaveChanges:=False
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    ' ---- 6. Mostrar resultado ----
    MsgBox "Importación completada." & vbCrLf & vbCrLf & _
           "? Tablas importadas: " & importados & vbCrLf & _
           "? Tablas no encontradas en origen: " & omitidos, _
           vbInformation, "Resultado de importación"
    
    Exit Sub

ErrorAbrir:
    MsgBox "No se pudo abrir el archivo:" & vbCrLf & srcFile & vbCrLf & vbCrLf & _
           "Error: " & Err.Description, vbCritical, "Error al abrir"
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub

Private Sub CopyTableData(ByVal destLO As ListObject, ByVal srcLO As ListObject)
    Dim vData As Variant
    Dim r     As Long
    
    If srcLO.DataBodyRange Is Nothing Then
        If Not destLO.DataBodyRange Is Nothing Then
            destLO.DataBodyRange.ClearContents
        End If
        Exit Sub
    End If
    
    vData = srcLO.DataBodyRange.Value
    
    Do While destLO.ListRows.Count > 0
        destLO.ListRows(destLO.ListRows.Count).Delete
    Loop
    
    If Not IsEmpty(vData) Then
        For r = 1 To UBound(vData, 1)
            destLO.ListRows.Add
        Next r
        
        destLO.DataBodyRange.Value = vData
    End If
End Sub

Private Function FindListObjectByName(ByVal wb As Workbook, ByVal name As String) As ListObject
    Dim ws As Worksheet
    Dim lo As ListObject
    
    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            If lo.name = name Then
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
'            Las tablas quedan vacías, solo con los encabezados.
' ============================================================================

Sub LimpiarTodasLasTablas()
    Dim ws  As Worksheet
    Dim lo  As ListObject
    Dim respuesta As VbMsgBoxResult
    
    ' ---- 1. Confirmar ----
    respuesta = MsgBox("¿Estás seguro de que querés ELIMINAR TODOS LOS DATOS " & _
                       "de TODAS las tablas?" & vbCrLf & vbCrLf & _
                       "Las filas se borran por completo. Esta acción NO se puede deshacer.", _
                       vbYesNo + vbExclamation, "Limpiar todas las tablas")
    
    If respuesta <> vbYes Then
        MsgBox "Operación cancelada.", vbInformation, "Cancelado"
        Exit Sub
    End If
    
    ' ---- 2. Desproteger libro ----
    On Error Resume Next
    ThisWorkbook.Unprotect PASS_LIBRO
    On Error GoTo 0
    
    ' ---- 3. Limpiar cada tabla ----
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    For Each ws In ThisWorkbook.Worksheets
        ' Desproteger hoja
        On Error Resume Next
        ws.Unprotect PASS_HOJA
        On Error GoTo 0
        
        For Each lo In ws.ListObjects
            ' Eliminar todas las filas de datos de abajo hacia arriba
            Do While lo.ListRows.Count > 0
                lo.ListRows(lo.ListRows.Count).Delete
            Loop
        Next lo
        
        ' Reproteger hoja
        On Error Resume Next
        ws.Protect password:=PASS_HOJA, UserInterfaceOnly:=True
        On Error GoTo 0
    Next ws
    
    ' Reproteger libro
    On Error Resume Next
    ThisWorkbook.Protect password:=PASS_LIBRO
    On Error GoTo 0
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    MsgBox "Todas las tablas quedaron vacías (solo encabezados).", vbInformation, "Completado"
End Sub

