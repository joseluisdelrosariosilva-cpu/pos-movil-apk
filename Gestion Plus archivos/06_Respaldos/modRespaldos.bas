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

' --- CREAR RESPALDO Y LIMPIAR BD (versión original) ------------------------
' Crea un respaldo en Excel .xlsx (solo datos, sin macros) y limpia la BD.

Public Sub CrearRespaldo()
    Dim wbNuevo As Workbook
    Dim wsOrigenVentas As Worksheet, wsOrigenGastos As Worksheet
    Dim wsDestinoVentas As Worksheet, wsDestinoGastos As Worksheet
    Dim tblventas As ListObject, tblGastos As ListObject
    Dim nombreArchivo As String
    Dim fechaMin As Date, fechaMax As Date
    Dim rutaRespaldos As String
    Dim ActualToInicial As ListRow
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    On Error GoTo ErrorHandler
    
    Set wsOrigenVentas = ThisWorkbook.Sheets(HOJA_VENTAS)
    Set tblventas = wsOrigenVentas.ListObjects(TBL_VENTAS)
    Set wsOrigenGastos = ThisWorkbook.Sheets(HOJA_GASTOS)
    Set tblGastos = wsOrigenGastos.ListObjects(TBL_GASTOS)
    
    ' Calcular rango de fechas
    fechaMin = WorksheetFunction.Min(ThisWorkbook.Sheets(HOJA_FECHAS).ListObjects(TBL_FECHAS_UNICAS).ListColumns("Fecha").DataBodyRange)
    fechaMax = WorksheetFunction.Max(ThisWorkbook.Sheets(HOJA_FECHAS).ListObjects(TBL_FECHAS_UNICAS).ListColumns("Fecha").DataBodyRange)
    
    nombreArchivo = "Respaldo_" & Format(fechaMin, "d-mmm-yyyy") & "_" & Format(fechaMax, "d-mmm-yyyy") & ".xlsx"
    rutaRespaldos = Environ("Jose") & "D:\Work\Respaldos\"
    
    If Dir(rutaRespaldos, vbDirectory) = "" Then MkDir rutaRespaldos
    
    ' Crear nuevo libro con los datos
    Set wbNuevo = Workbooks.Add
    tblventas.HeaderRowRange.Copy
    Set wsDestinoVentas = wbNuevo.Sheets(1)
    wsDestinoVentas.Name = "Historial de ventas"
    wsDestinoVentas.Range("A1").PasteSpecial Paste:=xlPasteAll
    
    If Not tblventas.DataBodyRange Is Nothing Then
        tblventas.DataBodyRange.Copy
        wsDestinoVentas.Range("A2").PasteSpecial Paste:=xlPasteAll
    End If
    
    tblGastos.HeaderRowRange.Copy
    Set wsDestinoGastos = wbNuevo.Sheets.Add(After:=wbNuevo.Sheets(wbNuevo.Sheets.Count))
    wsDestinoGastos.Name = "Historial de Gastos"
    wsDestinoGastos.Range("A1").PasteSpecial Paste:=xlPasteAll
    
    If Not tblGastos.DataBodyRange Is Nothing Then
        tblGastos.DataBodyRange.Copy
        wsDestinoGastos.Range("A2").PasteSpecial Paste:=xlPasteAll
    End If
    
    Application.CutCopyMode = False
    wbNuevo.SaveAs rutaRespaldos & nombreArchivo, FileFormat:=xlOpenXMLWorkbook
    wbNuevo.Close SaveChanges:=False
    
    ' Restaurar stock inicial = actual y limpiar tablas
    For Each ActualToInicial In ThisWorkbook.Worksheets(HOJA_ALMACEN).ListObjects(TBL_INVENTARIO).ListRows
        ActualToInicial.Range(COL_INV_CANT_INI) = ActualToInicial.Range(COL_INV_CANT_ACT)
    Next ActualToInicial
    
    LimpiarTabla tblventas
    LimpiarTabla tblGastos
    LimpiarTabla ThisWorkbook.Sheets(HOJA_FECHAS).ListObjects(TBL_FECHAS_UNICAS)
    
    MsgBox "Respaldo completado exitosamente." & vbCrLf & _
           "Archivo guardado en: " & rutaRespaldos & nombreArchivo, vbInformation
    
Exitsub:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    Resume Exitsub
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
            If ws.Shapes(i).Name Like "Boton_*" Then
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
