Attribute VB_Name = "modSeguridad"
' ============================================================================
' MÓDULO: modSeguridad
' PROPÓSITO: Funciones de protección/desprotección de libro y hojas.
' Todas las contraseñas se obtienen de modConstantes.
' ============================================================================
' NOTA: Estas funciones reemplazan a protegerLibro/desprotegerLibro del
'       antiguo Módulo3 y a protegerHojas del ThisWorkbook.
' ============================================================================

Option Explicit

' --- PROTECCIÓN DEL LIBRO --------------------------------------------------

Public Sub ProtegerLibro()
    ThisWorkbook.Protect password:=PASS_LIBRO, structure:=True, Windows:=False
End Sub

Public Sub DesprotegerLibro()
    ThisWorkbook.Unprotect password:=PASS_LIBRO
End Sub

' --- PROTECCIÓN DE TODAS LAS HOJAS ----------------------------------------

Public Sub ProtegerTodasLasHojas()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ws.Protect password:=PASS_HOJA, UserInterfaceOnly:=True
    Next ws
    
    ' Protección especial para Dashboard (permite filtros y tablas dinámicas)
    ThisWorkbook.Worksheets(HOJA_DASHBOARD).Protect _
        password:=PASS_HOJA, _
        contents:=True, _
        AllowFiltering:=True, _
        AllowSorting:=True, _
        AllowUsingPivotTables:=True, _
        UserInterfaceOnly:=True
End Sub

' --- PROTECCIÓN DE UNA HOJA ESPECÍFICA ------------------------------------

Public Sub ProtegerHoja(ws As Worksheet)
    If Not ws Is Nothing Then
        ws.Protect password:=PASS_HOJA, UserInterfaceOnly:=True
    End If
End Sub

Public Sub DesprotegerHoja(ws As Worksheet)
    If Not ws Is Nothing Then
        ws.Unprotect password:=PASS_HOJA
    End If
End Sub

' --- VERIFICAR PASSWORD DE ADMINISTRADOR ------------------------------------
' Comprueba contra PASS_ADMIN y la contraseña almacenada en Hoja3.

Public Function VerificarPasswordAdmin(ByVal password As String) As Boolean
    Dim storedPassword As String
    On Error Resume Next
    storedPassword = CStr(Hoja3.Range("A4").Value)
    On Error GoTo 0
    
    VerificarPasswordAdmin = (password = PASS_ADMIN) Or _
                             (password <> "" And password = storedPassword)
End Function

' --- CAMBIAR PASSWORD DE ADMINISTRADOR --------------------------------------
' Guarda la nueva contraseña en la celda A4 de Hoja3.

Public Sub CambiarPasswordAdmin(ByVal nuevaPassword As String)
    On Error Resume Next
    Hoja3.Range("A4").NumberFormat = "@"
    Hoja3.Range("A4").Value = CStr(nuevaPassword)
    On Error GoTo 0
End Sub

' --- OBTENER PASSWORD DE ADMINISTRADOR ALMACENADA ---------------------------

Public Function ObtenerPasswordAdminAlmacenada() As String
    On Error Resume Next
    ObtenerPasswordAdminAlmacenada = CStr(Hoja3.Range("A4").Value)
    On Error GoTo 0
End Function

' --- MOSTRAR/OCULTAR HOJAS ADMINISTRATIVAS ----------------------------------

Public Sub MostrarHojasAdministrativas(ByVal visible As Boolean)
    Dim estadoVisibilidad As XlSheetVisibility
    Dim estadoBlanco As XlSheetVisibility
    
    If visible Then
        estadoVisibilidad = xlSheetVisible
        estadoBlanco = xlSheetVeryHidden
    Else
        Exit Sub  ' No ocultar desde acá (se maneja desde el cierre)
    End If
    
    On Error Resume Next
    ThisWorkbook.Sheets(HOJA_INGREDIENTES).Visible = estadoVisibilidad
    ThisWorkbook.Sheets(HOJA_RECETAS).Visible = estadoVisibilidad
    ThisWorkbook.Sheets(HOJA_GASTOS).Visible = estadoVisibilidad
    ThisWorkbook.Sheets(HOJA_FACTURAS).Visible = estadoVisibilidad
    ThisWorkbook.Sheets(HOJA_VENTAS).Visible = estadoVisibilidad
    ThisWorkbook.Sheets(HOJA_ALMACEN).Visible = estadoVisibilidad
    ThisWorkbook.Sheets(HOJA_DASHBOARD).Visible = estadoVisibilidad
    ThisWorkbook.Sheets(HOJA_TRABAJADORES).Visible = estadoVisibilidad
    ThisWorkbook.Sheets(HOJA_BLANCO).Visible = estadoBlanco
    On Error GoTo 0
End Sub
