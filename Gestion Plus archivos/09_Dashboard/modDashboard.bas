Attribute VB_Name = "modDashboard"
' ============================================================================
' MÓDULO: modDashboard
' PROPÓSITO: Refresco del Dashboard con datos actualizados.
' ============================================================================

Option Explicit

' --- REFRESCAR DASHBOARD ---------------------------------------------------

Public Sub RefrescarDashboard()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(HOJA_DASHBOARD)
    
    ws.Unprotect password:=PASS_HOJA
    ThisWorkbook.RefreshAll
    
    ws.Protect password:=PASS_HOJA, _
        contents:=True, _
        AllowFiltering:=True, _
        AllowSorting:=True, _
        AllowUsingPivotTables:=True, _
        UserInterfaceOnly:=True
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error al refrescar Dashboard: " & Err.Description, vbCritical
End Sub

' --- VERIFICAR EXPIRACIÓN (fecha en Hidden) --------------------------------
' Si la fecha en Hidden.A1 ha pasado, muestra un mensaje.

Public Function VerificarExpiracion() As Boolean
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim fechaExpiracion As Date
    
    Set ws = ThisWorkbook.Sheets(HOJA_HIDDEN)
    
    On Error Resume Next
    fechaExpiracion = CDate(ws.Range("A1").Value)
    On Error GoTo ErrorHandler
    
    If fechaExpiracion > 0 And fechaExpiracion < Date Then
        MsgBox "La versión de prueba ha expirado.", vbExclamation
        VerificarExpiracion = True
    Else
        VerificarExpiracion = False
    End If
    
    Exit Function
    
ErrorHandler:
    VerificarExpiracion = False
End Function

' --- MOSTRAR/OCULTAR HOJAS DEL SISTEMA ------------------------------------

Public Sub MostrarHojasSistema()
    Dim ws As Worksheet
    
    DesprotegerLibro
    ThisWorkbook.Sheets(HOJA_FECHAS).Visible = xlSheetVisible
    ProtegerLibro
End Sub

Public Sub OcultarHojasSistema()
    DesprotegerLibro
    ThisWorkbook.Sheets(HOJA_FECHAS).Visible = xlSheetVeryHidden
    ProtegerLibro
End Sub

' --- ACTUALIZAR FORMAS DEL DASHBOARD CON VALORES DE PIVOT TABLES ------------
' Se ejecuta cada vez que se actualiza una tabla dinámica en el Dashboard.
' Actualiza las formas (Shape_Ingresos, Shape_Gastos, etc.) con los totales.

Public Sub ActualizarFormasDashboard()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim valor_Ingresos As Variant
    Dim valor_Gastos As Variant
    Dim valor_Beneficio As Variant
    Dim valor_Margen As Double
    Dim valor_Fondo As Variant

    Set ws = ThisWorkbook.Sheets(HOJA_DASHBOARD)
    Set pt = ws.PivotTables("Ingresos_Din")

    valor_Ingresos = pt.TableRange2.Cells(pt.TableRange2.Rows.Count, 2)
    valor_Gastos = pt.TableRange2.Cells(pt.TableRange2.Rows.Count, 3)
    valor_Beneficio = pt.TableRange2.Cells(pt.TableRange2.Rows.Count, 4)
    valor_Fondo = ws.PivotTables("Fondo_Total").TableRange2.Cells( _
        ws.PivotTables("Fondo_Total").TableRange2.Rows.Count, 1)

    If IsNumeric(valor_Ingresos) And IsNumeric(valor_Gastos) And _
       IsNumeric(valor_Beneficio) And IsNumeric(valor_Fondo) Then
        ws.Shapes("Shape_Ingresos").TextFrame.Characters.Text = _
            "Ingreso Total: " & vbCrLf & " CUP   " & Format(Round(valor_Ingresos, 3), "General Number")
        ws.Shapes("Shape_Fondo").TextFrame.Characters.Text = _
            "Fondo Total: " & vbCrLf & " CUP   " & Format(Round(valor_Fondo, 3), "General Number")
        ws.Shapes("Shape_Gastos").TextFrame.Characters.Text = _
            "Gasto Total: " & vbCrLf & " CUP   " & Format(Round(valor_Gastos, 3), "General Number")
        ws.Shapes("Shape_Beneficio").TextFrame.Characters.Text = _
            "Beneficio: " & vbCrLf & " CUP   " & Format(Round(valor_Beneficio, 3), "General Number")

        On Error GoTo MargenCero
        valor_Margen = CDbl(valor_Beneficio) * 100 / CDbl(valor_Ingresos)
        ws.Shapes("Shape_Margen").TextFrame.Characters.Text = _
            "Margen:" & vbCrLf & Format(Round(valor_Margen, 2), "General Number") & " %"
    Else
        ws.Shapes("Shape_Ingresos").TextFrame.Characters.Text = "Ingreso Total: " & vbCrLf & " CUP   0"
        ws.Shapes("Shape_Fondo").TextFrame.Characters.Text = "Fondo Total: " & vbCrLf & " CUP   0"
        ws.Shapes("Shape_Gastos").TextFrame.Characters.Text = "Gasto Total: " & vbCrLf & " CUP   0"
        ws.Shapes("Shape_Beneficio").TextFrame.Characters.Text = "Beneficio: " & vbCrLf & " CUP   0"
        ws.Shapes("Shape_Margen").TextFrame.Characters.Text = "Margen:" & vbCrLf & "0 %"
    End If

    Exit Sub

MargenCero:
    ws.Shapes("Shape_Margen").TextFrame.Characters.Text = "Margen: " & vbCrLf & "0 %"
    Resume Next
ErrorHandler:
    ' Silenciar errores de shapes faltantes o pivots vacíos
    Resume Next
End Sub
