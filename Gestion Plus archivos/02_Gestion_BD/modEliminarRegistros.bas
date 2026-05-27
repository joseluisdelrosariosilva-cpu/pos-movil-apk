Attribute VB_Name = "modEliminarRegistros"
' ============================================================================
' MÓDULO: modEliminarRegistros
' PROPÓSITO: Consolidación de todas las funciones de eliminación de
' registros en las tablas del sistema.
' ============================================================================

Option Explicit

' --- ELIMINAR FILA DE TABLA POR SELECCIÓN (genérico) -----------------------
' Helper para eliminar una fila seleccionada de cualquier tabla.

Private Function PrepararEliminacion(tbl As ListObject) As ListRow
    Dim selected As Range
    Dim selected_row As ListRow
    
    Set selected = Selection
    
    If Intersect(selected, tbl.Range) Is Nothing Then
        Set PrepararEliminacion = Nothing
        Exit Function
    End If
    
    Set selected_row = tbl.ListRows(selected.Row - tbl.HeaderRowRange.Row)
    Set PrepararEliminacion = selected_row
End Function

' --- ELIMINAR GASTO --------------------------------------------------------

Public Sub EliminarGasto()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim selected_row As ListRow
    Dim response As VbMsgBoxResult
    
    Set ws = ThisWorkbook.Sheets(HOJA_GASTOS)
    Set tbl = ws.ListObjects(TBL_GASTOS)
    Set selected_row = PrepararEliminacion(tbl)
    
    If selected_row Is Nothing Then
        MsgBox "Por favor seleccione un gasto dentro de la tabla", vbExclamation
        Exit Sub
    End If
    
    response = MsgBox("¿Desea eliminar el gasto: " & selected_row.Range(COL_GAS_DESCRIPCION).Value & "?", _
                      vbYesNo + vbQuestion, "Confirmar eliminación")
    
    If response = vbYes Then
        ws.Unprotect password:=PASS_HOJA
        selected_row.Delete
        ws.Protect password:=PASS_HOJA, UserInterfaceOnly:=True
        MsgBox "Gasto eliminado correctamente", vbInformation
    End If
End Sub

' --- ELIMINAR TRABAJADOR ---------------------------------------------------

Public Sub EliminarTrabajador()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim selected_row As ListRow
    Dim response As VbMsgBoxResult
    
    Set ws = ThisWorkbook.Sheets(HOJA_TRABAJADORES)
    Set tbl = ws.ListObjects(TBL_TRABAJADORES)
    Set selected_row = PrepararEliminacion(tbl)
    
    If selected_row Is Nothing Then
        MsgBox "Por favor seleccione un trabajador dentro de la tabla", vbExclamation
        Exit Sub
    End If
    
    response = MsgBox("¿Desea eliminar el trabajador: " & selected_row.Range(1).Value & "?", _
                      vbYesNo + vbQuestion, "Confirmar eliminación")
    
    If response = vbYes Then
        ws.Unprotect password:=PASS_HOJA
        selected_row.Delete
        ws.Protect password:=PASS_HOJA, UserInterfaceOnly:=True
        MsgBox "Trabajador eliminado correctamente", vbInformation
    End If
End Sub

' --- ELIMINAR INGREDIENTE --------------------------------------------------

Public Sub EliminarIngrediente()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim selected_row As ListRow
    Dim response As VbMsgBoxResult
    
    Set ws = ThisWorkbook.Sheets(HOJA_INGREDIENTES)
    Set tbl = ws.ListObjects(TBL_INGREDIENTES)
    Set selected_row = PrepararEliminacion(tbl)
    
    If selected_row Is Nothing Then
        MsgBox "Por favor seleccione un ingrediente dentro de la tabla", vbExclamation
        Exit Sub
    End If
    
    response = MsgBox("¿Desea eliminar el ingrediente: " & selected_row.Range(1).Value & "?", _
                      vbYesNo + vbQuestion, "Confirmar eliminación")
    
    If response = vbYes Then
        ws.Unprotect password:=PASS_HOJA
        selected_row.Delete
        ws.Protect password:=PASS_HOJA, UserInterfaceOnly:=True
        MsgBox "Ingrediente eliminado correctamente", vbInformation
    End If
End Sub

' --- ELIMINAR PRODUCTO (rango directo) -------------------------------------

Public Sub EliminarProducto()
    Dim fila_a_eliminar As Integer
    Dim respuesta As String
    
    fila_a_eliminar = Application.ActiveCell.Row
    
    If fila_a_eliminar > 4 And Range("B" & fila_a_eliminar) <> "" Then
        respuesta = MsgBox("¿Está seguro que desea eliminar el producto seleccionado?", vbYesNo)
        
        If respuesta = 6 Then
            Range("B" & fila_a_eliminar).EntireRow.Delete
        End If
    Else
        MsgBox "Seleccione la línea con el producto que desea eliminar"
    End If
End Sub

' --- ELIMINAR RECETA COMPLETA (con sus ingredientes) -----------------------

Public Sub EliminarReceta()
    Dim wsRecetas As Worksheet
    Dim wsIngredientes As Worksheet
    Dim tblRecetas As ListObject
    Dim tblIngredientesReceta As ListObject
    Dim celdaSeleccionada As Range
    Dim nombreReceta As String
    Dim respuesta As VbMsgBoxResult
    Dim filaTabla As Long
    Dim filaIngrediente As ListRow
    Dim i As Long
    
    On Error GoTo ErrorHandler
    
    Set wsRecetas = ThisWorkbook.Sheets(HOJA_RECETAS)
    Set wsIngredientes = ThisWorkbook.Sheets(HOJA_INGR_RECETA)
    Set tblRecetas = wsRecetas.ListObjects(TBL_RECETAS)
    Set tblIngredientesReceta = wsIngredientes.ListObjects(TBL_INGREDIENTES_RECETA)
    
    Set celdaSeleccionada = ActiveCell
    
    ' Validaciones
    If Not celdaSeleccionada.Parent.Name = wsRecetas.Name Then
        MsgBox "Por favor, seleccione una celda en la hoja '" & HOJA_RECETAS & "'.", vbExclamation
        Exit Sub
    End If
    
    If Intersect(celdaSeleccionada, tblRecetas.Range) Is Nothing Then
        MsgBox "Por favor, seleccione una celda dentro de la tabla de recetas.", vbExclamation
        Exit Sub
    End If
    
    filaTabla = celdaSeleccionada.Row - tblRecetas.HeaderRowRange.Row
    If filaTabla <= 0 Then
        MsgBox "Por favor, seleccione una celda en una fila de datos.", vbExclamation
        Exit Sub
    End If
    
    nombreReceta = tblRecetas.DataBodyRange.Cells(filaTabla, 1).Value
    
    respuesta = MsgBox("¿Está seguro que desea eliminar la receta: '" & nombreReceta & "'?" & vbCrLf & _
                       "Esta acción no se puede deshacer.", _
                       vbYesNo + vbExclamation + vbDefaultButton2, "Confirmar eliminación")
    
    If respuesta = vbNo Then Exit Sub
    
    Application.ScreenUpdating = False
    
    wsRecetas.Unprotect password:=PASS_HOJA
    wsIngredientes.Unprotect password:=PASS_HOJA
    
    ' Eliminar ingredientes asociados (recorrido inverso)
    For i = tblIngredientesReceta.ListRows.Count To 1 Step -1
        Set filaIngrediente = tblIngredientesReceta.ListRows(i)
        If filaIngrediente.Range.Cells(1, 1).Value = nombreReceta Then
            filaIngrediente.Delete
        End If
    Next i
    
    ' Eliminar la receta
    tblRecetas.ListRows(filaTabla).Delete
    
    tblRecetas.Sort.Apply
    wsRecetas.Protect password:=PASS_HOJA, UserInterfaceOnly:=True
    
    tblIngredientesReceta.Sort.Apply
    wsIngredientes.Protect password:=PASS_HOJA, UserInterfaceOnly:=True
    
    Application.ScreenUpdating = True
    MsgBox "Receta '" & nombreReceta & "' eliminada exitosamente.", vbInformation
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Error al eliminar la receta: " & Err.Description, vbCritical
End Sub
