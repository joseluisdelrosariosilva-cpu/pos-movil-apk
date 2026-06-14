Attribute VB_Name = "calling"
' ============================================================================
' MÓDULO: calling (MIGRADO - Fase 2)
' PROPÓSITO: Macros llamadoras desde botones y formularios.
' Los form-callers se mantienen; la lógica de negocio delega.
' ============================================================================
' NOTA: Las variables públicas fila_a_modificar, seleccion_destino_calendario,
' modoModificacion, recetaAModificar están declaradas en modGlobales.bas.
' ============================================================================

Option Explicit

' --- FORM-CALLERS (se mantienen igual) -------------------------------------

Sub llamar_ingresar_form()
    ingresar_form.modificar_registrar_textbox.Value = "R"
    ingresar_form.Show
End Sub

Sub llamar_Elaborar()
    Elaborar.Show
End Sub

Sub llamar_modificar_form()
    Dim respuesta As String
    fila_a_modificar = ActiveCell.Row

    If fila_a_modificar > 4 And Range("B" & fila_a_modificar) <> "" Then
        respuesta = MsgBox("¿Está seguro que desea modificar el producto seleccionado?", vbYesNo, "Eliminar producto")
        Select Case respuesta
            Case Is = 6
                ingresar_form.modificar_registrar_textbox.Value = "M"
                ingresar_form.Show
            Case Is = 7
                Unload ingresar_form
                Exit Sub
        End Select
    Else
        MsgBox "Seleccione la línea con el producto que desea modificar", vbExclamation
    End If
End Sub

Sub llamar_formulario_facturacion()
    facturacion.Show
End Sub

Sub reinicio_POS()
    Unload facturacion
    MsgBox "Venta realizada con éxito", vbInformation
    facturacion.Show
End Sub

' --- FORZAR ACTUALIZACIÓN DE FÓRMULAS (delega en modStock) -----------------

Sub setear_funcion()
    Call modStock.ForzarActualizacionStock
End Sub

' --- MODIFICAR RECETA (interfaz: validación + abrir formulario) -----------

Sub ModificarReceta()
    Dim ws As Worksheet
    Dim tblRecetas As ListObject
    Dim celdaSeleccionada As Range
    Dim filaSeleccionada As Long
    Dim nombreReceta As String
    
    Set ws = ThisWorkbook.Sheets(HOJA_RECETAS)
    Set tblRecetas = ws.ListObjects(TBL_RECETAS)
    
    On Error Resume Next
    Set celdaSeleccionada = Intersect(Selection, tblRecetas.Range)
    On Error GoTo 0
    
    If celdaSeleccionada Is Nothing Then
        MsgBox "Por favor, seleccione una celda dentro de la tabla de recetas", vbExclamation
        Exit Sub
    End If
    
    filaSeleccionada = celdaSeleccionada.Row
    nombreReceta = ws.Cells(filaSeleccionada, tblRecetas.ListColumns("Receta").Range.Column).Value
    
    Dim respuesta As VbMsgBoxResult
    respuesta = MsgBox("¿Está seguro que desea modificar la receta: '" & nombreReceta & "'?", _
                      vbYesNo + vbQuestion, "Confirmar modificación")
    
    If respuesta = vbYes Then
        ShowModoModificacion nombreReceta
    End If
End Sub

Public Sub ShowModoModificacion(nombreReceta As String)
    modoModificacion = True
    recetaAModificar = nombreReceta
    Ingr_Receta.Show
End Sub

' --- FORM-CALLERS DE INGREDIENTES (se mantienen) ---------------------------

Sub llamar_ingresar_nuevo_ingrediente()
    ingresar_ingr_nuevo_form.modificar_ingresar_nuevo_ingr_textbox.Value = "I"
    ingresar_ingr_nuevo_form.Show
End Sub

Sub llamar_modificar_form_ingrediente()
    Dim respuesta As String
    fila_a_modificar = ActiveCell.Row

    If fila_a_modificar > 4 And Range("B" & fila_a_modificar) <> "" Then
        respuesta = MsgBox("¿Está seguro que desea modificar el producto seleccionado?", vbYesNo, "Eliminar producto")
        Select Case respuesta
            Case Is = 6
                ingresar_ingr_nuevo_form.modificar_ingresar_nuevo_ingr_textbox = "M"
                ingresar_ingr_nuevo_form.Show
            Case Is = 7
                Unload ingresar_ingr_nuevo_form
                Exit Sub
        End Select
    Else
        MsgBox "Seleccione la línea con el producto que desea modificar", vbExclamation
    End If
End Sub

Sub llamar_ingresar_ingr_existente()
    ingresar_ingr_existente_form.Show
End Sub

Sub llamar_ingresarProdExistente()
    ingresarProdExistente.Show
End Sub


