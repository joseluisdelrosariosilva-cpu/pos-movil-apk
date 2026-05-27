Attribute VB_Name = "modRecetas"
' ============================================================================
' MÓDULO: modRecetas
' PROPÓSITO: Funciones reutilizables de CRUD para recetas e ingredientes.
' Extraídas de Ingr_Receta.frm y otros formularios.
' ============================================================================

Option Explicit

' --- VERIFICAR SI UNA RECETA YA EXISTE -------------------------------------

Public Function RecetaExiste(ByVal nombreReceta As String) As Boolean
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim fila As ListRow
    
    Set ws = ThisWorkbook.Sheets(HOJA_RECETAS)
    Set tbl = ws.ListObjects(TBL_RECETAS)
    
    RecetaExiste = False
    
    For Each fila In tbl.ListRows
        If LCase(fila.Range.Cells(1, COL_REC_NOMBRE).Value) = LCase(Trim(nombreReceta)) Then
            RecetaExiste = True
            Exit Function
        End If
    Next fila
End Function

' --- CARGAR NOMBRES DE RECETAS EN UN COMBOBOX ------------------------------

Public Sub CargarRecetasEnCombo(ByRef comboBox As Object)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim fila As ListRow
    
    Set ws = ThisWorkbook.Sheets(HOJA_RECETAS)
    Set tbl = ws.ListObjects(TBL_RECETAS)
    
    comboBox.Clear
    comboBox.AddItem ""
    
    For Each fila In tbl.ListRows
        comboBox.AddItem fila.Range.Cells(1, COL_REC_NOMBRE).Value
    Next fila
End Sub

' --- CARGAR INGREDIENTES EN UN COMBOBOX ------------------------------------

Public Sub CargarIngredientesEnCombo(ByRef comboBox As Object)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim fila As ListRow
    
    Set ws = ThisWorkbook.Sheets(HOJA_INGREDIENTES)
    Set tbl = ws.ListObjects(TBL_INGREDIENTES)
    
    comboBox.Clear
    comboBox.AddItem ""
    
    For Each fila In tbl.ListRows
        comboBox.AddItem fila.Range.Cells(1, COL_ING_NOMBRE).Value
    Next fila
End Sub

' --- CARGAR UNIDADES SEGÚN TIPO DE INGREDIENTE -----------------------------

Public Sub CargarUnidadesPorIngrediente(ByRef comboBox As Object, _
                                         ByVal nombreIngrediente As String)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim fila As ListRow
    Dim unidadBase As String
    
    comboBox.Clear
    
    If nombreIngrediente = "" Then Exit Sub
    
    Set ws = ThisWorkbook.Sheets(HOJA_INGREDIENTES)
    Set tbl = ws.ListObjects(TBL_INGREDIENTES)
    
    ' Buscar unidad base
    For Each fila In tbl.ListRows
        If fila.Range.Cells(1, COL_ING_NOMBRE).Value = nombreIngrediente Then
            unidadBase = fila.Range.Cells(1, COL_ING_UNIDAD_BASE).Value
            Exit For
        End If
    Next fila
    
    If unidadBase = "" Then Exit Sub
    
    ' Cargar unidades según tipo
    Select Case LCase(unidadBase)
        Case "g", "kg", "mg", "lb", "oz"
            comboBox.AddItem "g"
            comboBox.AddItem "kg"
            comboBox.AddItem "mg"
            comboBox.AddItem "Lb"
            comboBox.AddItem "oz"
        Case "ml", "l"
            comboBox.AddItem "mL"
            comboBox.AddItem "L"
        Case "u"
            comboBox.AddItem "u"
    End Select
End Sub

' --- OBTENER DATO DE UNA RECETA --------------------------------------------

Public Function ObtenerDatoReceta(ByVal nombreReceta As String, _
                                   ByVal columna As Long) As Variant
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim fila As ListRow
    
    Set ws = ThisWorkbook.Sheets(HOJA_RECETAS)
    Set tbl = ws.ListObjects(TBL_RECETAS)
    
    ObtenerDatoReceta = 0
    
    For Each fila In tbl.ListRows
        If fila.Range.Cells(1, COL_REC_NOMBRE).Value = nombreReceta Then
            ObtenerDatoReceta = fila.Range.Cells(1, columna).Value
            Exit Function
        End If
    Next fila
End Function

' --- AGREGAR UN PRODUCTO A TABLA PRODUCTOS_UNICOS --------------------------

Public Sub AgregarProductoUnico(ByVal nombreProducto As String)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim newRow As ListRow
    Dim fila As ListRow
    
    Set ws = ThisWorkbook.Sheets(HOJA_FECHAS)
    Set tbl = ws.ListObjects(TBL_PRODUCTOS_UNICOS)
    
    ' Verificar si ya existe
    For Each fila In tbl.ListRows
        If UCase(Trim(fila.Range(1).Value)) = UCase(Trim(nombreProducto)) Then
            Exit Sub
        End If
    Next fila
    
    ws.Unprotect password:=PASS_HOJA
    
    Set newRow = tbl.ListRows.Add
    newRow.Range(1).Value = nombreProducto
    tbl.Sort.Apply
    
    ws.Protect password:=PASS_HOJA, UserInterfaceOnly:=True
End Sub

' --- OBTENER RECETAS QUE USAN UN INGREDIENTE ESPECÍFICO ----------------------
' Devuelve una colección con los nombres de todas las recetas que usan
' el ingrediente indicado. Sin duplicados.

' --- CREAR UNA NUEVA RECETA EN LA TABLA --------------------------------------
' Crea el registro en TBL_RECETAS con todos los campos calculados.
' Devuelve True si se creó exitosamente.

Public Function AgregarReceta(ByVal nombre As String, _
                               ByVal cantidadLote As Double, _
                               ByVal precioUnidad As Double, _
                               ByVal inversionLote As Double, _
                               ByVal inversionUnidad As Double) As Boolean
    On Error GoTo ErrorHandler
    
    Dim wsRec As Worksheet
    Dim tblRecetas As ListObject
    Dim newRow As ListRow
    Dim precioLote As Double
    Dim gananciaLote As Double
    Dim gananciaUnidad As Double
    
    ' Calcular valores derivados
    precioLote = precioUnidad * cantidadLote
    
    gananciaLote = precioLote - inversionLote
    gananciaUnidad = precioUnidad - inversionUnidad
    
    Set wsRec = ThisWorkbook.Sheets(HOJA_RECETAS)
    Set tblRecetas = wsRec.ListObjects(TBL_RECETAS)
    
    modSeguridad.DesprotegerHoja wsRec
    
    Set newRow = tblRecetas.ListRows.Add
    With newRow
        .Range.Cells(1, COL_REC_NOMBRE).Value = nombre
        .Range.Cells(1, COL_REC_LOTE).Value = cantidadLote
        .Range.Cells(1, COL_REC_INV_LOTE).Value = inversionLote
        .Range.Cells(1, COL_REC_PRECIO_LOTE).Value = precioLote
        .Range.Cells(1, COL_REC_GAN_LOTE).Value = gananciaLote
        .Range.Cells(1, COL_REC_INV_UNIDAD).Value = inversionUnidad
        .Range.Cells(1, COL_REC_PRECIO_UNID).Value = precioUnidad
        .Range.Cells(1, COL_REC_GAN_UNIDAD).Value = gananciaUnidad
    End With
    
    tblRecetas.Sort.Apply
    modSeguridad.ProtegerHoja wsRec
    
    AgregarReceta = True
    Exit Function
    
ErrorHandler:
    MsgBox "Error al crear la receta." & vbCrLf & _
           "Detalle: " & Err.Description, vbExclamation
    AgregarReceta = False
End Function

' --- ELIMINAR UNA RECETA DE TBL_RECETAS --------------------------------------

Public Sub EliminarReceta(ByVal nombreReceta As String)
    Dim wsRec As Worksheet
    Dim tblRecetas As ListObject
    Dim j As Long
    Dim filaTabla As ListRow
    
    Set wsRec = ThisWorkbook.Sheets(HOJA_RECETAS)
    Set tblRecetas = wsRec.ListObjects(TBL_RECETAS)
    
    modSeguridad.DesprotegerHoja wsRec
    
    j = 1
    Do While j <= tblRecetas.ListRows.Count
        Set filaTabla = tblRecetas.ListRows(j)
        If filaTabla.Range(1, COL_REC_NOMBRE).Value = nombreReceta Then
            filaTabla.Delete
            Exit Do
        Else
            j = j + 1
        End If
    Loop
    
    tblRecetas.Sort.Apply
    modSeguridad.ProtegerHoja wsRec
End Sub

' --- ELIMINAR TODOS LOS INGREDIENTES DE UNA RECETA ---------------------------

Public Sub EliminarIngredientesReceta(ByVal nombreReceta As String)
    Dim wsIngRec As Worksheet
    Dim tblIngRec As ListObject
    Dim j As Long
    Dim filaTabla As ListRow
    
    Set wsIngRec = ThisWorkbook.Sheets(HOJA_INGR_RECETA)
    Set tblIngRec = wsIngRec.ListObjects(TBL_INGREDIENTES_RECETA)
    
    modSeguridad.DesprotegerHoja wsIngRec
    
    j = 1
    Do While j <= tblIngRec.ListRows.Count
        Set filaTabla = tblIngRec.ListRows(j)
        If filaTabla.Range(1, COL_IR_RECETA).Value = nombreReceta Then
            filaTabla.Delete
        Else
            j = j + 1
        End If
    Loop
    
    tblIngRec.Sort.Apply
    modSeguridad.ProtegerHoja wsIngRec
End Sub

' --- ELIMINAR RECETA + INGREDIENTES (combinado) ------------------------------

Public Sub EliminarRecetaYIngredientes(ByVal nombreReceta As String)
    Call EliminarIngredientesReceta(nombreReceta)
    Call EliminarReceta(nombreReceta)
End Sub

' --- AGREGAR UN INGREDIENTE A UNA RECETA ------------------------------------

Public Sub AgregarIngredienteReceta(ByVal nombreReceta As String, _
                                     ByVal nombreIngrediente As String, _
                                     ByVal cantidad As Double, _
                                     ByVal unidad As String)
    Dim wsIngRec As Worksheet
    Dim tblIngRec As ListObject
    Dim newRow As ListRow
    
    Set wsIngRec = ThisWorkbook.Sheets(HOJA_INGR_RECETA)
    Set tblIngRec = wsIngRec.ListObjects(TBL_INGREDIENTES_RECETA)
    
    modSeguridad.DesprotegerHoja wsIngRec
    
    Set newRow = tblIngRec.ListRows.Add
    With newRow
        .Range.Cells(1, COL_IR_RECETA).Value = nombreReceta
        .Range.Cells(1, COL_IR_INGREDIENTE).Value = nombreIngrediente
        .Range.Cells(1, COL_IR_CANTIDAD).Value = cantidad
        .Range.Cells(1, COL_IR_UNIDAD).Value = unidad
    End With
    
    tblIngRec.Sort.Apply
    modSeguridad.ProtegerHoja wsIngRec
End Sub

' --- CARGAR INGREDIENTES DE UNA RECETA (devuelve array 2D) -------------------
' Cada fila: (nombreIngrediente, cantidad, unidad)

Public Function CargarIngredientesReceta(ByVal nombreReceta As String) As Variant
    Dim wsIngRec As Worksheet
    Dim tblIngRec As ListObject
    Dim fila As ListRow
    Dim resultados() As Variant
    Dim i As Long
    
    Set wsIngRec = ThisWorkbook.Sheets(HOJA_INGR_RECETA)
    Set tblIngRec = wsIngRec.ListObjects(TBL_INGREDIENTES_RECETA)
    
    ' Contar ingredientes
    i = 0
    For Each fila In tblIngRec.ListRows
        If fila.Range.Cells(1, COL_IR_RECETA).Value = nombreReceta Then
            i = i + 1
        End If
    Next fila
    
    If i = 0 Then
        CargarIngredientesReceta = Array()
        Exit Function
    End If
    
    ReDim resultados(1 To i, 1 To 3)
    i = 1
    
    For Each fila In tblIngRec.ListRows
        If fila.Range.Cells(1, COL_IR_RECETA).Value = nombreReceta Then
            resultados(i, 1) = fila.Range.Cells(1, COL_IR_INGREDIENTE).Value
            resultados(i, 2) = fila.Range.Cells(1, COL_IR_CANTIDAD).Value
            resultados(i, 3) = fila.Range.Cells(1, COL_IR_UNIDAD).Value
            i = i + 1
        End If
    Next fila
    
    CargarIngredientesReceta = resultados
End Function

Public Function ObtenerRecetasPorIngrediente(ByVal ingrediente As String) As Collection
    Dim wsIngRec As Worksheet
    Dim tblIngRec As ListObject
    Dim fila As ListRow
    Dim recetas As New Collection
    Dim recetaNombre As String

    Set wsIngRec = ThisWorkbook.Sheets(HOJA_INGR_RECETA)
    Set tblIngRec = wsIngRec.ListObjects(TBL_INGREDIENTES_RECETA)

    For Each fila In tblIngRec.ListRows
        If fila.Range.Cells(1, COL_IR_INGREDIENTE).Value = ingrediente Then
            recetaNombre = fila.Range.Cells(1, COL_IR_RECETA).Value
            ' Agregar sin duplicados (usando key)
            On Error Resume Next
            recetas.Add recetaNombre, recetaNombre
            On Error GoTo 0
        End If
    Next

    Set ObtenerRecetasPorIngrediente = recetas
End Function

