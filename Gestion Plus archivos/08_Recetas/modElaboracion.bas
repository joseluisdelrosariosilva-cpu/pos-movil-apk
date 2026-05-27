Attribute VB_Name = "modElaboracion"
' ============================================================================
' MÓDULO: modElaboracion
' PROPÓSITO: Lógica de negocio para la elaboración de productos a partir de
' recetas. Valida stock, descuenta ingredientes y registra en inventario.
' ============================================================================

Option Explicit

' --- VALIDAR STOCK DE INGREDIENTES PARA UNA ELABORACIÓN --------------------
' Devuelve True si hay suficiente stock.
' En ingredientesFaltantes devuelve la lista de faltantes (si los hay).

Public Function ValidarStockElaboracion(ByVal nombreReceta As String, _
                                         ByVal numLotes As Double, _
                                         ByRef ingredientesFaltantes As String) As Boolean
    Dim wsIng As Worksheet
    Dim wsIngrRec As Worksheet
    Dim tblIng As ListObject
    Dim tblIngrRec As ListObject
    Dim filaIngrRec As ListRow
    Dim filaIngrediente As ListRow
    Dim cantidadNecesaria As Double
    Dim cantidadPorLote As Double
    Dim cantidadTotal As Double
    Dim stockActual As Double
    Dim encontrado As Boolean
    Dim suficiente As Boolean
    
    Set wsIng = ThisWorkbook.Sheets(HOJA_INGREDIENTES)
    Set wsIngrRec = ThisWorkbook.Sheets(HOJA_INGR_RECETA)
    Set tblIng = wsIng.ListObjects(TBL_INGREDIENTES)
    Set tblIngrRec = wsIngrRec.ListObjects(TBL_INGREDIENTES_RECETA)
    
    suficiente = True
    ingredientesFaltantes = ""
    
    ' Recorrer ingredientes de la receta
    For Each filaIngrRec In tblIngrRec.ListRows
        If filaIngrRec.Range.Cells(1, COL_IR_RECETA).Value = nombreReceta Then
            cantidadPorLote = CDbl(filaIngrRec.Range.Cells(1, COL_IR_CANTIDAD).Value)
            cantidadTotal = cantidadPorLote * numLotes
            
            ' Convertir cantidad de la receta a unidad base del ingrediente
            Dim unidadReceta As String
            Dim ingredienteNombre As String
            ingredienteNombre = filaIngrRec.Range.Cells(1, COL_IR_INGREDIENTE).Value
            unidadReceta = filaIngrRec.Range.Cells(1, COL_IR_UNIDAD).Value
            cantidadTotal = modConversor.ConvertirUnidad(ingredienteNombre, cantidadPorLote * numLotes, unidadReceta)
            
            ' Buscar en tabla de ingredientes
            encontrado = False
            For Each filaIngrediente In tblIng.ListRows
                If filaIngrediente.Range.Cells(1, COL_ING_NOMBRE).Value = ingredienteNombre Then
                    encontrado = True
                    stockActual = filaIngrediente.Range.Cells(1, COL_ING_STOCK).Value
                    
                    If stockActual < cantidadTotal Then
                        suficiente = False
                        ingredientesFaltantes = ingredientesFaltantes & _
                            "- " & ingredienteNombre & _
                            ": Necesario: " & Format(Round(cantidadTotal, 3), "General Number") & _
                            " " & filaIngrediente.Range.Cells(1, COL_ING_UNIDAD_BASE).Value & _
                            " | Disponible: " & Format(Round(stockActual, 3), "General Number") & _
                            " " & filaIngrediente.Range.Cells(1, COL_ING_UNIDAD_BASE).Value & vbCrLf
                    End If
                    Exit For
                End If
            Next filaIngrediente
            
            If Not encontrado Then
                suficiente = False
                ingredientesFaltantes = ingredientesFaltantes & _
                    "- " & filaIngrRec.Range.Cells(1, COL_IR_INGREDIENTE).Value & _
                    " (no encontrado en ingredientes)" & vbCrLf
            End If
        End If
    Next filaIngrRec
    
    ValidarStockElaboracion = suficiente
End Function

' --- DESCONTAR STOCK DE INGREDIENTES TRAS ELABORAR -------------------------

Public Sub DescontarStockIngredientes(ByVal nombreReceta As String, _
                                       ByVal numLotes As Double)
    Dim wsIng As Worksheet
    Dim wsIngrRec As Worksheet
    Dim tblIng As ListObject
    Dim tblIngrRec As ListObject
    Dim filaIngrRec As ListRow
    Dim filaIngrediente As ListRow
    Dim cantidadPorLote As Double
    Dim cantidadTotal As Double
    Dim stockActual As Double
    Dim fondoActual As Double
    Dim precioUnitario As Double
    
    Set wsIng = ThisWorkbook.Sheets(HOJA_INGREDIENTES)
    Set wsIngrRec = ThisWorkbook.Sheets(HOJA_INGR_RECETA)
    Set tblIng = wsIng.ListObjects(TBL_INGREDIENTES)
    Set tblIngrRec = wsIngrRec.ListObjects(TBL_INGREDIENTES_RECETA)
    
    wsIng.Unprotect password:=PASS_HOJA
    
    For Each filaIngrRec In tblIngrRec.ListRows
        If filaIngrRec.Range.Cells(1, COL_IR_RECETA).Value = nombreReceta Then
            cantidadPorLote = CDbl(filaIngrRec.Range.Cells(1, COL_IR_CANTIDAD).Value)
            
            ' Convertir a unidad base antes de descontar
            Dim ingredienteNombre As String
            Dim unidadReceta As String
            ingredienteNombre = filaIngrRec.Range.Cells(1, COL_IR_INGREDIENTE).Value
            unidadReceta = filaIngrRec.Range.Cells(1, COL_IR_UNIDAD).Value
            cantidadTotal = modConversor.ConvertirUnidad(ingredienteNombre, cantidadPorLote * numLotes, unidadReceta)
            
            For Each filaIngrediente In tblIng.ListRows
                If filaIngrediente.Range.Cells(1, COL_ING_NOMBRE).Value = ingredienteNombre Then
                    stockActual = filaIngrediente.Range.Cells(1, COL_ING_STOCK).Value
                    fondoActual = filaIngrediente.Range.Cells(1, COL_ING_FONDO).Value
                    precioUnitario = filaIngrediente.Range.Cells(1, COL_ING_PRECIO).Value
                    
                    filaIngrediente.Range.Cells(1, COL_ING_STOCK).Value = stockActual - cantidadTotal
                    filaIngrediente.Range.Cells(1, COL_ING_FONDO).Value = _
                        fondoActual - (cantidadTotal * precioUnitario)
                    Exit For
                End If
            Next filaIngrediente
        End If
    Next filaIngrRec
    
    wsIng.Protect password:=PASS_HOJA, UserInterfaceOnly:=True
End Sub

' --- REGISTRAR PRODUCTO ELABORADO EN INVENTARIO ----------------------------

Public Sub RegistrarProductoElaborado(ByVal nombreReceta As String, _
                                       ByVal cantidadElaborada As Double, _
                                       ByVal precioVenta As Double, _
                                       ByVal costoUnidad As Double, _
                                       Optional ByVal codigoExterno As String = "")
    Dim wsAlmacen As Worksheet
    Dim tbl As ListObject
    Dim newRow As ListRow
    Dim codigoProducto As String
    
    ' Usar código externo si se provee, o generar uno nuevo
    If codigoExterno = "" Then
        Call GenerarCodigoProducto
        codigoProducto = Codigo_Producto
    Else
        codigoProducto = codigoExterno
    End If
    
    Set wsAlmacen = ThisWorkbook.Sheets(HOJA_ALMACEN)
    Set tbl = wsAlmacen.ListObjects(TBL_INVENTARIO)
    
    wsAlmacen.Unprotect password:=PASS_HOJA
    
    Set newRow = tbl.ListRows.Add
    With newRow
        .Range.Cells(1, COL_INV_CODIGO).Value = codigoProducto
        .Range.Cells(1, COL_INV_NOMBRE).Value = nombreReceta
        .Range.Cells(1, COL_INV_FECHA).Value = Date
        .Range.Cells(1, COL_INV_CANT_INI).Value = cantidadElaborada
        .Range.Cells(1, COL_INV_CANT_ACT).Value = cantidadElaborada
        .Range.Cells(1, COL_INV_PRECIO_V).Value = precioVenta
        .Range.Cells(1, COL_INV_PRECIO_C).Value = costoUnidad
        .Range.Cells(1, COL_INV_FONDO).Value = costoUnidad * cantidadElaborada
        .Range.Cells(1, COL_INV_INGRESO).Value = precioVenta * cantidadElaborada
        .Range.Cells(1, COL_INV_GANANCIA).Value = (precioVenta - costoUnidad) * cantidadElaborada
    End With
    
    tbl.Sort.Apply
    wsAlmacen.Protect password:=PASS_HOJA, UserInterfaceOnly:=True
    
    ' Registrar en tablas auxiliares
    Call AgregarProductoUnico(nombreReceta)
    Call AgregarFecha(Date)
End Sub

' --- OBTENER INGREDIENTES DE UNA RECETA ------------------------------------
' Devuelve un array con los ingredientes (nombre, cantidad, unidad).

Public Function ObtenerIngredientesReceta(ByVal nombreReceta As String) As Variant
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim fila As ListRow
    Dim resultados() As Variant
    Dim i As Long
    
    Set ws = ThisWorkbook.Sheets(HOJA_INGR_RECETA)
    Set tbl = ws.ListObjects(TBL_INGREDIENTES_RECETA)
    
    ' Contar ingredientes
    i = 0
    For Each fila In tbl.ListRows
        If fila.Range.Cells(1, COL_IR_RECETA).Value = nombreReceta Then
            i = i + 1
        End If
    Next fila
    
    If i = 0 Then
        ObtenerIngredientesReceta = Array()
        Exit Function
    End If
    
    ReDim resultados(1 To i, 1 To 3)
    i = 1
    
    For Each fila In tbl.ListRows
        If fila.Range.Cells(1, COL_IR_RECETA).Value = nombreReceta Then
            resultados(i, 1) = fila.Range.Cells(1, COL_IR_INGREDIENTE).Value
            resultados(i, 2) = fila.Range.Cells(1, COL_IR_CANTIDAD).Value
            resultados(i, 3) = fila.Range.Cells(1, COL_IR_UNIDAD).Value
            i = i + 1
        End If
    Next fila
    
    ObtenerIngredientesReceta = resultados
End Function

' --- CALCULAR COSTO TOTAL DE UNA RECETA (por lote) --------------------------

Public Function CalcularCostoReceta(ByVal nombreReceta As String) As Double
    Dim wsIngRec As Worksheet
    Dim wsIng As Worksheet
    Dim tblIngRec As ListObject
    Dim tblIng As ListObject
    Dim fila As ListRow
    Dim ingrediente As String
    Dim cantidad As Double
    Dim unidad As String
    Dim cantidadBase As Double
    Dim precioUnitario As Double
    Dim costototal As Double

    Set wsIngRec = ThisWorkbook.Sheets(HOJA_INGR_RECETA)
    Set wsIng = ThisWorkbook.Sheets(HOJA_INGREDIENTES)
    Set tblIngRec = wsIngRec.ListObjects(TBL_INGREDIENTES_RECETA)
    Set tblIng = wsIng.ListObjects(TBL_INGREDIENTES)

    costototal = 0

    ' Recorrer ingredientes de la receta
    For Each fila In tblIngRec.ListRows
        If fila.Range.Cells(1, COL_IR_RECETA).Value = nombreReceta Then
            ingrediente = fila.Range.Cells(1, COL_IR_INGREDIENTE).Value
            cantidad = fila.Range.Cells(1, COL_IR_CANTIDAD).Value
            unidad = fila.Range.Cells(1, COL_IR_UNIDAD).Value

            ' Convertir la cantidad a unidad base
            cantidadBase = modConversor.ConvertirUnidad(ingrediente, cantidad, unidad)

            ' Obtener precio por unidad base del ingrediente
            precioUnitario = modCalculosInventario.ObtenerPrecioPorUnidadBase(ingrediente)

            ' Acumular costo
            costototal = costototal + (cantidadBase * precioUnitario)
        End If
    Next

    CalcularCostoReceta = costototal
End Function

' --- ACTUALIZAR COSTOS DE UNA RECETA (costo, ganancia por lote/unidad) ------

Public Sub ActualizarCostosReceta(ByVal nombreReceta As String)
    Dim wsRec As Worksheet
    Dim tblRecetas As ListObject
    Dim filaReceta As ListRow
    Dim cantidadLote As Double
    Dim precioLote As Double
    Dim precioUnidad As Double
    Dim costoTotal As Double
    Dim costoUnidad As Double
    Dim gananciaLote As Double
    Dim gananciaUnidad As Double

    Set wsRec = ThisWorkbook.Sheets(HOJA_RECETAS)
    Set tblRecetas = wsRec.ListObjects(TBL_RECETAS)

    ' Buscar la receta
    For Each filaReceta In tblRecetas.ListRows
        If filaReceta.Range.Cells(1, COL_REC_NOMBRE).Value = nombreReceta Then
            cantidadLote = filaReceta.Range.Cells(1, COL_REC_LOTE).Value
            precioLote = filaReceta.Range.Cells(1, COL_REC_PRECIO_LOTE).Value
            precioUnidad = filaReceta.Range.Cells(1, COL_REC_PRECIO_UNID).Value

            ' Recalcular
            costoTotal = CalcularCostoReceta(nombreReceta)
            costoUnidad = costoTotal / cantidadLote
            gananciaLote = precioLote - costoTotal
            gananciaUnidad = precioUnidad - costoUnidad

            ' Actualizar tabla
            filaReceta.Range.Cells(1, COL_REC_INV_LOTE).Value = costoTotal
            filaReceta.Range.Cells(1, COL_REC_GAN_LOTE).Value = gananciaLote
            filaReceta.Range.Cells(1, COL_REC_INV_UNIDAD).Value = costoUnidad
            filaReceta.Range.Cells(1, COL_REC_GAN_UNIDAD).Value = gananciaUnidad

            Exit For
        End If
    Next filaReceta
End Sub

' --- ACTUALIZAR COSTOS DE TODAS LAS RECETAS QUE USAN UN INGREDIENTE ----------
' Se llama cuando cambia el precio de un ingrediente.

Public Sub ActualizarRecetasPorIngrediente(ByVal nombreIngrediente As String)
    Dim recetasAfectadas As Collection
    Dim receta As Variant

    Set recetasAfectadas = modRecetas.ObtenerRecetasPorIngrediente(nombreIngrediente)

    If recetasAfectadas.Count > 0 Then
        For Each receta In recetasAfectadas
            Call ActualizarCostosReceta(CStr(receta))
        Next receta
    End If
End Sub
