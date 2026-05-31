VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Elaborar 
   Caption         =   "Elaborar Producto"
   ClientHeight    =   5240
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   15060
   OleObjectBlob   =   "Elaborar.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Elaborar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

' ============================================================================
' FORMULARIO: Elaborar
' PROPOSITO:  Interfaz para elaborar productos a partir de recetas.
'              Toda la logica de negocio delega en modulos reorganizados.
' ============================================================================

Private recetaSeleccionada As String
Private cantidadPorLote As Double
Private inversionPorLote As Double
Private precioPorUnidad As Double
Private productosCoincidentes() As Variant   ' Array 2D de productos en almacen
Private numProductosCoincidentes As Long      ' Cantidad de coincidencias

Private Sub CargarRecetas()
    Call modRecetas.CargarRecetasEnCombo(Me.Recetas)
End Sub

' --- ELABORAR PRODUCTO -------------------------------------------------------
Private Sub Btn_Active_Click()
    Dim numLotes As Double
    Dim ingredientesFaltantes As String
    
    ' Validar receta seleccionada
    If Me.Recetas.Value = "" Then
        MsgBox "Por favor, seleccione una receta.", vbExclamation
        Me.Recetas.SetFocus
        Exit Sub
    End If
    
    ' Validar numero de lotes
    If Me.Lotes.Value = "" Or Not IsNumeric(Me.Lotes.Value) Then
        MsgBox "Por favor, ingrese una cantidad valida de lotes.", vbExclamation
        Me.Lotes.SetFocus
        Exit Sub
    End If
    
    numLotes = CDbl(Me.Lotes.Value)
    
    If Me.Ingredientes.ListCount = 0 Then
        MsgBox "La receta seleccionada no tiene ingredientes.", vbExclamation
        Exit Sub
    End If
    
    ' PASO 1: VALIDAR STOCK
    Application.ScreenUpdating = False
    
    If Not modElaboracion.ValidarStockElaboracion(Me.Recetas.Value, numLotes, ingredientesFaltantes) Then
        Application.ScreenUpdating = True
        MsgBox "Stock insuficiente para los siguientes ingredientes:" & vbCrLf & vbCrLf & _
               ingredientesFaltantes & vbCrLf & _
               "Por favor, reponga el stock y vuelva a intentar.", vbExclamation
        Exit Sub
    End If
    
    ' PASO 2: DESCONTAR STOCK DE INGREDIENTES (comun a ambos modos)
    Call modElaboracion.DescontarStockIngredientes(Me.Recetas.Value, numLotes)
    
    ' PASO 3: BIFURCAR SEGUN MODO
    If Me.cboModo.Enabled And Me.cboModo.Value = "Abastecer" Then
        Call FlujoAbastecer(numLotes)
    Else
        Call FlujoNuevo(numLotes)
    End If
    
    Application.ScreenUpdating = True
    Unload Me
End Sub

' --- FLUJO: NUEVO PRODUCTO (comportamiento original) -------------------------
Private Sub FlujoNuevo(ByVal numLotes As Double)
    Dim codigoProducto As String
    Dim cantidadElaborada As Double
    Dim inversionTotal As Double
    Dim costoPorUnidad As Double
    Dim ingresoEsperado As Double
    Dim beneficioEsperado As Double
    
    ' Generar codigo
    Call modGeneradorCodigo.GenerarCodigoProducto
    codigoProducto = Codigo_Producto
    
    ' Registrar en inventario
    cantidadElaborada = CDbl(Me.cantidad.Value)
    inversionTotal = CDbl(Me.Inversion.Value)
    costoPorUnidad = inversionTotal / cantidadElaborada
    
    Call modElaboracion.RegistrarProductoElaborado( _
        Me.Recetas.Value, cantidadElaborada, _
        CDbl(Me.Precio_Venta.Value), costoPorUnidad, _
        codigoProducto)
    
    ' Mostrar mensaje de exito
    ingresoEsperado = CDbl(Me.Ingreso.Value)
    beneficioEsperado = CDbl(Me.Beneficio.Value)
    
    Dim resumen As String
    resumen = "ELABORACION COMPLETADA EXITOSAMENTE" & vbCrLf & vbCrLf & _
              "Receta: " & Me.Recetas.Value & vbCrLf & _
              "Lotes producidos: " & numLotes & vbCrLf & _
              "Unidades totales: " & cantidadElaborada & vbCrLf & _
              "Codigo del producto: " & codigoProducto & vbCrLf & _
              "Fecha de produccion: " & Format(Date, "dd/mm/yyyy") & vbCrLf & _
              "Inversion total: $" & modFormateo.FormatearNumero(inversionTotal) & vbCrLf & _
              "Ingreso esperado: $" & modFormateo.FormatearNumero(ingresoEsperado) & vbCrLf & _
              "Beneficio esperado: $" & modFormateo.FormatearNumero(beneficioEsperado) & vbCrLf & _
              "Costo por unidad: $" & modFormateo.FormatearNumero(costoPorUnidad) & vbCrLf & _
              "Precio de venta: $" & modFormateo.FormatearNumero(CDbl(Me.Precio_Venta.Value)) & vbCrLf & _
              "Ganancia por unidad: $" & modFormateo.FormatearNumero(CDbl(Me.Precio_Venta.Value) - costoPorUnidad)
    
    MsgBox resumen, vbInformation, "Elaboracion Exitosa"
End Sub

' --- FLUJO: ABASTECER PRODUCTO EXISTENTE -------------------------------------
Private Sub FlujoAbastecer(ByVal numLotes As Double)
    Dim cantidadElaborada As Double
    Dim inversionTotal As Double
    Dim costoPorUnidad As Double
    Dim indiceSeleccion As Long
    Dim codigoElegido As String
    Dim i As Long
    
    cantidadElaborada = CDbl(Me.cantidad.Value)
    inversionTotal = CDbl(Me.Inversion.Value)
    costoPorUnidad = inversionTotal / cantidadElaborada
    
    ' Si hay mas de un producto coincidente, pedir seleccion
    If numProductosCoincidentes > 1 Then
        Dim mensaje As String
        mensaje = "Se encontraron " & numProductosCoincidentes & " productos con el nombre """ & _
                  recetaSeleccionada & """:" & vbCrLf & vbCrLf
        
        For i = 1 To numProductosCoincidentes
            mensaje = mensaje & "  " & i & " - Codigo: " & productosCoincidentes(i, 1) & _
                      " | Precio: $" & modFormateo.FormatearNumero(CDbl(productosCoincidentes(i, 2))) & vbCrLf
        Next i
        
        mensaje = mensaje & vbCrLf & "Ingrese el numero del producto a abastecer:"
        
        Dim inputRespuesta As String
        inputRespuesta = InputBox(mensaje, "Seleccionar Producto", "1")
        
        If inputRespuesta = "" Then
            Application.ScreenUpdating = True
            Exit Sub
        End If
        
        If Not IsNumeric(inputRespuesta) Then
            MsgBox "Debe ingresar un numero valido.", vbExclamation
            Application.ScreenUpdating = True
            Exit Sub
        End If
        
        indiceSeleccion = CLng(inputRespuesta)
        
        If indiceSeleccion < 1 Or indiceSeleccion > numProductosCoincidentes Then
            MsgBox "Numero fuera de rango. Debe elegir entre 1 y " & numProductosCoincidentes & ".", vbExclamation
            Application.ScreenUpdating = True
            Exit Sub
        End If
        
        codigoElegido = productosCoincidentes(indiceSeleccion, 1)
    Else
        ' Un solo producto, usar directamente
        codigoElegido = productosCoincidentes(1, 1)
    End If
    
    ' Abastecer producto existente
    Call modElaboracion.AbastecerProductoElaborado( _
        codigoElegido, cantidadElaborada, _
        CDbl(Me.Precio_Venta.Value), costoPorUnidad)
    
    ' Mostrar mensaje de exito
    Dim resumen As String
    resumen = "ABASTECIMIENTO COMPLETADO EXITOSAMENTE" & vbCrLf & vbCrLf & _
              "Receta: " & Me.Recetas.Value & vbCrLf & _
              "Producto: " & codigoElegido & vbCrLf & _
              "Lotes producidos: " & numLotes & vbCrLf & _
              "Unidades agregadas: " & cantidadElaborada & vbCrLf & _
              "Fecha: " & Format(Date, "dd/mm/yyyy") & vbCrLf & _
              "Inversion total: $" & modFormateo.FormatearNumero(inversionTotal) & vbCrLf & _
              "Costo por unidad: $" & modFormateo.FormatearNumero(costoPorUnidad) & vbCrLf & _
              "Precio de venta actualizado: $" & modFormateo.FormatearNumero(CDbl(Me.Precio_Venta.Value))
    
    MsgBox resumen, vbInformation, "Abastecimiento Exitoso"
End Sub

' --- ACTUALIZAR ESTADO DEL COMBOBOX SEGUN ALMACEN ---------------------------
Private Sub ActualizarEstadoModoElaboracion()
    Dim resultados As Variant
    
    If recetaSeleccionada = "" Then
        Me.cboModo.Enabled = False
        Me.cboModo.Value = "Nuevo"
        numProductosCoincidentes = 0
        Erase productosCoincidentes
        Exit Sub
    End If
    
    resultados = modStock.BuscarProductosEnAlmacenPorNombre(recetaSeleccionada)
    
    On Error Resume Next
    numProductosCoincidentes = UBound(resultados) - LBound(resultados) + 1
    If Err.Number <> 0 Then numProductosCoincidentes = 0
    On Error GoTo 0
    
    If numProductosCoincidentes > 0 Then
        productosCoincidentes = resultados
        Me.cboModo.Enabled = True
        Me.cboModo.Value = "Abastecer"
    Else
        Erase productosCoincidentes
        Me.cboModo.Enabled = False
        Me.cboModo.Value = "Nuevo"
    End If
End Sub

Private Sub Btn_Reposo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.Btn_Active.ZOrder msoBringToFront
    Me.Txt_Reposo.ZOrder msoBringToFront
End Sub

Private Sub Lotes_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Call modValidacionNumerica.SoloNumeros(KeyAscii)
End Sub

' --- ACTUALIZAR CÁLCULOS AL CAMBIAR PRECIO DE VENTA -------------------------
Private Sub Precio_venta_Change()
    If Me.Recetas.Value <> "" And Me.Lotes.Value <> "" And IsNumeric(Me.Lotes.Value) Then
        Dim ingresoEsperado As Double
        Dim beneficioEsperado As Double
        
        If Me.Precio_Venta.Value <> "" And IsNumeric(Me.Precio_Venta.Value) And _
           Me.cantidad.Value <> "" And IsNumeric(Me.cantidad.Value) Then
            
            ingresoEsperado = CDbl(Me.Precio_Venta.Value) * CDbl(Me.cantidad.Value)
            Me.Ingreso.Value = modFormateo.FormatearNumero(ingresoEsperado)
            
            If Me.Inversion.Value <> "" And IsNumeric(Me.Inversion.Value) Then
                beneficioEsperado = ingresoEsperado - CDbl(Me.Inversion.Value)
                Me.Beneficio.Value = modFormateo.FormatearNumero(beneficioEsperado)
            End If
        Else
            Me.Ingreso.Value = ""
            Me.Beneficio.Value = ""
        End If
    End If
End Sub

' --- CARGAR DATOS DE LA RECETA SELECCIONADA ---------------------------------
Private Sub Recetas_Change()
    Dim wsRecetas As Worksheet
    Dim wsIngredientes As Worksheet
    Dim tblRecetas As ListObject
    Dim tblIngredientesReceta As ListObject
    Dim filaReceta As ListRow
    Dim filaIngrediente As ListRow
    
    ' Limpiar datos anteriores
    recetaSeleccionada = ""
    cantidadPorLote = 0
    inversionPorLote = 0
    precioPorUnidad = 0
    
    Me.Ingredientes.Clear
    Me.Lotes.Value = ""
    Me.cantidad.Value = ""
    Me.Inversion.Value = ""
    Me.Precio_Venta.Value = ""
    Me.Ingreso.Value = ""
    Me.Beneficio.Value = ""
    
    If Me.Recetas.Value = "" Then
        Call ActualizarEstadoModoElaboracion
        Exit Sub
    End If
    
    recetaSeleccionada = Me.Recetas.Value
    
    Set wsRecetas = ThisWorkbook.Sheets(HOJA_RECETAS)
    Set wsIngredientes = ThisWorkbook.Sheets(HOJA_INGR_RECETA)
    Set tblRecetas = wsRecetas.ListObjects(TBL_RECETAS)
    Set tblIngredientesReceta = wsIngredientes.ListObjects(TBL_INGREDIENTES_RECETA)
    
    ' Buscar la receta seleccionada
    For Each filaReceta In tblRecetas.ListRows
        If filaReceta.Range.Cells(1, 1).Value = recetaSeleccionada Then
            cantidadPorLote = filaReceta.Range.Cells(1, COL_REC_LOTE).Value
            inversionPorLote = filaReceta.Range.Cells(1, COL_REC_INV_LOTE).Value
            precioPorUnidad = filaReceta.Range.Cells(1, COL_REC_PRECIO_UNID).Value
            Me.Precio_Venta.Value = precioPorUnidad
            Exit For
        End If
    Next filaReceta
    
    ' Cargar ingredientes
    For Each filaIngrediente In tblIngredientesReceta.ListRows
        If filaIngrediente.Range.Cells(1, COL_IR_RECETA).Value = recetaSeleccionada Then
            Me.Ingredientes.AddItem
            With Me.Ingredientes
                .List(.ListCount - 1, 0) = filaIngrediente.Range.Cells(1, COL_IR_INGREDIENTE).Value
                .List(.ListCount - 1, 1) = filaIngrediente.Range.Cells(1, COL_IR_CANTIDAD).Value
                .List(.ListCount - 1, 2) = filaIngrediente.Range.Cells(1, COL_IR_UNIDAD).Value
            End With
        End If
    Next filaIngrediente
    
    ' Inicializar con 1 lote por defecto
    If Me.Ingredientes.ListCount > 0 And Me.Lotes.Value = "" Then
        Me.Lotes.Value = 1
        ActualizarCalculos
    End If
    
    ' Actualizar modo de elaboracion segun existencias en almacen
    Call ActualizarEstadoModoElaboracion
End Sub

Private Sub Lotes_Change()
    If Me.Lotes.Value = "" Then
        Me.cantidad.Value = ""
        Me.Inversion.Value = ""
        ActualizarListaIngredientes 0
        Exit Sub
    End If
    If IsNumeric(Me.Lotes.Value) Then
        ActualizarCalculos
    End If
End Sub

Private Sub ActualizarCalculos()
    Dim numLotes As Double
    Dim ingresoEsperado As Double
    Dim beneficioEsperado As Double
    
    If Me.Lotes.Value = "" Then
        numLotes = 0
    Else
        numLotes = CDbl(Me.Lotes.Value)
    End If
    
    Me.cantidad.Value = cantidadPorLote * numLotes
    Me.Inversion.Value = modFormateo.FormatearNumero(inversionPorLote * numLotes)
    
    If Me.Precio_Venta.Value <> "" And IsNumeric(Me.Precio_Venta.Value) Then
        ingresoEsperado = CDbl(Me.Precio_Venta.Value) * CDbl(Me.cantidad.Value)
        Me.Ingreso.Value = modFormateo.FormatearNumero(ingresoEsperado)
    Else
        Me.Ingreso.Value = ""
    End If
    
    If Me.Inversion.Value <> "" And IsNumeric(Me.Inversion.Value) And _
       Me.Ingreso.Value <> "" And IsNumeric(Me.Ingreso.Value) Then
        beneficioEsperado = CDbl(Me.Ingreso.Value) - CDbl(Me.Inversion.Value)
        Me.Beneficio.Value = modFormateo.FormatearNumero(beneficioEsperado)
    Else
        Me.Beneficio.Value = ""
    End If
    
    ActualizarListaIngredientes numLotes
End Sub

Private Sub ActualizarListaIngredientes(numLotes As Double)
    Dim wsIngredientes As Worksheet
    Dim tblIngredientesReceta As ListObject
    Dim filaIngrediente As ListRow
    Dim i As Long
    Dim ingredienteNombre As String
    Dim cantidadOriginal As Double
    Dim encontrado As Boolean
    
    Set wsIngredientes = ThisWorkbook.Sheets(HOJA_INGR_RECETA)
    Set tblIngredientesReceta = wsIngredientes.ListObjects(TBL_INGREDIENTES_RECETA)
    
    For i = 0 To Me.Ingredientes.ListCount - 1
        ingredienteNombre = Me.Ingredientes.List(i, 0)
        encontrado = False
        
        For Each filaIngrediente In tblIngredientesReceta.ListRows
            If filaIngrediente.Range.Cells(1, COL_IR_RECETA).Value = recetaSeleccionada And _
               filaIngrediente.Range.Cells(1, COL_IR_INGREDIENTE).Value = ingredienteNombre Then
                cantidadOriginal = CDbl(filaIngrediente.Range.Cells(1, COL_IR_CANTIDAD).Value)
                Me.Ingredientes.List(i, 1) = cantidadOriginal * numLotes
                encontrado = True
                Exit For
            End If
        Next filaIngrediente
    Next i
End Sub

Private Sub Lotes_AfterUpdate()
    ActualizarCalculos
End Sub

Private Sub Txt_Reposo_Click()
    Btn_Active_Click
End Sub

Private Sub UserForm_Initialize()
    CargarRecetas
    
    Me.Lotes.Value = ""
    Me.cantidad.Value = ""
    Me.Inversion.Value = ""
    Me.Precio_Venta.Value = ""
    Me.Ingreso.Value = ""
    Me.Beneficio.Value = ""
    
    With Me.Ingredientes
        .ColumnCount = 3
        .ColumnWidths = "120;80;80"
    End With
    
    ' Configurar ComboBox de modo (Nuevo / Abastecer)
    With Me.cboModo
        .Clear
        .AddItem "Nuevo"
        .AddItem "Abastecer"
        .Value = "Nuevo"
        .Enabled = False
    End With
End Sub

Private Sub LimpiarFormularioElaborar()
    Me.Recetas.Value = ""
    Me.Lotes.Value = ""
    Me.cantidad.Value = ""
    Me.Inversion.Value = ""
    Me.Precio_Venta.Value = ""
    Me.Ingreso.Value = ""
    Me.Beneficio.Value = ""
    Me.Ingredientes.Clear
    Me.Recetas.SetFocus
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.Btn_Reposo.ZOrder msoBringToFront
    Me.Txt_Reposo.ZOrder msoBringToFront
End Sub
