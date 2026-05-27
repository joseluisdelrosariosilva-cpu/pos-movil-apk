VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Ingr_Receta 
   Caption         =   "Ingresar Receta"
   ClientHeight    =   6620
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   14820
   OleObjectBlob   =   "Ingr_Receta.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Ingr_Receta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ============================================================================
' FORMULARIO: Ingr_Receta
' PROPÓSITO:  Interfaz para crear y modificar recetas con hasta 10
'              ingredientes. Toda la lógica de negocio delega en módulos.
' ============================================================================

Private recetaSeleccionada As String

' --- BOTÓN PRINCIPAL (Agregar / Modificar) ------------------------------------
Private Sub Btn_Active_Click()
    If modoModificacion Then
        ModificarRecetaExistente
    Else
        AgregarNuevaReceta
    End If
End Sub

Private Sub Agregar_Click()
    Btn_Active_Click
End Sub

' --- AGREGAR NUEVA RECETA ----------------------------------------------------
Private Sub AgregarNuevaReceta()
    Dim recetaNombre As String
    Dim cantidadLote As Double
    Dim precioUnidad As Double
    Dim Inversion_lote As Double
    Dim Inversion_unidad As Double
    Dim i As Integer
    
    ' Validar datos básicos
    recetaNombre = Trim(Me.Txt_Receta.Value)
    If recetaNombre = "" Then
        MsgBox "Por favor, ingrese un nombre para la receta.", vbExclamation
        Me.Txt_Receta.SetFocus: Exit Sub
    End If
    
    If modRecetas.RecetaExiste(recetaNombre) Then
        MsgBox "Ya existe una receta con este nombre.", vbExclamation
        Me.Txt_Receta.SetFocus: Exit Sub
    End If
    
    If Me.Txt_Cantidad.Value = "" Or Not IsNumeric(Me.Txt_Cantidad.Value) Then
        MsgBox "Por favor, ingrese una cantidad válida para el lote.", vbExclamation
        Me.Txt_Cantidad.SetFocus: Exit Sub
    End If
    cantidadLote = CDbl(Me.Txt_Cantidad.Value)
    
    If Me.Txt_Precio_u.Value = "" Or Not IsNumeric(Me.Txt_Precio_u.Value) Then
        MsgBox "Por favor, ingrese un precio de venta por unidad válido.", vbExclamation
        Me.Txt_Precio_u.SetFocus: Exit Sub
    End If
    precioUnidad = CDbl(Me.Txt_Precio_u.Value)
    
    ' Verificar al menos un ingrediente
    If Not HayIngredientesValidos() Then
        MsgBox "Por favor, agregue al menos un ingrediente a la receta.", vbExclamation
        Exit Sub
    End If
    
    Inversion_lote = CDbl(Me.Txt_Costo.Value)
    Inversion_unidad = CDbl(Me.Txt_Costo_u.Value)
    
    ' Crear receta en BD
    If Not modRecetas.AgregarReceta(recetaNombre, cantidadLote, precioUnidad, Inversion_lote, Inversion_unidad) Then
        Exit Sub
    End If
    
    ' Agregar ingredientes
    For i = 1 To 10
        Dim cbIng As MSForms.comboBox
        Dim txtCant As MSForms.TextBox
        Dim cbUnd As MSForms.comboBox
        
        Set cbIng = Me.Controls("Cmb" & i)
        Set txtCant = Me.Controls("TextBox" & i)
        Set cbUnd = Me.Controls("Cmb_" & i)
        
        If cbIng.Value <> "" And cbUnd.Value <> "" And txtCant.Value <> "" Then
            If IsNumeric(txtCant.Value) Then
                Call modRecetas.AgregarIngredienteReceta( _
                    recetaNombre, cbIng.Value, CDbl(txtCant.Value), cbUnd.Value)
            End If
        End If
    Next i
    
    MsgBox "Receta '" & recetaNombre & "' guardada exitosamente.", vbInformation
    Unload Me
End Sub

' --- MODIFICAR RECETA EXISTENTE ----------------------------------------------
Private Sub ModificarRecetaExistente()
    Dim recetaNombre As String
    Dim cantidadLote As Double
    Dim precioUnidad As Double
    Dim Inversion_lote As Double
    Dim Inversion_unidad As Double
    Dim i As Integer
    
    ' Validar datos básicos
    recetaNombre = Trim(Me.Txt_Receta.Value)
    If recetaNombre = "" Then
        MsgBox "Por favor, ingrese un nombre para la receta.", vbExclamation
        Me.Txt_Receta.SetFocus: Exit Sub
    End If
    
    If Me.Txt_Cantidad.Value = "" Or Not IsNumeric(Me.Txt_Cantidad.Value) Then
        MsgBox "Por favor, ingrese una cantidad válida para el lote.", vbExclamation
        Me.Txt_Cantidad.SetFocus: Exit Sub
    End If
    cantidadLote = CDbl(Me.Txt_Cantidad.Value)
    
    If Me.Txt_Precio_u.Value = "" Or Not IsNumeric(Me.Txt_Precio_u.Value) Then
        MsgBox "Por favor, ingrese un precio de venta por unidad válido.", vbExclamation
        Me.Txt_Precio_u.SetFocus: Exit Sub
    End If
    precioUnidad = CDbl(Me.Txt_Precio_u.Value)
    
    If Not HayIngredientesValidos() Then
        MsgBox "Por favor, agregue al menos un ingrediente a la receta.", vbExclamation
        Exit Sub
    End If
    
    Inversion_lote = CDbl(Me.Txt_Costo.Value)
    Inversion_unidad = CDbl(Me.Txt_Costo_u.Value)
    
    ' Eliminar receta anterior
    Call modRecetas.EliminarRecetaYIngredientes(recetaAModificar)
    
    ' Crear la nueva receta
    If Not modRecetas.AgregarReceta(recetaNombre, cantidadLote, precioUnidad, Inversion_lote, Inversion_unidad) Then
        Exit Sub
    End If
    
    ' Agregar ingredientes
    For i = 1 To 10
        Dim cbIng As MSForms.comboBox
        Dim txtCant As MSForms.TextBox
        Dim cbUnd As MSForms.comboBox
        
        Set cbIng = Me.Controls("Cmb" & i)
        Set txtCant = Me.Controls("TextBox" & i)
        Set cbUnd = Me.Controls("Cmb_" & i)
        
        If cbIng.Value <> "" And cbUnd.Value <> "" And txtCant.Value <> "" Then
            If IsNumeric(txtCant.Value) Then
                Call modRecetas.AgregarIngredienteReceta( _
                    recetaNombre, cbIng.Value, CDbl(txtCant.Value), cbUnd.Value)
            End If
        End If
    Next i
    
    MsgBox "Receta '" & recetaNombre & "' modificada exitosamente.", vbInformation
    Unload Me
End Sub

' --- VERIFICAR SI HAY AL MENOS UN INGREDIENTE VÁLIDO -------------------------
Private Function HayIngredientesValidos() As Boolean
    Dim i As Integer
    
    HayIngredientesValidos = False
    
    For i = 1 To 10
        Dim cbIng As MSForms.comboBox
        Dim txtCant As MSForms.TextBox
        Dim cbUnd As MSForms.comboBox
        
        Set cbIng = Me.Controls("Cmb" & i)
        Set txtCant = Me.Controls("TextBox" & i)
        Set cbUnd = Me.Controls("Cmb_" & i)
        
        If cbIng.Value <> "" And cbUnd.Value <> "" And txtCant.Value <> "" Then
            If IsNumeric(txtCant.Value) Then
                HayIngredientesValidos = True
                Exit Function
            End If
        End If
    Next i
End Function

' --- LIMPIAR FORMULARIO ------------------------------------------------------
Private Sub LimpiarFormulario()
    Dim i As Integer
    
    Me.Txt_Receta.Value = ""
    Me.Txt_Cantidad.Value = ""
    Me.Txt_Costo.Value = ""
    Me.Txt_Costo_u.Value = ""
    Me.Txt_Precio_u.Value = ""
    Me.Txt_Precio.Value = ""
    
    For i = 1 To 10
        Me.Controls("Cmb" & i).Value = ""
        Me.Controls("TextBox" & i).Value = ""
        Me.Controls("Cmb_" & i).Value = ""
    Next i
    
    Me.Txt_Receta.SetFocus
End Sub

' --- EVENTOS KeyPress (todos delegan a modValidacionNumerica) ----------------
Private Sub TextBox1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger): Call modValidacionNumerica.SoloNumeros(KeyAscii): End Sub
Private Sub TextBox2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger): Call modValidacionNumerica.SoloNumeros(KeyAscii): End Sub
Private Sub TextBox3_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger): Call modValidacionNumerica.SoloNumeros(KeyAscii): End Sub
Private Sub TextBox4_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger): Call modValidacionNumerica.SoloNumeros(KeyAscii): End Sub
Private Sub TextBox5_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger): Call modValidacionNumerica.SoloNumeros(KeyAscii): End Sub
Private Sub TextBox6_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger): Call modValidacionNumerica.SoloNumeros(KeyAscii): End Sub
Private Sub TextBox7_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger): Call modValidacionNumerica.SoloNumeros(KeyAscii): End Sub
Private Sub TextBox8_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger): Call modValidacionNumerica.SoloNumeros(KeyAscii): End Sub
Private Sub TextBox9_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger): Call modValidacionNumerica.SoloNumeros(KeyAscii): End Sub
Private Sub TextBox10_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger): Call modValidacionNumerica.SoloNumeros(KeyAscii): End Sub
Private Sub Txt_Cantidad_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger): Call modValidacionNumerica.SoloNumeros(KeyAscii): End Sub
Private Sub Txt_Precio_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger): Call modValidacionNumerica.SoloNumeros(KeyAscii): End Sub
Private Sub Txt_Precio_u_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger): Call modValidacionNumerica.SoloNumeros(KeyAscii): End Sub

' --- EVENTO Change: recalcular costo al modificar cantidades -----------------
Private Sub TextBox1_Change(): CalcularCostoTotal: End Sub
Private Sub TextBox2_Change(): CalcularCostoTotal: End Sub
Private Sub TextBox3_Change(): CalcularCostoTotal: End Sub
Private Sub TextBox4_Change(): CalcularCostoTotal: End Sub
Private Sub TextBox5_Change(): CalcularCostoTotal: End Sub
Private Sub TextBox6_Change(): CalcularCostoTotal: End Sub
Private Sub TextBox7_Change(): CalcularCostoTotal: End Sub
Private Sub TextBox8_Change(): CalcularCostoTotal: End Sub
Private Sub TextBox9_Change(): CalcularCostoTotal: End Sub
Private Sub TextBox10_Change(): CalcularCostoTotal: End Sub

' --- EVENTO Change: actualizar unidades al cambiar ingrediente ---------------
Private Sub Cmb1_Change(): ActualizarUnidades 1: CalcularCostoTotal: End Sub
Private Sub Cmb2_Change(): ActualizarUnidades 2: CalcularCostoTotal: End Sub
Private Sub Cmb3_Change(): ActualizarUnidades 3: CalcularCostoTotal: End Sub
Private Sub Cmb4_Change(): ActualizarUnidades 4: CalcularCostoTotal: End Sub
Private Sub Cmb5_Change(): ActualizarUnidades 5: CalcularCostoTotal: End Sub
Private Sub Cmb6_Change(): ActualizarUnidades 6: CalcularCostoTotal: End Sub
Private Sub Cmb7_Change(): ActualizarUnidades 7: CalcularCostoTotal: End Sub
Private Sub Cmb8_Change(): ActualizarUnidades 8: CalcularCostoTotal: End Sub
Private Sub Cmb9_Change(): ActualizarUnidades 9: CalcularCostoTotal: End Sub
Private Sub Cmb10_Change(): ActualizarUnidades 10: CalcularCostoTotal: End Sub

' --- EVENTO Change: recalcular costo al cambiar unidad -----------------------
Private Sub Cmb_1_Change(): CalcularCostoTotal: End Sub
Private Sub Cmb_2_Change(): CalcularCostoTotal: End Sub
Private Sub Cmb_3_Change(): CalcularCostoTotal: End Sub
Private Sub Cmb_4_Change(): CalcularCostoTotal: End Sub
Private Sub Cmb_5_Change(): CalcularCostoTotal: End Sub
Private Sub Cmb_6_Change(): CalcularCostoTotal: End Sub
Private Sub Cmb_7_Change(): CalcularCostoTotal: End Sub
Private Sub Cmb_8_Change(): CalcularCostoTotal: End Sub
Private Sub Cmb_9_Change(): CalcularCostoTotal: End Sub
Private Sub Cmb_10_Change(): CalcularCostoTotal: End Sub

' --- ACTUALIZAR UNIDADES SEGÚN INGREDIENTE SELECCIONADO ----------------------
Private Sub ActualizarUnidades(numeroModulo As Integer)
    Dim cbIngrediente As MSForms.comboBox
    Dim cbUnidad As MSForms.comboBox
    Dim txtCantidad As MSForms.TextBox
    
    Set cbIngrediente = Me.Controls("Cmb" & numeroModulo)
    Set cbUnidad = Me.Controls("Cmb_" & numeroModulo)
    Set txtCantidad = Me.Controls("TextBox" & numeroModulo)
    
    txtCantidad.Value = ""
    Call modRecetas.CargarUnidadesPorIngrediente(cbUnidad, cbIngrediente.Value)
End Sub

' --- CALCULAR COSTO TOTAL DE LA RECETA (en tiempo real) ----------------------
Private Sub CalcularCostoTotal()
    Dim totalCosto As Double
    Dim i As Integer
    
    totalCosto = 0
    
    For i = 1 To 10
        Dim cbIng As MSForms.comboBox
        Dim txtCant As MSForms.TextBox
        Dim cbUnd As MSForms.comboBox
        
        Set cbIng = Me.Controls("Cmb" & i)
        Set txtCant = Me.Controls("TextBox" & i)
        Set cbUnd = Me.Controls("Cmb_" & i)
        
        If cbIng.Value <> "" And cbUnd.Value <> "" Then
            If IsNumeric(txtCant.Value) And txtCant.Value <> "" Then
                Dim cantidadBase As Double
                Dim precioUnitario As Double
                
                cantidadBase = modConversor.ConvertirUnidad(cbIng.Value, CDbl(txtCant.Value), cbUnd.Value)
                precioUnitario = modCalculosInventario.ObtenerPrecioPorUnidadBase(cbIng.Value)
                totalCosto = totalCosto + (cantidadBase * precioUnitario)
            End If
        End If
    Next i
    
    Me.Txt_Costo.Value = modFormateo.FormatearNumero(totalCosto)
    
    ' Calcular costo por unidad
    If Me.Txt_Cantidad.Value <> "" And IsNumeric(Me.Txt_Cantidad.Value) Then
        Dim cantLote As Double
        cantLote = CDbl(Me.Txt_Cantidad.Value)
        If cantLote > 0 Then
            Me.Txt_Costo_u.Value = modFormateo.FormatearNumero(totalCosto / cantLote)
        End If
    End If
End Sub

' --- EVENTO Change: recalcular precio del lote al cambiar cantidad -----------
Private Sub Txt_Cantidad_Change()
    If Me.Txt_Cantidad.Value <> "" And IsNumeric(Me.Txt_Cantidad.Value) Then
        Dim cantidad As Double
        cantidad = CDbl(Me.Txt_Cantidad.Value)
        
        If Me.Txt_Precio_u.Value <> "" And IsNumeric(Me.Txt_Precio_u.Value) Then
            Dim precioUnitario As Double
            precioUnitario = CDbl(Me.Txt_Precio_u.Value)
            Me.Txt_Precio.Value = modFormateo.FormatearNumero(precioUnitario * cantidad)
        End If
        
        If Me.Txt_Costo.Value <> "" And IsNumeric(Me.Txt_Costo.Value) And cantidad <> 0 Then
            Dim costoTotal As Double
            costoTotal = CDbl(Me.Txt_Costo.Value)
            Me.Txt_Costo_u.Value = modFormateo.FormatearNumero(costoTotal / cantidad)
        End If
    End If
End Sub

' --- EVENTO Change: recalcular precio del lote al cambiar precio unitario ----
Private Sub Txt_Precio_u_Change()
    If Me.Txt_Precio_u.Value <> "" And IsNumeric(Me.Txt_Precio_u.Value) Then
        If Me.Txt_Cantidad.Value <> "" And IsNumeric(Me.Txt_Cantidad.Value) Then
            Me.Txt_Precio.Value = modFormateo.FormatearNumero( _
                CDbl(Me.Txt_Precio_u.Value) * CDbl(Me.Txt_Cantidad.Value))
        End If
    End If
End Sub

' --- INICIALIZACIÓN DEL FORMULARIO -------------------------------------------
Private Sub UserForm_Initialize()
    ' Cargar ingredientes en los combos
    Dim i As Integer
    For i = 1 To 10
        Call modRecetas.CargarIngredientesEnCombo(Me.Controls("Cmb" & i))
        Me.Controls("Cmb_" & i).Clear
    Next i
    
    ' Si estamos en modo modificación, cargar datos
    If modoModificacion Then
        Me.Caption = "Modificar Receta"
        Me.Agregar.Caption = "Modificar"
        CargarDatosReceta
    Else
        Me.Caption = "Nueva Receta"
        Me.Agregar.Caption = "Agregar"
    End If
End Sub

' --- CARGAR DATOS DE LA RECETA (para modo modificación) ----------------------
Private Sub CargarDatosReceta()
    Dim wsRecetas As Worksheet
    Dim tblRecetas As ListObject
    Dim filaReceta As Range
    
    Set wsRecetas = ThisWorkbook.Sheets(HOJA_RECETAS)
    Set tblRecetas = wsRecetas.ListObjects(TBL_RECETAS)
    
    ' Buscar la receta
    For Each filaReceta In tblRecetas.ListColumns(COL_REC_NOMBRE).DataBodyRange
        If filaReceta.Value = recetaAModificar Then
            Dim fila As Long
            fila = filaReceta.Row
            Me.Txt_Receta.Value = wsRecetas.Cells(fila, tblRecetas.ListColumns(COL_REC_NOMBRE).Range.Column).Value
            Me.Txt_Cantidad.Value = wsRecetas.Cells(fila, tblRecetas.ListColumns(COL_REC_LOTE).Range.Column).Value
            Me.Txt_Precio_u.Value = wsRecetas.Cells(fila, tblRecetas.ListColumns(COL_REC_PRECIO_UNID).Range.Column).Value
            Exit For
        End If
    Next filaReceta
    
    ' Cargar ingredientes usando el módulo
    Dim ingredientes As Variant
    ingredientes = modRecetas.CargarIngredientesReceta(recetaAModificar)
    
    Dim i As Integer
    If IsArray(ingredientes) Then
        For i = 1 To UBound(ingredientes, 1)
            If i > 10 Then Exit For
            
            Dim cbIng As MSForms.comboBox
            Dim txtCant As MSForms.TextBox
            Dim cbUnd As MSForms.comboBox
            
            Set cbIng = Me.Controls("Cmb" & i)
            Set txtCant = Me.Controls("TextBox" & i)
            Set cbUnd = Me.Controls("Cmb_" & i)
            
            cbIng.Value = ingredientes(i, 1)
            txtCant.Value = modFormateo.FormatearNumero(ingredientes(i, 2))
            
            ' Cargar unidades según el ingrediente
            Call modRecetas.CargarUnidadesPorIngrediente(cbUnd, ingredientes(i, 1))
            cbUnd.Value = ingredientes(i, 3)
        Next i
    End If
    
    CalcularCostoTotal
End Sub

' --- CARGAR INGREDIENTES (wrapper para modRecetas) ---------------------------
Private Sub CargarIngredientes()
    Dim i As Integer
    For i = 1 To 10
        Call modRecetas.CargarIngredientesEnCombo(Me.Controls("Cmb" & i))
    Next i
End Sub

' --- EVENTOS MOUSE MOVE (UI) -------------------------------------------------
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.Btn_Reposo.ZOrder msoBringToFront
    Me.Agregar.ZOrder msoBringToFront
End Sub

Private Sub Btn_Reposo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.Btn_Active.ZOrder msoBringToFront
    Me.Agregar.ZOrder msoBringToFront
End Sub

' --- LIMPIAR VARIABLES AL CERRAR ---------------------------------------------
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    modoModificacion = False
    recetaAModificar = ""
End Sub


