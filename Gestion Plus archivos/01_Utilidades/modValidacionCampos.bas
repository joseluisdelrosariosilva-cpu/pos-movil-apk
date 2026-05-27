Attribute VB_Name = "modValidacionCampos"
' ============================================================================
' MÓDULO: modValidacionCampos
' PROPÓSITO: Funciones de validación de campos vacíos para todos los
' formularios. Unifica los módulos Validar_campos_vacios,
' Validar_vacios_2 y Validar_vacios_3.
' ============================================================================
' NOTA: Las variables contador/contador2/contador3 están declaradas
'       en modGlobales.
' ============================================================================

Option Explicit

' --- VALIDACIÓN: INGRESAR PRODUCTO (ingresar_form) -------------------------

Public Sub ValidarVaciosProducto(frm As ingresar_form)
    contador = 0
    
    If frm.nombre_de_producto_TextBox.Value = "" Then contador = contador + 1
    If frm.fecha_de_entrada_TextBox.Value = "" Then contador = contador + 1
    If frm.cantidad_TextBox.Value = "" Then contador = contador + 1
    If frm.precio_venta_TextBox.Value = "" Then contador = contador + 1
    If frm.precio_costo_TextBox.Value = "" Then contador = contador + 1
End Sub

' --- VALIDACIÓN: INGRESAR NUEVO INGREDIENTE (ingresar_ingr_nuevo_form) -----

Public Sub ValidarVaciosIngredienteNuevo(frm As ingresar_ingr_nuevo_form)
    contador2 = 0
    
    ' Resaltar campos vacíos con fondo amarillo
    If frm.Nombre_ingrediente_Nuevo_TextBox.Value = "" Then
        frm.Nombre_ingrediente_Nuevo_TextBox.BackStyle = fmBackStyleOpaque
        frm.Nombre_ingrediente_Nuevo_TextBox.BackColor = vbYellow
        contador2 = contador2 + 1
    Else
        frm.Nombre_ingrediente_Nuevo_TextBox.BackStyle = fmBackStyleTransparent
    End If
    
    If frm.Cantidad_ingresar_nuevo_producto_TextBox.Value = "" Then
        frm.Cantidad_ingresar_nuevo_producto_TextBox.BackStyle = fmBackStyleOpaque
        frm.Cantidad_ingresar_nuevo_producto_TextBox.BackColor = vbYellow
        contador2 = contador2 + 1
    Else
        frm.Cantidad_ingresar_nuevo_producto_TextBox.BackStyle = fmBackStyleTransparent
    End If
    
    If frm.Cantidad_referencia_TextBox.Value = "" Then
        frm.Cantidad_referencia_TextBox.BackStyle = fmBackStyleOpaque
        frm.Cantidad_referencia_TextBox.BackColor = vbYellow
        contador2 = contador2 + 1
    Else
        frm.Cantidad_referencia_TextBox.BackStyle = fmBackStyleTransparent
    End If
    
    If frm.Costo_Total_Ingr_Nuevo_TextBox.Value = "" Then
        frm.Costo_Total_Ingr_Nuevo_TextBox.BackStyle = fmBackStyleOpaque
        frm.Costo_Total_Ingr_Nuevo_TextBox.BackColor = vbYellow
        contador2 = contador2 + 1
    Else
        frm.Costo_Total_Ingr_Nuevo_TextBox.BackStyle = fmBackStyleTransparent
    End If
    
    If frm.UBase_nuevo_ComboBox.Value = "" Then
        frm.UBase_nuevo_ComboBox.BackStyle = fmBackStyleOpaque
        frm.UBase_nuevo_ComboBox.BackColor = vbYellow
        contador2 = contador2 + 1
    Else
        frm.UBase_nuevo_ComboBox.BackStyle = fmBackStyleTransparent
    End If
End Sub

' --- VALIDACIÓN: INGRESAR INGREDIENTE EXISTENTE (ingresar_ingr_existente_form) --

Public Sub ValidarVaciosIngredienteExistente(frm As ingresar_ingr_existente_form)
    contador3 = 0
    
    If frm.Nombre_ingr_ComboBox.Value = "" Then
        frm.Nombre_ingr_ComboBox.BackStyle = fmBackStyleOpaque
        frm.Nombre_ingr_ComboBox.BackColor = vbYellow
        contador3 = contador3 + 1
    Else
        frm.Nombre_ingr_ComboBox.BackStyle = fmBackStyleTransparent
    End If
    
    If frm.Cantidad_ingresar_producto_existente_TextBox.Value = "" Then
        frm.Cantidad_ingresar_producto_existente_TextBox.BackStyle = fmBackStyleOpaque
        frm.Cantidad_ingresar_producto_existente_TextBox.BackColor = vbYellow
        contador3 = contador3 + 1
    Else
        frm.Cantidad_ingresar_producto_existente_TextBox.BackStyle = fmBackStyleTransparent
    End If
    
    If frm.Unidad_ComboBox.Value = "" Then
        frm.Unidad_ComboBox.BackStyle = fmBackStyleOpaque
        frm.Unidad_ComboBox.BackColor = vbYellow
        contador3 = contador3 + 1
    Else
        frm.Unidad_ComboBox.BackStyle = fmBackStyleTransparent
    End If
    
    If frm.Costo_TextBox.Value = "" Then
        frm.Costo_TextBox.BackStyle = fmBackStyleOpaque
        frm.Costo_TextBox.BackColor = vbYellow
        contador3 = contador3 + 1
    Else
        frm.Costo_TextBox.BackStyle = fmBackStyleTransparent
    End If
End Sub
