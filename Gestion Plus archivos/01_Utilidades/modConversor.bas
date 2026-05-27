Attribute VB_Name = "modConversor"
' ============================================================================
' MÓDULO: modConversor
' PROPÓSITO: Conversión de unidades de medida para ingredientes.
' Versión mejorada con soporte completo de unidades.
' ============================================================================

Option Explicit

' --- CONVERTIR UNIDAD DE UN INGREDIENTE A SU UNIDAD BASE -------------------
' Devuelve la cantidad convertida a la unidad base del ingrediente.
' Ejemplo: ConvertirUnidad("Harina", 1000, "g") con base kg -> 1

Public Function ConvertirUnidad(ByVal ingrediente As String, _
                                 ByVal cantidad As Double, _
                                 ByVal unidadOrigen As String) As Double
    Dim tblConversiones As ListObject
    Dim fila As ListRow
    Dim unidadBase As String
    
    ' Si la unidad origen es "u", no convertir (unidades discretas)
    If LCase(Trim(unidadOrigen)) = "u" Then
        ConvertirUnidad = cantidad
        Exit Function
    End If
    
    ' Obtener unidad base del ingrediente
    Set tblConversiones = ThisWorkbook.Sheets(HOJA_INGREDIENTES).ListObjects(TBL_INGREDIENTES)
    
    For Each fila In tblConversiones.ListRows
        If fila.Range.Cells(1, COL_ING_NOMBRE).Value = ingrediente Then
            unidadBase = fila.Range.Cells(1, COL_ING_UNIDAD_BASE).Value
            Exit For
        End If
    Next fila
    
    ' Convertir
    ConvertirUnidad = ConvertirABase(cantidad, unidadOrigen, unidadBase)
End Function

' --- CONVERTIR DE UNIDAD ORIGEN A UNIDAD BASE ESPECÍFICA -------------------

Private Function ConvertirABase(ByVal cantidad As Double, _
                                 ByVal unidadOrigen As String, _
                                 ByVal unidadBase As String) As Double
    Dim cantidadEnBase As Double
    
    ' Paso 1: Convertir a unidad común (gramos o mL)
    Select Case LCase(Trim(unidadOrigen))
        ' Masa
        Case "kg"
            cantidadEnBase = cantidad * 1000     ' -> gramos
        Case "g"
            cantidadEnBase = cantidad             ' gramos
        Case "mg"
            cantidadEnBase = cantidad / 1000      ' -> gramos
        Case "lb"
            cantidadEnBase = cantidad * 453.592   ' -> gramos
        Case "oz"
            cantidadEnBase = cantidad * 28.3495   ' -> gramos
        
        ' Volumen
        Case "l"
            cantidadEnBase = cantidad * 1000      ' -> mL
        Case "ml", "mL"
            cantidadEnBase = cantidad             ' mL
        
        Case Else
            ' Unidad no reconocida, devolver tal cual
            ConvertirABase = cantidad
            Exit Function
    End Select
    
    ' Paso 2: Convertir de unidad común a la unidad base específica
    Select Case LCase(Trim(unidadBase))
        ' Masa desde gramos
        Case "kg"
            ConvertirABase = cantidadEnBase / 1000
        Case "g"
            ConvertirABase = cantidadEnBase
        Case "mg"
            ConvertirABase = cantidadEnBase * 1000
        Case "lb"
            ConvertirABase = cantidadEnBase / 453.592
        Case "oz"
            ConvertirABase = cantidadEnBase / 28.3495
        
        ' Volumen desde mL
        Case "l"
            ConvertirABase = cantidadEnBase / 1000
        Case "ml", "mL"
            ConvertirABase = cantidadEnBase
        
        ' Unidad discreta
        Case "u"
            ConvertirABase = cantidad
        
        Case Else
            ConvertirABase = cantidadEnBase
    End Select
End Function

' --- NORMALIZAR NOMBRE DE UNIDAD (evitar errores de mayúsculas) ------------

Public Function NormalizarUnidad(ByVal unidad As String) As String
    Dim u As String
    u = LCase(Trim(unidad))
    
    Select Case u
        Case "kg", "kilo", "kilos"
            NormalizarUnidad = "kg"
        Case "g", "gramo", "gramos"
            NormalizarUnidad = "g"
        Case "mg", "miligramo", "miligramos"
            NormalizarUnidad = "mg"
        Case "lb", "libra", "libras"
            NormalizarUnidad = "lb"
        Case "oz", "onza", "onzas"
            NormalizarUnidad = "oz"
        Case "l", "litro", "litros"
            NormalizarUnidad = "l"
        Case "ml", "mililitro", "mililitros", "mL"
            NormalizarUnidad = "ml"
        Case "u", "unidad", "unidades"
            NormalizarUnidad = "u"
        Case Else
            NormalizarUnidad = unidad
    End Select
End Function
