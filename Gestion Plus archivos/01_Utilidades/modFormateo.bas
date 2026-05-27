Attribute VB_Name = "modFormateo"
' ============================================================================
' MÓDULO: modFormateo
' PROPÓSITO: Funciones auxiliares de formateo de valores numéricos y fechas.
' ============================================================================

Option Explicit

' --- FORMATEAR NÚMERO ------------------------------------------------------

Public Function FormatearNumero(valor As Variant) As String
    If IsNumeric(valor) Then
        FormatearNumero = Format(Round(CDbl(valor), 3), "General Number")
    Else
        FormatearNumero = ""
    End If
End Function

' --- FORMATEAR FECHA -------------------------------------------------------

Public Function FormatearFecha(fecha As Variant) As String
    If IsDate(fecha) Then
        FormatearFecha = Format(CDate(fecha), "d-mmm-yyyy")
    Else
        FormatearFecha = ""
    End If
End Function

' --- CONVERTIR A DOUBLE (0 si vacío) --------------------------------------

Public Function ToDoubleOrZero(ByVal v As Variant) As Double
    On Error GoTo ErrorHandler
    
    If IsNull(v) Or IsEmpty(v) Then
        ToDoubleOrZero = 0#
        Exit Function
    End If
    
    If VarType(v) = vbString Then
        If Len(Trim$(CStr(v))) = 0 Then
            ToDoubleOrZero = 0#
            Exit Function
        End If
    End If
    
    If IsNumeric(v) Then
        ToDoubleOrZero = CDbl(v)
        Exit Function
    End If
    
    Err.Raise vbObjectError + 3201, "ToDoubleOrZero", "Valor no numérico: " & CStr(v)
    
ErrorHandler:
    Err.Raise vbObjectError + 3299, "ToDoubleOrZero", "No se pudo convertir a número: " & CStr(v)
End Function
