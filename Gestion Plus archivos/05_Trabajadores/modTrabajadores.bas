Attribute VB_Name = "modTrabajadores"
' ============================================================================
' MÓDULO: modTrabajadores
' PROPÓSITO: Gestión de trabajadores - pago de salarios y llamadas a
' formularios relacionados.
' ============================================================================

Option Explicit

' --- PAGAR SALARIO A TRABAJADOR -------------------------------------------
' Calcula el monto a pagar basado en el % de comisión sobre ventas
' o el salario mínimo, lo que sea mayor.

Public Sub PagarTrabajador()
    Dim tbl_sel As ListObject
    Dim selected As Range
    Dim selected_row As ListRow
    Dim response As VbMsgBoxResult
    
    Set tbl_sel = ThisWorkbook.Sheets(HOJA_TRABAJADORES).ListObjects(TBL_TRABAJADORES)
    Set selected = Selection
    
    If Intersect(selected, tbl_sel.Range) Is Nothing Then
        MsgBox "Por favor seleccione un trabajador dentro de la tabla", vbExclamation
        Exit Sub
    End If
    
    Set selected_row = tbl_sel.ListRows(selected.Row - tbl_sel.HeaderRowRange.Row)
    
    ' Calcular pago: max(comisión, salario mínimo)
    ' Columna 3 = % comisión, Columna 4 = Salario mínimo
    ' J5 en Historial de ventas = Total de ventas (asumido)
    Dim totalVentas As Double
    totalVentas = ThisWorkbook.Worksheets(HOJA_VENTAS).Range("J5").Value
    
    If totalVentas * selected_row.Range(3).Value < selected_row.Range(4).Value Then
        pago = selected_row.Range(4).Value  ' Salario mínimo
    Else
        pago = totalVentas * selected_row.Range(3).Value  ' Comisión
    End If
    
    cargo = selected_row.Range(2).Value  ' Cargo para la descripción del gasto
    
    response = MsgBox("¿Desea pagar al trabajador: " & selected_row.Range(1).Value & "?" & vbCrLf & _
                      "A pagar: " & pago & " CUP", _
                      vbYesNo + vbQuestion, "Confirmar pago")
    
    If response = vbYes Then
        Call AgregarSalario(cargo, pago)
        Call AgregarFecha
        MsgBox "Salario pagado e incluido en la tabla de gastos correctamente", vbInformation
    End If
End Sub

' --- MODIFICAR TRABAJADOR (desde botón en hoja) ----------------------------

Public Sub ModificarTrabajador()
    Dim ws As Worksheet
    Dim tbl_sel As ListObject
    Dim selected As Range
    Dim selected_row As ListRow
    Dim response As VbMsgBoxResult
    
    Set ws = ThisWorkbook.Sheets(HOJA_TRABAJADORES)
    Set tbl_sel = ws.ListObjects(TBL_TRABAJADORES)
    Set selected = Selection
    
    If Intersect(selected, tbl_sel.Range) Is Nothing Then
        MsgBox "Por favor seleccione un trabajador dentro de la tabla", vbExclamation
        Exit Sub
    End If
    
    Set selected_row = tbl_sel.ListRows(selected.Row - tbl_sel.HeaderRowRange.Row)
    
    response = MsgBox("¿Desea modificar el trabajador: " & selected_row.Range(1).Value & "?", _
                      vbYesNo + vbQuestion, "Confirmar modificación")
    
    If response = vbYes Then
        fila_sel = selected_row.Index
        Trabajador_Mod.Show
    End If
End Sub

' --- AGREGAR TRABAJADOR NUEVO ------------------------------------------------

Public Sub AgregarTrabajador(ByVal nombre As String, _
                              ByVal cargo As String, _
                              ByVal salarioPct As Double, _
                              ByVal salarioMin As Double)
    Dim ws As Worksheet
    Dim tbl As ListObject
    
    Set ws = ThisWorkbook.Sheets(HOJA_TRABAJADORES)
    Set tbl = ws.ListObjects(TBL_TRABAJADORES)
    
    ws.Unprotect password:=PASS_HOJA
    
    With tbl.ListRows.Add
        .Range(COL_TRAB_NOMBRE) = nombre
        .Range(COL_TRAB_CARGO) = cargo
        .Range(COL_TRAB_PORCENTAJE) = Round(salarioPct, 4)
        .Range(COL_TRAB_SAL_MIN) = Format(Round(salarioMin, 3), "General Number")
    End With
    
    tbl.Sort.Apply
    ws.Protect password:=PASS_HOJA, UserInterfaceOnly:=True
End Sub

' --- MODIFICAR TRABAJADOR EXISTENTE ------------------------------------------

Public Sub ModificarTrabajadorEnTabla(ByVal indiceFila As Long, _
                                       ByVal nombre As String, _
                                       ByVal cargo As String, _
                                       ByVal salarioPct As Double, _
                                       ByVal salarioMin As Double)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim sel_row As ListRow
    
    Set ws = ThisWorkbook.Sheets(HOJA_TRABAJADORES)
    Set tbl = ws.ListObjects(TBL_TRABAJADORES)
    Set sel_row = tbl.ListRows(indiceFila)
    
    ws.Unprotect password:=PASS_HOJA
    
    With sel_row
        .Range(COL_TRAB_NOMBRE) = nombre
        .Range(COL_TRAB_CARGO) = cargo
        .Range(COL_TRAB_PORCENTAJE) = Format(Round(salarioPct, 4), "General Number")
        .Range(COL_TRAB_SAL_MIN) = Format(Round(salarioMin, 3), "General Number")
    End With
    
    tbl.Sort.Apply
    ws.Protect password:=PASS_HOJA, UserInterfaceOnly:=True
End Sub
