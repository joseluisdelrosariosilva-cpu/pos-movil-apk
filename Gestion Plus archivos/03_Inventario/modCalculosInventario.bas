Attribute VB_Name = "modCalculosInventario"
' ============================================================================
' MÓDULO: modCalculosInventario
' PROPÓSITO: Cálculos de inventario como precio por unidad base, costo de
' inversión, etc.
' ============================================================================

Option Explicit

' --- CALCULAR PRECIO POR UNIDAD BASE (en Gestión de Ingredientes) ----------

Public Sub CalcularPrecioUBase()
    Dim i As Long
    Dim resultado As Double
    Dim ultima_fila As Long
    Dim HojaBD As Worksheet
    
    Set HojaBD = ThisWorkbook.Sheets(HOJA_INGREDIENTES)
    ultima_fila = HojaBD.Range("B" & Rows.Count).End(xlUp).Row
    
    For i = 5 To ultima_fila
        If HojaBD.Range("G" & i).Value <> "" And HojaBD.Range("D" & i).Value <> "" Then
            If HojaBD.Range("D" & i).Value <> 0 Then
                resultado = HojaBD.Range("G" & i).Value / HojaBD.Range("D" & i).Value
                HojaBD.Range("F" & i).Value = resultado
            Else
                HojaBD.Range("F" & i).Value = "#DIV/0!"
            End If
        Else
            HojaBD.Range("F" & i).Value = ""
        End If
    Next i
End Sub

' --- CALCULAR MONTO DE INVERSIÓN DE UN PRODUCTO VENDIDO --------------------

Public Function CalcularMontoInversion(idProducto As String, cantidad As Double) As Double
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim fila As ListRow
    Dim costo As Double
    
    Set ws = ThisWorkbook.Sheets(HOJA_ALMACEN)
    Set tbl = ws.ListObjects(TBL_INVENTARIO)
    
    For Each fila In tbl.ListRows
        If fila.Range(COL_INV_CODIGO).Value = idProducto Then
            costo = fila.Range(COL_INV_PRECIO_C).Value
            CalcularMontoInversion = costo * cantidad
            Exit Function
        End If
    Next fila
    
    CalcularMontoInversion = 0
End Function

' --- OBTENER PRECIO POR UNIDAD BASE DE INGREDIENTE -------------------------

Public Function ObtenerPrecioPorUnidadBase(ingrediente As String) As Double
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim fila As ListRow
    
    Set ws = ThisWorkbook.Sheets(HOJA_INGREDIENTES)
    Set tbl = ws.ListObjects(TBL_INGREDIENTES)
    
    For Each fila In tbl.ListRows
        If fila.Range.Cells(1, 1).Value = ingrediente Then
            ' Columna 5 = Precio por Unidad Base
            ObtenerPrecioPorUnidadBase = fila.Range.Cells(1, 5).Value
            Exit Function
        End If
    Next fila
    
    ObtenerPrecioPorUnidadBase = 0
End Function
