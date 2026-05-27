Attribute VB_Name = "modGeneradorCodigo"
' ============================================================================
' MÓDULO: modGeneradorCodigo
' PROPÓSITO: Generación de códigos de producto en formato "Pr_00001".
' Reemplaza al antiguo módulo Generador_codigo.
' ============================================================================

Option Explicit

' --- GENERAR CÓDIGO DE PRODUCTO --------------------------------------------

Public Sub GenerarCodigoProducto()
    Dim wsAlmacen As Worksheet
    Dim tblInventario As ListObject
    Dim i As Long
    Dim codigoActual As String
    Dim numeroActual As Long
    Dim maxNumero As Long
    
    Set wsAlmacen = ThisWorkbook.Sheets(HOJA_ALMACEN)
    Set tblInventario = wsAlmacen.ListObjects(TBL_INVENTARIO)
    
    maxNumero = 0
    
    ' Recorrer códigos existentes para encontrar el máximo
    If tblInventario.ListRows.Count > 0 Then
        For i = 1 To tblInventario.ListRows.Count
            With tblInventario.ListRows(i)
                codigoActual = .Range.Cells(1, 1).Value
                
                If codigoActual <> "" And Len(codigoActual) > 3 Then
                    On Error Resume Next
                    numeroActual = CLng(Mid(codigoActual, 4))
                    On Error GoTo 0
                    
                    If numeroActual > maxNumero Then
                        maxNumero = numeroActual
                    End If
                End If
            End With
        Next i
    End If
    
    ' Generar nuevo código
    maxNumero = maxNumero + 1
    Codigo_Producto = "Pr_" & Format(maxNumero, "00000")
End Sub
