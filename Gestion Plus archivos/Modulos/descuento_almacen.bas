Attribute VB_Name = "descuento_almacen"
' ============================================================================
' MODULO: descuento_almacen (MIGRADO - Fase 2)
' PROPOSITO: Fachada que delega en modMermas y modStock.
' Mantiene las funciones UDF para compatibilidad con formulas de Excel.
' ============================================================================

'Function ObtenerMermasProducto(codigo As String) As Double
'    ObtenerMermasProducto = modMermas.ObtenerMermasProducto(codigo)
'End Function

Function actualizar_almacen(codigo As String)
    actualizar_almacen = modStock.ObtenerStockRestante(codigo)
End Function
