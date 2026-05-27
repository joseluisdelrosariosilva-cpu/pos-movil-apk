Attribute VB_Name = "modMermas"
' ============================================================================
' MÓDULO: modMermas
' PROPÓSITO: Gestión de mermas - consulta, registro y procesamiento.
' ============================================================================

Option Explicit

' --- OBTENER MERMAS DE UN PRODUCTO (desde Historial de Gastos) -------------
' Función de hoja: recorre todos los gastos con categoría "Merma" y suma
' las cantidades asociadas al producto.

Public Function ObtenerMermasProducto(codigo As String) As Double
    Dim wsGastos As Worksheet
    Dim tblGastos As ListObject
    Dim fila As ListRow
    Dim nombreProducto As String
    Dim totalMermas As Double
    Dim descripcion As String
    Dim posX As Long
    Dim cantidadStr As String
    Dim wsInv As Worksheet
    Dim tblInv As ListObject
    Dim filaInv As ListRow
    
    On Error GoTo ErrorHandler
    
    ' Obtener el nombre del producto desde INVENTARIO
    Set wsInv = ThisWorkbook.Sheets(HOJA_ALMACEN)
    Set tblInv = wsInv.ListObjects(TBL_INVENTARIO)
    
    nombreProducto = ""
    For Each filaInv In tblInv.ListRows
        If CStr(filaInv.Range(COL_INV_CODIGO).Value) = codigo Then
            nombreProducto = CStr(filaInv.Range(COL_INV_NOMBRE).Value)
            Exit For
        End If
    Next filaInv
    
    If nombreProducto = "" Then
        ObtenerMermasProducto = 0
        Exit Function
    End If
    
    ' Buscar en Historial de Gastos
    Set wsGastos = ThisWorkbook.Sheets(HOJA_GASTOS)
    Set tblGastos = wsGastos.ListObjects(TBL_GASTOS)
    
    totalMermas = 0
    
    For Each fila In tblGastos.ListRows
        If CStr(fila.Range(COL_GAS_CATEGORIA).Value) = CAT_MERMA Then
            descripcion = CStr(fila.Range(COL_GAS_DESCRIPCION).Value)
            
            ' Formato esperado: "NombreProducto x cantidad"
            If InStr(1, descripcion, nombreProducto & " x ") = 1 Then
                posX = InStrRev(descripcion, " x ")
                If posX > 0 Then
                    cantidadStr = Trim(Mid(descripcion, posX + 3))
                    If IsNumeric(cantidadStr) Then
                        totalMermas = totalMermas + CDbl(cantidadStr)
                    End If
                End If
            End If
        End If
    Next fila
    
    ObtenerMermasProducto = totalMermas
    Exit Function
    
ErrorHandler:
    ObtenerMermasProducto = 0
End Function

' --- REGISTRAR MERMA DESDE INVENTARIO (manual) -----------------------------

Public Sub RegistrarMerma()
    Dim ws As Worksheet
    Dim selectedRow As Long
    Dim codigo As String
    Dim nombre As String
    Dim stock As Double
    Dim precioCosto As Double
    Dim Cantidad As Double
    Dim monto As Double
    Dim inputStr As String
    
    Set ws = ThisWorkbook.Sheets(HOJA_ALMACEN)
    
    fila_a_modificar = ActiveCell.Row
    
    ' Validar selección
    If fila_a_modificar <= 4 Or ws.Range("B" & fila_a_modificar) = "" Then
        MsgBox "Selecciona una fila del inventario.", vbExclamation
        Exit Sub
    End If
    
    selectedRow = Selection.Row
    
    ' Leer valores del inventario
    codigo = CStr(ws.Cells(selectedRow, 2).Value)
    nombre = CStr(ws.Cells(selectedRow, 3).Value)
    stock = CDbl(ws.Cells(selectedRow, 6).Value)
    precioCosto = CDbl(ws.Cells(selectedRow, 8).Value)
    
    If stock <= 0 Then
        MsgBox "El producto no tiene stock disponible.", vbExclamation
        Exit Sub
    End If
    
    ' Solicitar cantidad
    inputStr = InputBox("Ingresa la cantidad de merma para " & nombre & _
                        " (Stock disponible: " & stock & "):", "Registrar Merma", "1")
    
    If inputStr = "" Then Exit Sub
    If Not IsNumeric(inputStr) Then
        MsgBox "Cantidad inválida. Ingresa un número.", vbExclamation
        Exit Sub
    End If
    
    Cantidad = CDbl(inputStr)
    
    If Cantidad <= 0 Then
        MsgBox "La cantidad debe ser mayor a 0.", vbExclamation
        Exit Sub
    End If
    
    If Cantidad > stock Then
        MsgBox "La cantidad no puede superar el stock disponible (" & stock & ").", vbExclamation
        Exit Sub
    End If
    
    ' Registrar gasto como merma
    monto = precioCosto * Cantidad
    Call AgregarGasto(Date, CAT_MERMA, nombre & " x " & Cantidad, monto)
    Call modAgregarRegistros.AgregarFecha(Date)
    
    ' Forzar actualización de stock
    Call ForzarActualizacionStock
    
    MsgBox "Merma registrada correctamente: " & nombre & " x " & Cantidad, vbInformation
End Sub

' --- PROCESAR MERMAS DESDE datos.xlsx -------------------------------------

Public Sub ProcesarMermas()
    Dim wb As Workbook
    Dim wsMerma As Worksheet
    Dim wsInventario As Worksheet
    Dim ruta As String
    Dim ultimaFilaMerma As Long
    Dim i As Long
    Dim codigo As String
    Dim nombre As String
    Dim Cantidad As Double
    Dim precioCosto As Double
    Dim fila As ListRow
    Dim monto As Double
    
    ruta = ThisWorkbook.Path & "\" & CARPETA_WEBAPP & ARCHIVO_DATOS
    If Dir(ruta) = "" Then
        MsgBox "No se encontró datos.xlsx", vbExclamation
        Exit Sub
    End If
    
    Set wb = Workbooks.Open(ruta, ReadOnly:=False)
    
    On Error Resume Next
    Set wsMerma = wb.Sheets("Merma")
    On Error GoTo 0
    
    If wsMerma Is Nothing Then
        wb.Close False
        Exit Sub
    End If
    
    If wsMerma.Cells(2, 1).Value = "" Then
        wb.Save
        wb.Close False
        Exit Sub
    End If
    
    Set wsInventario = ThisWorkbook.Sheets(HOJA_ALMACEN)
    ultimaFilaMerma = wsMerma.Cells(wsMerma.Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To ultimaFilaMerma
        codigo = CStr(wsMerma.Cells(i, 1).Value)
        nombre = CStr(wsMerma.Cells(i, 2).Value)
        Cantidad = CDbl(wsMerma.Cells(i, 3).Value)
        
        ' Buscar precio costo en inventario
        For Each fila In wsInventario.ListObjects(TBL_INVENTARIO).ListRows
            If CStr(fila.Range(COL_INV_CODIGO).Value) = codigo Then
                precioCosto = CDbl(fila.Range(COL_INV_PRECIO_C).Value)
                monto = precioCosto * Cantidad
                Call AgregarGasto(Date, CAT_MERMA, nombre & " x " & Cantidad, monto)
                Call modAgregarRegistros.AgregarFecha(Date)
                Exit For
            End If
        Next fila
    Next i
    
    ' Actualizar stock
    Call ForzarActualizacionStock
    
    ' Eliminar filas procesadas
    For i = ultimaFilaMerma To 2 Step -1
        If wsMerma.Cells(i, 1).Value = "" Then Exit For
        wsMerma.Rows(i).Delete
    Next i
    
    wb.Save
    wb.Close False
    
    MsgBox "Mermas procesadas correctamente.", vbInformation
End Sub
