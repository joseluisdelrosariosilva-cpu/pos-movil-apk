Attribute VB_Name = "modSincronizacion"
' ============================================================================
' MÓDULO: modSincronizacion
' PROPÓSITO: Sincronización de datos entre el Excel y el archivo datos.xlsx
' que usa la webapp.
' ============================================================================

Option Explicit

' --- SINCRONIZAR PRODUCTOS HACIA datos.xlsx --------------------------------

Public Sub SincronizarProductosEnDatos()
    Dim wbOrigen As Workbook
    Dim wsOrigen As Worksheet
    Dim wbDestino As Workbook
    Dim wsDestino As Worksheet
    Dim rutaDestino As String
    Dim paso As String
    Dim abiertoPorMacro As Boolean
    Dim tbl As ListObject
    Dim dataArr As Variant
    Dim arrOut() As Variant
    Dim i As Long, numFilas As Long
    
    On Error GoTo ErrHandler
    
    paso = "Inicializar referencias"
    Set wbOrigen = ThisWorkbook
    Set wsOrigen = wbOrigen.Worksheets(HOJA_ALMACEN)
    rutaDestino = wbOrigen.Path & "\" & CARPETA_WEBAPP & ARCHIVO_DATOS
    
    paso = "Abrir libro destino"
    Set wbDestino = GetWorkbookByFullName(rutaDestino)
    If wbDestino Is Nothing Then
        Set wbDestino = Workbooks.Open(Filename:=rutaDestino, ReadOnly:=False)
        abiertoPorMacro = True
    End If
    
    paso = "Obtener hoja destino"
    Set wsDestino = wbDestino.Worksheets("Productos")
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    paso = "Leer tabla INVENTARIO"
    Set tbl = wsOrigen.ListObjects(TBL_INVENTARIO)
    
    If tbl.DataBodyRange Is Nothing Then GoTo Finalizar
    
    dataArr = tbl.DataBodyRange.Value2
    numFilas = UBound(dataArr, 1)
    
    ' Mapeo: Col1(B)->A(Código), Col2(C)->B(Nombre),
    '        Col5(F)->C(Stock), Col6(G)->D(PrecioVta)
    paso = "Mapear columnas"
    ReDim arrOut(1 To numFilas, 1 To 4)
    For i = 1 To numFilas
        arrOut(i, 1) = dataArr(i, COL_INV_CODIGO)
        arrOut(i, 2) = dataArr(i, COL_INV_NOMBRE)
        arrOut(i, 3) = dataArr(i, COL_INV_CANT_ACT)
        arrOut(i, 4) = dataArr(i, COL_INV_PRECIO_V)
    Next i
    
    paso = "Volcar datos al destino"
    wsDestino.Range("A2:D" & wsDestino.Rows.Count).ClearContents
    wsDestino.Range("A2").Resize(numFilas, 4).Value = arrOut
    
Finalizar:
    paso = "Guardar y cerrar"
    wbDestino.Save
    
    If abiertoPorMacro Then
        wbDestino.Close SaveChanges:=False
    End If
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub
    
ErrHandler:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    On Error Resume Next
    If Not wbDestino Is Nothing Then
        If abiertoPorMacro Then wbDestino.Close SaveChanges:=False
    End If
    On Error GoTo 0
    
    MsgBox "Error al sincronizar productos." & vbCrLf & _
           "Paso: " & paso & vbCrLf & _
           "Error: " & Err.Description, vbCritical
End Sub

' --- COPIAR ÚLTIMA FACTURA DEL DÍA A datos.xlsx ----------------------------

Public Sub CopiarUltimaFacturaDelDia()
    Dim wbOrigen As Workbook
    Dim wbDestino As Workbook
    Dim wsOrigen As Worksheet
    Dim wsDestino As Worksheet
    Dim tblOrigen As ListObject
    Dim rngFacturas As Range
    Dim celda As Range
    Dim facturaID As String
    Dim facturaActual As String
    Dim fechaHoyNumero As String
    Dim maxContador As Long
    Dim contadorActual As Long
    Dim tieneFacturasHoy As Boolean
    Dim rutaDestino As String
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    fechaHoyNumero = Format(CLng(Date), "00000")
    rutaDestino = ThisWorkbook.Path & "\" & CARPETA_WEBAPP & ARCHIVO_DATOS
    
    Set wbOrigen = ThisWorkbook
    Set wsOrigen = wbOrigen.Sheets(HOJA_FACTURAS)
    Set tblOrigen = wsOrigen.ListObjects(TBL_FACTURAS)
    
    On Error Resume Next
    Set rngFacturas = tblOrigen.DataBodyRange.Columns(1)
    On Error GoTo 0
    
    If rngFacturas Is Nothing Then
        facturaID = fechaHoyNumero & "00000"
        maxContador = 0
    Else
        maxContador = -1
        tieneFacturasHoy = False
        
        For Each celda In rngFacturas.Cells
            facturaActual = Trim(CStr(celda.Value))
            
            If Left(facturaActual, 5) = fechaHoyNumero Then
                tieneFacturasHoy = True
                contadorActual = CLng(Right(facturaActual, 5))
                
                If contadorActual > maxContador Then
                    maxContador = contadorActual
                    facturaID = facturaActual
                End If
            End If
        Next celda
    End If
    
    ' Abrir y escribir en datos.xlsx
    Set wbDestino = Workbooks.Open(rutaDestino)
    Set wsDestino = wbDestino.Sheets("Pendientes")
    
    With wsDestino.ListObjects("Tabla1").ListRows.Add.Range
        .Cells(1, 1).Value = facturaID
        .Cells(1, 10).Value = "VERDADERO"
    End With
    
    wbDestino.Save
    wbDestino.Close SaveChanges:=False
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub

' --- OBTENER WORKBOOK POR NOMBRE COMPLETO ----------------------------------

Public Function GetWorkbookByFullName(ByVal fullName As String) As Workbook
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If StrComp(wb.fullName, fullName, vbTextCompare) = 0 Then
            Set GetWorkbookByFullName = wb
            Exit Function
        End If
    Next wb
End Function

' --- ACTUALIZAR DATOS DEL FORMULARIO POS -----------------------------------

Public Sub ActualizarDatosFormulario()
    On Error GoTo SalirSeguro
    
    Dim frm As Object
    For Each frm In VBA.UserForms
        If TypeName(frm) = "frmPOSMovil" Then
            If frm.visible Then frm.ActualizarDatosFormulario
            Exit For
        End If
    Next frm
    
SalirSeguro:
    On Error Resume Next
End Sub

