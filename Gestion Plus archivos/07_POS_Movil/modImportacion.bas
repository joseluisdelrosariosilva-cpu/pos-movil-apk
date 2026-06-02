Attribute VB_Name = "modImportacion"
' ============================================================================
' MÓDULO: modImportacion
' PROPÓSITO: Importación de datos desde datos.xlsx hacia el libro principal.
' Incluye: ventas, gastos, entradas, abastecimientos, mermas.
' ============================================================================

Option Explicit

' --- MACRO PRINCIPAL: IMPORTAR VENTAS --------------------------------------

Public Sub ImportarVentas()
    On Error GoTo ErrorHandler
    
    Dim resultado As Variant
    Dim datos As Variant
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim facturas As Object
    Dim facturaData As Object
    Dim lineas As Collection
    Dim clave As Variant
    Dim linea As Variant
    Dim facturaID As String
    Dim i As Long
    Dim filaExcel As Long
    Dim ultimaFila As Long
    Dim fechaVenta As Date
    Dim totalFactura As Double
    Dim efectivo As Double
    Dim transferencia As Double
    Dim contadorFacturas As Long
    Dim rutaLog As String
    
    Application.ScreenUpdating = False
    
    ' 1. Importar datos complementarios primero
    Call ImportarEntradas
    Application.ScreenUpdating = False
    Call ActualizarInventarioDesdeAbastecimiento
    Application.ScreenUpdating = False
    Call ProcesarMermas
    Application.ScreenUpdating = False
    Call ProcesarElaboraciones
    Application.ScreenUpdating = False
    Call ImportarGastos
    Application.ScreenUpdating = False
    
    ' 2. Leer ventas pendientes
    resultado = ObtenerVentasPendientes()
    
    On Error Resume Next
    If (Not IsArray(resultado)) Or UBound(resultado) < 1 Then
        On Error GoTo 0
        MsgBox "No hay ventas pendientes para importar.", vbInformation
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    
    datos = resultado(0)
    Set wb = resultado(1)
    Set ws = wb.Sheets("Pendientes")
    
    On Error Resume Next
    If UBound(datos, 1) < 1 Then
        On Error GoTo 0
        MsgBox "No hay ventas pendientes para importar.", vbInformation
        wb.Close False
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    
    ' 3. Agrupar por factura
    Set facturas = CreateObject("Scripting.Dictionary")
    
    For i = 1 To UBound(datos, 1)
        facturaID = CStr(datos(i, 1))
        If Len(facturaID) = 0 Then GoTo SiguienteFila
        
        If Not facturas.Exists(facturaID) Then
            Set facturaData = CreateObject("Scripting.Dictionary")
            Set lineas = New Collection
            facturaData.Add "fecha", ParseFechaVenta(datos(i, 2))
            facturaData.Add "efectivo", ToDoubleOrZero(datos(i, 8))
            facturaData.Add "transferencia", ToDoubleOrZero(datos(i, 9))
            facturaData.Add "total", 0#
            facturaData.Add "lineas", lineas
            facturas.Add facturaID, facturaData
        End If
        
        Set facturaData = facturas(facturaID)
        facturaData("total") = ToDoubleOrZero(facturaData("total")) + ToDoubleOrZero(datos(i, 7))
        facturaData("lineas").Add Array(datos(i, 3), datos(i, 4), _
                                        ToDoubleOrZero(datos(i, 5)), _
                                        ToDoubleOrZero(datos(i, 6)), _
                                        ToDoubleOrZero(datos(i, 7)))
SiguienteFila:
    Next i
    
    ' 4. Insertar facturas y líneas en el libro
    contadorFacturas = 0
    
    For Each clave In facturas.Keys
        facturaID = CStr(clave)
        Set facturaData = facturas(facturaID)
        
        fechaVenta = ParseFechaVenta(facturaData("fecha"))
        efectivo = ToDoubleOrZero(facturaData("efectivo"))
        transferencia = ToDoubleOrZero(facturaData("transferencia"))
        totalFactura = ToDoubleOrZero(facturaData("total"))
        
        If InsertarFacturaEnLibro(facturaID, fechaVenta, totalFactura, efectivo, transferencia) Then
            contadorFacturas = contadorFacturas + 1
            For Each linea In facturaData("lineas")
                Call InsertarDetalleVenta(facturaID, fechaVenta, _
                    CStr(linea(0)), CStr(linea(1)), _
                    ToDoubleOrZero(linea(2)), ToDoubleOrZero(linea(3)), _
                    ToDoubleOrZero(linea(4)))
                Call ActualizarStockProducto(CStr(linea(0)), ToDoubleOrZero(linea(2)))
            Next linea
        End If
    Next clave
    
    ' 5. Limpiar pendientes
    ultimaFila = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For filaExcel = ultimaFila To 2 Step -1
        ws.Rows(filaExcel).Delete
    Next filaExcel
    
    wb.Save
    wb.Close False
    
    ' 6. Limpiar log
    rutaLog = ThisWorkbook.Path & "\" & CARPETA_WEBAPP & ARCHIVO_VENTAS_LOG
    If Dir(rutaLog) <> "" Then Kill rutaLog
    
    Application.ScreenUpdating = True
    MsgBox "Importación completada: " & contadorFacturas & " facturas importadas.", vbInformation
    Exit Sub
    
ErrorHandler:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    On Error GoTo 0
    Application.ScreenUpdating = True
    MsgBox "Error al importar ventas: " & Err.Description, vbCritical
End Sub

' --- OBTENER VENTAS PENDIENTES DESDE datos.xlsx ----------------------------

Private Function ObtenerVentasPendientes() As Variant
    On Error GoTo ErrorHandler
    
    Dim rutaIntermediario As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim datos() As Variant
    Dim fila As Long
    Dim i As Long
    Dim ultimaFila As Long
    
    rutaIntermediario = ThisWorkbook.Path & "\" & CARPETA_WEBAPP & ARCHIVO_DATOS
    
    If Dir(rutaIntermediario) = "" Then
        ObtenerVentasPendientes = Array()
        Exit Function
    End If
    
    Set wb = Workbooks.Open(rutaIntermediario, ReadOnly:=False)
    Set ws = wb.Sheets("Pendientes")
    ultimaFila = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Contar pendientes (columna J = False)
    i = 0
    For fila = 2 To ultimaFila
        If ws.Cells(fila, 10).Value = False Then i = i + 1
    Next fila
    
    If i = 0 Then
        wb.Close False
        ObtenerVentasPendientes = Array()
        Exit Function
    End If
    
    ' Almacenar datos pendientes
    ReDim datos(1 To i, 1 To 10)
    i = 1
    For fila = 2 To ultimaFila
        If ws.Cells(fila, 10).Value = False Then
            datos(i, 1) = ws.Cells(fila, 1).Value   ' FacturaID
            datos(i, 2) = ws.Cells(fila, 2).Value   ' FechaHora
            datos(i, 3) = ws.Cells(fila, 3).Value   ' Codigo
            datos(i, 4) = ws.Cells(fila, 4).Value   ' Nombre
            datos(i, 5) = ws.Cells(fila, 5).Value   ' Cantidad
            datos(i, 6) = ws.Cells(fila, 6).Value   ' PrecioUnitario
            datos(i, 7) = ws.Cells(fila, 7).Value   ' Subtotal
            datos(i, 8) = ws.Cells(fila, 8).Value   ' Efectivo
            datos(i, 9) = ws.Cells(fila, 9).Value   ' Transferencia
            datos(i, 10) = ws.Cells(fila, 10).Value ' Procesado
            i = i + 1
        End If
    Next fila
    
    ObtenerVentasPendientes = Array(datos, wb)
    Exit Function
    
ErrorHandler:
    If Not wb Is Nothing Then wb.Close False
    ObtenerVentasPendientes = Array()
End Function

' --- IMPORTAR ENTRADAS DE PRODUCTOS ----------------------------------------

Private Sub ImportarEntradas()
    On Error GoTo ErrorHandler
    
    Dim wb As Workbook, wsEntrada As Worksheet, wsAlmacen As Worksheet
    Dim tbl As ListObject, fila As ListRow
    Dim ruta As String, ultimaFilaEntrada As Long, i As Long
    Dim codigo As String, nombre As String, fechaEntrada As Date
    Dim cantInicial As Double, precioVenta As Double, precioCosto As Double
    Dim cantActual As Double, fondo As Double
    Dim ingresoEsperado As Double, gananciaEsperada As Double
    
    ruta = ThisWorkbook.Path & "\" & CARPETA_WEBAPP & ARCHIVO_DATOS
    If Dir(ruta) = "" Then Exit Sub
    
    Set wb = Workbooks.Open(ruta, ReadOnly:=False)
    
    On Error Resume Next
    Set wsEntrada = wb.Sheets("Entrada")
    On Error GoTo 0
    
    If wsEntrada Is Nothing Then
        wb.Close False
        Exit Sub
    End If
    
    If wsEntrada.Cells(2, 1).Value = "" Then
        wb.Save
        wb.Close False
        Exit Sub
    End If
    
    Set wsAlmacen = ThisWorkbook.Sheets(HOJA_ALMACEN)
    Set tbl = wsAlmacen.ListObjects(TBL_INVENTARIO)
    ultimaFilaEntrada = wsEntrada.Cells(wsEntrada.Rows.Count, 1).End(xlUp).Row
    
    wsAlmacen.Unprotect password:=PASS_HOJA
    
    For i = 2 To ultimaFilaEntrada
        codigo = CStr(wsEntrada.Cells(i, 1).Value)
        nombre = CStr(wsEntrada.Cells(i, 2).Value)
        On Error Resume Next
        fechaEntrada = CDate(wsEntrada.Cells(i, 3).Value)
        On Error GoTo 0
        cantInicial = CDbl(wsEntrada.Cells(i, 4).Value)
        precioVenta = CDbl(wsEntrada.Cells(i, 5).Value)
        precioCosto = CDbl(wsEntrada.Cells(i, 6).Value)
        
        cantActual = cantInicial
        fondo = cantActual * precioCosto
        ingresoEsperado = cantActual * precioVenta
        gananciaEsperada = ingresoEsperado - fondo
        
        Set fila = tbl.ListRows.Add
        With fila.Range
            .Cells(1, COL_INV_CODIGO) = codigo
            .Cells(1, COL_INV_NOMBRE) = nombre
            .Cells(1, COL_INV_FECHA) = fechaEntrada
            .Cells(1, COL_INV_CANT_INI) = cantInicial
            .Cells(1, COL_INV_CANT_ACT) = cantActual
            .Cells(1, COL_INV_PRECIO_V) = precioVenta
            .Cells(1, COL_INV_PRECIO_C) = precioCosto
            .Cells(1, COL_INV_FONDO) = fondo
            .Cells(1, COL_INV_INGRESO) = ingresoEsperado
            .Cells(1, COL_INV_GANANCIA) = gananciaEsperada
        End With
    Next i
    
    Call modAgregarRegistros.AgregarProductoUnico(nombre)
    Call modAgregarRegistros.AgregarFecha(fechaEntrada)
    
    tbl.Sort.Apply
    wsAlmacen.Protect password:=PASS_HOJA, UserInterfaceOnly:=True
    
    ' Limpiar hoja Entrada
    For i = ultimaFilaEntrada To 2 Step -1
        If wsEntrada.Cells(i, 1).Value = "" Then Exit For
        wsEntrada.Rows(i).Delete
    Next i
    
    wb.Save
    wb.Close False
    
    Exit Sub
    
ErrorHandler:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    On Error GoTo 0
    MsgBox "Error al importar entradas: " & Err.Description, vbCritical
End Sub

' --- ACTUALIZAR INVENTARIO DESDE ABASTECIMIENTO ----------------------------

Private Sub ActualizarInventarioDesdeAbastecimiento()
    Dim wbGestion As Workbook, wbDatos As Workbook
    Dim wsGestion As Worksheet, wsAbastecimiento As Worksheet
    Dim rngInventario As Range, foundCell As Range
    Dim lastRow As Long, r As Long
    Dim codigo As Variant, Cantidad As Double, fecha As Date
    Dim datosPath As String
    
    Set wbGestion = ThisWorkbook
    
    datosPath = wbGestion.Path & "\" & CARPETA_WEBAPP & ARCHIVO_DATOS
    If Dir(datosPath) = "" Then Exit Sub
    
    Set wbDatos = Workbooks.Open(datosPath)
    Set wsAbastecimiento = wbDatos.Worksheets("Abastecimiento")
    Set wsGestion = wbGestion.Worksheets(HOJA_ALMACEN)
    Set rngInventario = wsGestion.ListObjects(TBL_INVENTARIO).DataBodyRange
    
    lastRow = wsAbastecimiento.Cells(wsAbastecimiento.Rows.Count, "A").End(xlUp).Row
    
    For r = 2 To lastRow
        codigo = wsAbastecimiento.Cells(r, 1).Value
        If codigo = "" Then GoTo NextRow
        
        Cantidad = Val(wsAbastecimiento.Cells(r, 2).Value)
        fecha = ParseFechaVenta(wsAbastecimiento.Cells(r, 3).Value)
        
        Set foundCell = rngInventario.Columns(1).Find(codigo, LookAt:=xlWhole)
        If Not foundCell Is Nothing Then
            foundCell.Offset(0, 2).Value = fecha
            foundCell.Offset(0, 3).Value = Val(foundCell.Offset(0, 3).Value) + Cantidad
        End If
NextRow:
    Next r
    
    If lastRow >= 2 Then wsAbastecimiento.Range("A2:C" & lastRow).ClearContents
    
    wbDatos.Save
    wbDatos.Close
End Sub

' --- IMPORTAR GASTOS VARIABLES ---------------------------------------------

Private Sub ImportarGastos()
    On Error GoTo ErrorHandler
    
    Dim wb As Workbook, wsGastos As Worksheet
    Dim ruta As String, ultimaFila As Long, i As Long
    Dim fecha As Date, descripcion As String, monto As Double
    
    ruta = ThisWorkbook.Path & "\" & CARPETA_WEBAPP & ARCHIVO_DATOS
    If Dir(ruta) = "" Then Exit Sub
    
    Set wb = Workbooks.Open(ruta, ReadOnly:=False)
    
    On Error Resume Next
    Set wsGastos = wb.Sheets("Gastos")
    On Error GoTo 0
    
    If wsGastos Is Nothing Then
        wb.Close False
        Exit Sub
    End If
    
    If wsGastos.Cells(2, 1).Value = "" Then
        wb.Save
        wb.Close False
        Exit Sub
    End If
    
    ultimaFila = wsGastos.Cells(wsGastos.Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To ultimaFila
        fecha = CDate(wsGastos.Cells(i, 1).Value)
        descripcion = CStr(wsGastos.Cells(i, 2).Value)
        monto = CDbl(wsGastos.Cells(i, 3).Value)
        
        If descripcion <> "" And monto > 0 Then
            Call AgregarGasto(fecha, CAT_VARIABLE, descripcion, monto)
            Call modAgregarRegistros.AgregarFecha(fecha)
        End If
    Next i
    
    ' Limpiar hoja Gastos
    For i = ultimaFila To 2 Step -1
        If wsGastos.Cells(i, 1).Value = "" Then Exit For
        wsGastos.Rows(i).Delete
    Next i
    
    wb.Save
    wb.Close False
    Exit Sub
    
ErrorHandler:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    On Error GoTo 0
End Sub

' --- PARSEAR FECHA DE VENTA (ISO o Date) -----------------------------------

Private Function ParseFechaVenta(ByVal v As Variant) As Date
    On Error GoTo ErrorHandler
    
    Dim s As String, partes As Variant, fechaPart As String
    Dim fechaSeg As Variant, anio As Integer, mes As Integer, dia As Integer
    
    If IsDate(v) Then
        ParseFechaVenta = DateValue(v)
        Exit Function
    End If
    
    If IsNull(v) Or IsEmpty(v) Then
        Err.Raise vbObjectError + 3101, "ParseFechaVenta", "Fecha vacía."
    End If
    
    s = Trim$(CStr(v))
    If Len(s) = 0 Then Err.Raise vbObjectError + 3102, "ParseFechaVenta", "Fecha vacía."
    
    ' Normalizar ISO
    s = Replace$(s, "T", " ")
    If Right$(s, 1) = "Z" Then s = Left$(s, Len(s) - 1)
    
    partes = Split(s, " ")
    fechaPart = partes(0)
    fechaSeg = Split(fechaPart, "-")
    
    anio = CInt(fechaSeg(0))
    mes = CInt(fechaSeg(1))
    dia = CInt(fechaSeg(2))
    
    ParseFechaVenta = DateSerial(anio, mes, dia)
    Exit Function
    
ErrorHandler:
    Err.Raise vbObjectError + 3199, "ParseFechaVenta", "No se pudo convertir la fecha: " & CStr(v)
End Function

' --- ACTIVAR POS MÓVIL (macro principal desde botón) -----------------------

Public Sub ActivarPOSMovil()
    ' 1. Verificar si ya está activo
    If ServidorEstaActivo() Then
        GoTo MostrarFormulario
    End If
    
    ' 2. Validar carpeta
    If Dir(ThisWorkbook.Path & "\" & CARPETA_WEBAPP, vbDirectory) = "" Then
        MsgBox "No se encuentra la carpeta webapp-beta", vbCritical
        Exit Sub
    End If
    
    ' 3. Copiar última factura
    Call CopiarUltimaFacturaDelDia
    
    ' 4. Iniciar servidor
    If Not IniciarServidor() Then
        MsgBox "No se pudo iniciar el servidor. Verifica que Node.js esté instalado.", vbCritical
        Exit Sub
    End If
    
    MsgBox "POS Móvil activado correctamente", vbInformation
    
MostrarFormulario:
    Call SincronizarProductosEnDatos
    Call SincronizarRecetasEnDatos
    Call OcultarInterfaz
    frmPOSMovil.Show vbModeless
End Sub

Private Sub ProcesarElaboraciones()
    On Error GoTo ErrorHandler
    
    Dim wb As Workbook
    Dim wsElaboracion As Worksheet
    Dim ruta As String
    Dim ultimaFila As Long
    Dim i As Long
    Dim nombreReceta As String
    Dim Lotes As Double
    Dim faltantes As String
    Dim procesadas As Long
    Dim pendientes As Long
    Dim mensajePendientes As String
    
    ruta = ThisWorkbook.Path & "\" & CARPETA_WEBAPP & ARCHIVO_DATOS
    If Dir(ruta) = "" Then Exit Sub
    
    Set wb = Workbooks.Open(ruta, ReadOnly:=False)
    
    On Error Resume Next
    Set wsElaboracion = wb.Sheets("Elaboracion")
    On Error GoTo ErrorHandler
    
    If wsElaboracion Is Nothing Then
        wb.Close False
        Exit Sub
    End If
    
    If wsElaboracion.Cells(2, 1).Value = "" Then
        wb.Save
        wb.Close False
        Exit Sub
    End If
    
    ultimaFila = wsElaboracion.Cells(wsElaboracion.Rows.Count, 1).End(xlUp).Row
    procesadas = 0
    pendientes = 0
    mensajePendientes = ""
    
    ' Iterar de abajo hacia arriba para borrar filas sin problemas
    For i = ultimaFila To 2 Step -1
        nombreReceta = CStr(wsElaboracion.Cells(i, 1).Value)
        Lotes = Val(wsElaboracion.Cells(i, 2).Value)
        
        If nombreReceta = "" Or Lotes <= 0 Then
            ' Fila inválida, borrarla igual
            wsElaboracion.Rows(i).Delete
            GoTo Siguiente
        End If
        
        ' 1. VALIDAR si hay stock suficiente de ingredientes
        faltantes = ""
        On Error Resume Next
        If modElaboracion.ValidarStockElaboracion(nombreReceta, Lotes, faltantes) Then
            On Error GoTo ErrorHandler
            
            ' 2. Stock suficiente ? descontar y borrar la fila
            Call modElaboracion.DescontarStockIngredientes(nombreReceta, Lotes)
            wsElaboracion.Rows(i).Delete
            procesadas = procesadas + 1
        Else
            On Error GoTo ErrorHandler
            
            ' 3. Stock insuficiente ? dejar la fila para después
            pendientes = pendientes + 1
            mensajePendientes = mensajePendientes & _
                "- " & nombreReceta & " x" & Lotes & " lote(s)" & vbCrLf & _
                "  " & Replace(faltantes, vbCrLf, vbCrLf & "  ") & vbCrLf
        End If
        
Siguiente:
    Next i
    
    wb.Save
    wb.Close False
    
    ' Mostrar resumen al usuario
    If procesadas > 0 And pendientes = 0 Then
        MsgBox "Elaboraciones procesadas correctamente: " & procesadas, vbInformation
    ElseIf procesadas > 0 And pendientes > 0 Then
        MsgBox "Elaboraciones procesadas: " & procesadas & vbCrLf & vbCrLf & _
               "QUEDARON PENDIENTES (" & pendientes & ") por stock insuficiente:" & vbCrLf & _
               vbCrLf & mensajePendientes & vbCrLf & _
               "Agregá los ingredientes faltantes y volvé a importar.", _
               vbExclamation
    ElseIf procesadas = 0 And pendientes > 0 Then
        MsgBox "NINGUNA elaboración pudo procesarse por stock insuficiente:" & vbCrLf & _
               vbCrLf & mensajePendientes & vbCrLf & _
               "Agregá los ingredientes faltantes y volvé a importar.", _
               vbExclamation
    End If
    
    Exit Sub
    
ErrorHandler:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    On Error GoTo 0
End Sub

