Attribute VB_Name = "modGeneradorCodigo"
' ============================================================================
' MÓDULO: modGeneradorCodigo
' PROPÓSITO: Generación de códigos de producto en formato "Pr_00001".
' ============================================================================
' CAMBIO: contador persistente en la hoja oculta "Hidden" (celda B1).
'
'   GenerarCodigoProducto   -> SOLO sugiere el próximo número. NO toca el contador.
'   ConfirmarCodigoProducto -> lo llama el guardado; sube el contador al número
'                              realmente usado.
'   SincronizarContador     -> opcional, después de importar productos.
'
'   Reglas:
'   - El contador NUNCA retrocede: borrar el producto de código más alto no
'     libera el número, y vaciar el almacén no reinicia en Pr_00001.
'   - Contador := MAX(contador, número usado, máximo real de la tabla), así los
'     códigos que entran por importación / POS móvil / respaldo quedan cubiertos.
'   - Cancelar el formulario ya no quema números: el contador solo se mueve al
'     guardar.
' ============================================================================
' Celdas ocupadas de la hoja oculta:
'   A1 = fecha de expiración     (modLicencia)
'   A2 = última fecha usada      (modLicencia)
'   A4 = contraseña de admin     (modSeguridad)
'   B1 = último número usado     <-- ESTE MÓDULO
' ============================================================================

Option Explicit

Private Const CELDA_ULTIMO_CODIGO As String = "B1"

' --- 1. GENERAR (solo sugiere, NO persiste) --------------------------------

Public Sub GenerarCodigoProducto()
    Codigo_Producto = FormatearCodigo(SiguienteNumeroDisponible(TablaInventario()))
End Sub

' --- 2. CONFIRMAR (llamar SIEMPRE después de escribir la fila) -------------

Public Sub ConfirmarCodigoProducto(ByVal codigo As String)
    AjustarContador ExtraerNumero(codigo)
End Sub

' --- 3. SINCRONIZAR (opcional, después de importar productos) --------------

Public Sub SincronizarContador()
    AjustarContador 0
End Sub

' --- ¿El código sigue libre? (chequeo previo al guardado) ------------------

Public Function CodigoDisponible(ByVal codigo As String) As Boolean
    CodigoDisponible = Not ExisteCodigo(TablaInventario(), codigo)
End Function

' --- LÓGICA INTERNA --------------------------------------------------------

' Sube el contador a MAX(contador, numeroMinimo, máximo real de la tabla).
' Nunca lo baja.
Private Sub AjustarContador(ByVal numeroMinimo As Long)
    Dim nuevo As Long
    Dim maxTabla As Long
    
    nuevo = LeerUltimoNumero()
    If numeroMinimo > nuevo Then nuevo = numeroMinimo
    
    maxTabla = ObtenerMaximoNumeroEnTabla(TablaInventario())
    If maxTabla > nuevo Then nuevo = maxTabla
    
    If nuevo <= 0 Then Exit Sub
    
    If Not GuardarUltimoNumero(nuevo) Then
        MsgBox "No se pudo actualizar el contador de códigos en la hoja oculta """ & _
               HOJA_HIDDEN & """ (celda " & CELDA_ULTIMO_CODIGO & ")." & vbCrLf & vbCrLf & _
               "El producto quedó guardado, pero el próximo código podría repetirse." & vbCrLf & _
               "Revise la protección de la hoja.", _
               vbExclamation, "Generador de códigos"
    End If
End Sub

' Próximo número libre: MAYOR entre contador y máximo real de la tabla, más
' uno. Salta cualquier código que ya exista.
Private Function SiguienteNumeroDisponible(ByVal tbl As ListObject) As Long
    Dim numeroGuardado As Long
    Dim numeroEnTabla As Long
    Dim candidato As Long
    
    numeroGuardado = LeerUltimoNumero()
    numeroEnTabla = ObtenerMaximoNumeroEnTabla(tbl)
    
    If numeroEnTabla > numeroGuardado Then numeroGuardado = numeroEnTabla
    
    candidato = numeroGuardado + 1
    If candidato < 1 Then candidato = 1
    
    Do While ExisteCodigo(tbl, FormatearCodigo(candidato))
        candidato = candidato + 1
    Loop
    
    SiguienteNumeroDisponible = candidato
End Function

Private Function FormatearCodigo(ByVal numero As Long) As String
    FormatearCodigo = "Pr_" & Format$(numero, "00000")
End Function

Private Function TablaInventario() As ListObject
    Set TablaInventario = ThisWorkbook.Sheets(HOJA_ALMACEN).ListObjects(TBL_INVENTARIO)
End Function

' Mayor número "Pr_#####" presente en la tabla de inventario.
Private Function ObtenerMaximoNumeroEnTabla(ByVal tbl As ListObject) As Long
    Dim i As Long
    Dim numero As Long
    Dim maximo As Long
    Dim celda As Variant
    
    If tbl Is Nothing Then Exit Function
    If tbl.DataBodyRange Is Nothing Then Exit Function
    
    For i = 1 To tbl.ListRows.Count
        celda = tbl.ListRows(i).Range.Cells(1, COL_INV_CODIGO).Value
        numero = ExtraerNumero(ValorTexto(celda))
        If numero > maximo Then maximo = numero
    Next i
    
    ObtenerMaximoNumeroEnTabla = maximo
End Function

' Extrae el número de un código "Pr_#####". Devuelve 0 si no cumple el
' formato, para que un código escrito a mano no altere el contador.
Private Function ExtraerNumero(ByVal codigo As String) As Long
    Dim digitos As String
    Dim i As Long
    Dim caracter As String
    
    codigo = Trim$(codigo)
    If Len(codigo) < 4 Then Exit Function
    If UCase$(Left$(codigo, 3)) <> "PR_" Then Exit Function
    
    digitos = Mid$(codigo, 4)
    For i = 1 To Len(digitos)
        caracter = Mid$(digitos, i, 1)
        If caracter < "0" Or caracter > "9" Then Exit Function
    Next i
    
    If Len(digitos) = 0 Then Exit Function
    If Len(digitos) > 9 Then Exit Function    ' evita desbordamiento de Long
    
    ExtraerNumero = CLng(digitos)
End Function

' ¿Ya existe ese código en el inventario?
Private Function ExisteCodigo(ByVal tbl As ListObject, ByVal codigo As String) As Boolean
    Dim rango As Range
    Dim encontrado As Range
    
    If tbl Is Nothing Then Exit Function
    
    On Error Resume Next
    Set rango = tbl.ListColumns(COL_INV_CODIGO).DataBodyRange
    On Error GoTo 0
    
    If rango Is Nothing Then Exit Function
    
    Set encontrado = rango.Find(What:=codigo, LookIn:=xlValues, LookAt:=xlWhole)
    ExisteCodigo = Not (encontrado Is Nothing)
End Function

' Convierte a texto sin explotar con celdas con error o Null.
Private Function ValorTexto(ByVal v As Variant) As String
    If IsError(v) Then Exit Function
    If IsNull(v) Then Exit Function
    ValorTexto = CStr(v)
End Function

' Lee el contador de la hoja oculta. Devuelve 0 si está vacío o inaccesible.
Private Function LeerUltimoNumero() As Long
    Dim valor As Variant
    
    On Error GoTo Fallo
    valor = ThisWorkbook.Sheets(HOJA_HIDDEN).Range(CELDA_ULTIMO_CODIGO).Value
    If IsNumeric(valor) Then LeerUltimoNumero = CLng(valor)
    Exit Function

Fallo:
    LeerUltimoNumero = 0
End Function

' Guarda el contador en la hoja oculta. True si se pudo escribir.
Private Function GuardarUltimoNumero(ByVal numero As Long) As Boolean
    Dim ws As Worksheet
    
    On Error GoTo Fallo
    
    Set ws = ThisWorkbook.Sheets(HOJA_HIDDEN)
    modSeguridad.DesprotegerHoja ws
    ws.Range(CELDA_ULTIMO_CODIGO).NumberFormat = "0"
    ws.Range(CELDA_ULTIMO_CODIGO).Value = numero
    modSeguridad.ProtegerHoja ws
    
    GuardarUltimoNumero = True
    Exit Function

Fallo:
    On Error Resume Next
    If Not ws Is Nothing Then modSeguridad.ProtegerHoja ws
    GuardarUltimoNumero = False
End Function

