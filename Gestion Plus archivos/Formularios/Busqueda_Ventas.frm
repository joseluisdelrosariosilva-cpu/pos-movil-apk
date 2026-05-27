VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Busqueda_Ventas 
   Caption         =   "Buscar"
   ClientHeight    =   5410
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   16440
   OleObjectBlob   =   "Busqueda_Ventas.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Busqueda_Ventas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

' ============================================================================
' FORMULARIO: Busqueda_Ventas
' PROPÓSITO:  Búsqueda de ventas por rango de fechas y descripción.
'              Toda la lógica de negocio delega en modBusquedas.
' ============================================================================

Private filasOriginales As Collection

' --- BÚSQUEDA PRINCIPAL ------------------------------------------------------
Private Sub Btn_Active_Click()
    Dim fechaInicio As Date
    Dim fechaFin As Date
    Dim descripcion As String
    Dim resultados As Variant
    
    fechaInicio = CDate(Trim(Me.Txt_Fecha))
    fechaFin = CDate(Trim(Me.Txt_Fecha_Final))
    
    If Me.Txt_Descripcion.ListIndex > 0 Then
        descripcion = Trim(Me.Txt_Descripcion.Value)
    Else
        descripcion = ""
    End If
    
    Set filasOriginales = New Collection
    resultados = modBusquedas.BuscarVentas(fechaInicio, fechaFin, descripcion, filasOriginales)
    
    If IsArray(resultados) And (Not IsEmpty(resultados)) Then
        If UBound(resultados, 1) > 0 Then
            Me.lstResultados.List = resultados
            Exit Sub
        End If
    End If
    
    Me.lstResultados.Clear
    MsgBox "No se encontraron ventas para el período seleccionado.", vbInformation, "Sin resultados"
End Sub

' --- DOBLE CLIC: NAVEGAR A LA FILA ORIGINAL ----------------------------------
Private Sub lstResultados_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim filaSeleccionada As Long
    Dim numFilaOriginal As Long
    
    If Me.lstResultados.ListIndex = -1 Then Exit Sub
    
    filaSeleccionada = Me.lstResultados.ListIndex + 1
    numFilaOriginal = filasOriginales(CStr(filaSeleccionada))
    
    Call modBusquedas.NavegarAFilaOriginal(numFilaOriginal, HOJA_VENTAS)
    Unload Me
End Sub

' --- INICIALIZACIÓN ----------------------------------------------------------
Private Sub UserForm_Initialize()
    Dim ws As Worksheet
    Dim tbl As ListObject
    
    Set filasOriginales = New Collection
    
    With Me.lstResultados
        .ColumnCount = 6
        .ColumnWidths = "95;80;80;110;120;120"
    End With
    
    Me.Txt_Fecha = Format(Date, "d-mmm-yyyy")
    Me.Txt_Fecha_Final = Format(Date, "d-mmm-yyyy")
    
    Set ws = ThisWorkbook.Sheets(HOJA_VENTAS)
    Set tbl = ws.ListObjects(TBL_VENTAS)
    Call modBusquedas.CargarDescripcionesUnicas(tbl, 2, Me.Txt_Descripcion)
End Sub

' --- CALENDARIO ---------------------------------------------------------------
Private Sub boton_calendario_covered_Click()
    seleccion_destino_calendario = 5
    frmCalendario.Show
End Sub

Private Sub boton_calendario_covered2_Click()
    seleccion_destino_calendario = 6
    frmCalendario.Show
End Sub

' --- EVENTOS MOUSE MOVE -------------------------------------------------------
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.Btn_Reposo.ZOrder msoBringToFront
    Me.boton_calendario_reposo.ZOrder msoBringToFront
    Me.boton_calendario_reposo2.ZOrder msoBringToFront
End Sub

Private Sub Btn_Reposo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.Btn_Active.ZOrder msoBringToFront
End Sub

Private Sub boton_calendario_reposo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.boton_calendario_covered.ZOrder msoBringToFront
End Sub

Private Sub boton_calendario_reposo2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.boton_calendario_covered2.ZOrder msoBringToFront
End Sub
