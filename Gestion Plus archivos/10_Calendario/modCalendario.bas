Attribute VB_Name = "modCalendario"
' ============================================================================
' MODULO: modCalendario
' PROPOSITO: Logica del formulario calendario. Controla la navegacion de
' meses/años, carga de dias, marcado/desmarcado y envio de fecha a los
' formularios segun seleccion_destino_calendario.
' Reemplaza a ModuloCalendario.bas.
' ============================================================================
' NOTA: Mantiene la misma interfaz publica que ModuloCalendario.bas para no
' romper los eventos de frmCalendario.frm (31 labels con DblClick/MouseMove).
' ============================================================================
Option Explicit
Option Private Module
Private SenalCambioMes As Long
' --- RECIBIR LA FECHA SELECCIONADA EN EL CALENDARIO -------------------------
Public Sub RecibeLaFecha(dia As Long, mes As Long, Ano As Long)
    Dim FechaRecibida As Date
    FechaRecibida = VBA.DateSerial(VBA.CInt(Ano), VBA.CInt(mes), VBA.CInt(dia))
    Select Case seleccion_destino_calendario
        Case 1
            ingresar_form.fecha_de_entrada_TextBox.Value = Format(FechaRecibida, "d-mmm-yyyy")
        Case 2
            facturacion.fecha_de_orden_TextBox.Value = Format(FechaRecibida, "d-mmm-yyyy")
        Case 3
            Busqueda.Txt_Fecha.Value = Format(FechaRecibida, "d-mmm-yyyy")
            If CDate(Busqueda.Txt_Fecha.Value) > CDate(Busqueda.Txt_Fecha_Final.Value) Then
                Busqueda.Txt_Fecha_Final.Value = Format(FechaRecibida, "d-mmm-yyyy")
            End If
        Case 4
            Busqueda.Txt_Fecha_Final.Value = Format(FechaRecibida, "d-mmm-yyyy")
            If CDate(Busqueda.Txt_Fecha.Value) > CDate(Busqueda.Txt_Fecha_Final.Value) Then
                Busqueda.Txt_Fecha.Value = Format(FechaRecibida, "d-mmm-yyyy")
            End If
        Case 5
            Busqueda_Ventas.Txt_Fecha.Value = Format(FechaRecibida, "d-mmm-yyyy")
            If CDate(Busqueda_Ventas.Txt_Fecha.Value) > CDate(Busqueda_Ventas.Txt_Fecha_Final.Value) Then
                Busqueda_Ventas.Txt_Fecha_Final.Value = Format(FechaRecibida, "d-mmm-yyyy")
            End If
        Case 6
            Busqueda_Ventas.Txt_Fecha_Final.Value = Format(FechaRecibida, "d-mmm-yyyy")
            If CDate(Busqueda_Ventas.Txt_Fecha.Value) > CDate(Busqueda_Ventas.Txt_Fecha_Final.Value) Then
                Busqueda_Ventas.Txt_Fecha.Value = Format(FechaRecibida, "d-mmm-yyyy")
            End If
        Case 7
            Busqueda_Facturacion.Txt_Fecha.Value = Format(FechaRecibida, "d-mmm-yyyy")
            If CDate(Busqueda_Facturacion.Txt_Fecha.Value) > CDate(Busqueda_Facturacion.Txt_Fecha_Final.Value) Then
                Busqueda_Facturacion.Txt_Fecha_Final.Value = Format(FechaRecibida, "d-mmm-yyyy")
            End If
        Case 8
            Busqueda_Facturacion.Txt_Fecha_Final.Value = Format(FechaRecibida, "d-mmm-yyyy")
            If CDate(Busqueda_Facturacion.Txt_Fecha.Value) > CDate(Busqueda_Facturacion.Txt_Fecha_Final.Value) Then
                Busqueda_Facturacion.Txt_Fecha.Value = Format(FechaRecibida, "d-mmm-yyyy")
            End If
    End Select
End Sub
' --- INICIALIZAR FORMULARIO CALENDARIO --------------------------------------
Public Sub InicializaFormularioCalendario()
    SenalCambioMes = 1
    With frmCalendario.cboMes
        .AddItem 1: .List(0, 1) = "enero"
        .AddItem 2: .List(1, 1) = "febrero"
        .AddItem 3: .List(2, 1) = "marzo"
        .AddItem 4: .List(3, 1) = "abril"
        .AddItem 5: .List(4, 1) = "mayo"
        .AddItem 6: .List(5, 1) = "junio"
        .AddItem 7: .List(6, 1) = "julio"
        .AddItem 8: .List(7, 1) = "agosto"
        .AddItem 9: .List(8, 1) = "septiembre"
        .AddItem 10: .List(9, 1) = "octubre"
        .AddItem 11: .List(10, 1) = "noviembre"
        .AddItem 12: .List(11, 1) = "diciembre"
    End With
    frmCalendario.cboMes.ListIndex = VBA.Month(VBA.Date) - 1
    frmCalendario.spbAño.Value = VBA.Year(VBA.Date)
    frmCalendario.lblAno.Caption = VBA.Year(VBA.Date)
    Dim Ano As Long, mes As Long
    Ano = VBA.Year(VBA.Date)
    mes = VBA.Month(VBA.Date)
    Call CargarLosDias(Ano, mes)
    frmCalendario.lblHoy.Caption = Format(VBA.Date, "d-mmm-yyyy")
End Sub
' --- CARGAR LOS DIAS DEL MES ------------------------------------------------
Public Sub CargarLosDias(Ano As Long, mes As Long)
    Dim FechaDelPrimerDia As Date, FechaDelUltimoDia As Date
    Dim DiaSemanaPrimerDia As Long, VariableControl As Control, contador As Long
    FechaDelPrimerDia = VBA.DateSerial(Ano, mes, 1)
    FechaDelUltimoDia = Application.WorksheetFunction.EoMonth(VBA.DateSerial(Ano, mes, 1), 0)
    DiaSemanaPrimerDia = Application.WorksheetFunction.Weekday(FechaDelPrimerDia, 2)
    contador = 1
    For Each VariableControl In frmCalendario.mrcDias.Controls
        VariableControl.Caption = "-"
        If VariableControl.Tag >= DiaSemanaPrimerDia And contador <= VBA.Day(FechaDelUltimoDia) Then
            VariableControl.Caption = contador
            contador = contador + 1
        End If
    Next VariableControl
End Sub
' --- CAMBIO DE MES ----------------------------------------------------------
Public Sub CambioDeMes()
    If SenalCambioMes > 1 Then
        Dim MesEnElCombo As Long, AnoEnElLabel As Long
        If Not IsNull(frmCalendario.cboMes.Value) And Not IsNull(frmCalendario.lblAno.Caption) Then
            MesEnElCombo = VBA.CLng(frmCalendario.cboMes.Value)
            AnoEnElLabel = VBA.CLng(frmCalendario.lblAno.Caption)
            Call DesmarcarDias: Call CargarLosDias(AnoEnElLabel, MesEnElCombo)
        End If
    End If
    SenalCambioMes = SenalCambioMes + 1
End Sub
' --- CAMBIO DE AñO ----------------------------------------------------------
Public Sub CambioDeAno()
    Dim MesEnElCombo As Long, AnoEnElLabel As Long
    frmCalendario.lblAno.Caption = frmCalendario.spbAño.Value
    MesEnElCombo = VBA.CLng(frmCalendario.cboMes.Value)
    AnoEnElLabel = VBA.CLng(frmCalendario.lblAno.Caption)
    Call DesmarcarDias: Call CargarLosDias(AnoEnElLabel, MesEnElCombo)
End Sub
' --- VOLVER AL DIA DE HOY ---------------------------------------------------
Public Sub UnClickEnHoyEs()
    Dim mes As Long, Ano As Long, FechaActual As Date
    FechaActual = VBA.CDate(frmCalendario.lblHoy.Caption)
    mes = VBA.CLng(VBA.Month(FechaActual))
    Ano = VBA.CLng(VBA.Year(FechaActual))
    frmCalendario.lblAno.Caption = Ano
    frmCalendario.cboMes.ListIndex = mes - 1
    frmCalendario.spbAño.Value = Ano
    frmCalendario.spbAño.SetFocus
    Call DesmarcarDias: Call CargarLosDias(Ano, mes)
End Sub
' --- SALIR CON TECLA ESCAPE -------------------------------------------------
Sub SalirConEscape()
    Unload frmCalendario
End Sub
' --- MARCAR DIA -------------------------------------------------------------
Sub MarcarDia(ControlDeEtiqueta As Control)
    Call DesmarcarDias
    ControlDeEtiqueta.Font.Bold = True
    ControlDeEtiqueta.ForeColor = VBA.RGB(255, 0, 0)
End Sub
' --- DESMARCAR TODOS LOS DIAS -----------------------------------------------
Sub DesmarcarDias()
    Dim ControlEtiqueta As Control
    For Each ControlEtiqueta In frmCalendario.mrcDias.Controls
        ControlEtiqueta.Font.Bold = False
        ControlEtiqueta.ForeColor = VBA.RGB(0, 0, 0)
    Next ControlEtiqueta
End Sub
