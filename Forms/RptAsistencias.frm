VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form RptAsistencias 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6465
   Icon            =   "RptAsistencias.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   6465
   Begin CRVIEWER9LibCtl.CRViewer9 CRViewer91 
      Height          =   3945
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6390
      lastProp        =   500
      _cx             =   11271
      _cy             =   6959
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
   End
End
Attribute VB_Name = "RptAsistencias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Report As New CrystalReport4
Dim report3 As New CrystalReport3 'Reporte Prima-AFP
Dim report2 As New CrystalReport5
Dim report4 As New CrystalReport6 'Reporte Afp Horizonte
Dim report5 As New CrystalReport7 'Reporte Afp Integra
Dim report6 As New CrystalReport8 'Reporte Profuturo
Dim report7 As New CrystalReport9 'Comprobante de Retenciones
Dim report8 As New CrystalReport10 'Detalle Promedios
Dim rs_rptasistencias As ADODB.Recordset
Dim s_RucEmpresa As String
Dim s_Tipo_Trabajador As String
Public tdFecIni As Date

Private Sub CRViewer91_PrintButtonClicked(UseDefault As Boolean)
    Select Case i_Direccion_Reportes
        Case 1: Report.PrinterSetup 0
        Case 2: report2.PrinterSetup 0
        Case 3: report3.PrinterSetup 0
        Case 4: report4.PrinterSetup 0
        Case 5: report5.PrinterSetup 0
        Case 6: report6.PrinterSetup 0
        Case 7: report7.PrinterSetup 0
        Case 8: report8.PrinterSetup 0
    End Select
End Sub
Private Sub Form_Activate()
    Select Case i_Direccion_Reportes
        Case 1
            Call Llena_Informacion_Empresa
            Call Llena_Datos_Fecha(CStr(tdFecIni))
            Call Llena_Informacion_Adicional
            Call Titulo_General
            
            Report.DiscardSavedData
    
            Report.Database.Tables(1).ConnectionProperties("Password") = "USURPT"
            Report.Database.Tables(1).SetDataSource rs

            
            Report.SQLQueryString = "select Nombres,DNI from Reporte_Pla1 order by dias,orden"
            Screen.MousePointer = vbHourglass
            CRViewer91.ReportSource = Report
            CRViewer91.ViewReport
            Screen.MousePointer = vbDefault
        Case 2
            Call Llena_Informacion_Empresa
            Call Recupera_tipo_trabajador
            Call Titulo_General
                        
            report2.DiscardSavedData
    
            report2.Database.Tables(1).ConnectionProperties("Password") = "USURPT"
            report2.Database.Tables(1).SetDataSource rs
            
            report2.SQLQueryString = "select Codigo,Nombres,Fecha_Ingreso,Dni,Cargo,Direccion," & _
            "NombreAFP,CUSPP,Basico,AsigFamiliar from Reporte_Pla2 where TipoTrab='" & s_Tipo_Trabajador & "'"
            
            Screen.MousePointer = vbHourglass
            CRViewer91.ReportSource = report2
            CRViewer91.ViewReport
            Screen.MousePointer = vbDefault
        Case 3
            Call Enviar_Informacion_Proceso
            Call Enviar_Totales_Segundo_Nivel
            Screen.MousePointer = vbHourglass
            CRViewer91.ReportSource = report3
            CRViewer91.ViewReport
            Screen.MousePointer = vbDefault
        Case 4
            Call Recupera_Ruc
            Call Enviar_Ruc_Empresa_Reporte
            Call Enviar_Informacion_Proceso
            Call Enviar_Totales_Segundo_Nivel
            Screen.MousePointer = vbHourglass
            CRViewer91.ReportSource = report4
            CRViewer91.ViewReport
            Screen.MousePointer = vbDefault
        Case 5
            Call Enviar_Totales_Segundo_Nivel
            Screen.MousePointer = vbHourglass
            CRViewer91.ReportSource = report5
            CRViewer91.ViewReport
            Screen.MousePointer = vbDefault
        Case 6
            Call Enviar_Totales_Segundo_Nivel
            Screen.MousePointer = vbHourglass
            CRViewer91.ReportSource = report6
            CRViewer91.ViewReport
            Screen.MousePointer = vbDefault
        Case 7
            Call Llena_Informacion_Empresa
            Call Envia_Informacion_Comprobante_Retencion
            Screen.MousePointer = vbHourglass
            CRViewer91.ReportSource = report7
            CRViewer91.ViewReport
            Screen.MousePointer = vbDefault
        Case 8
            Call Envia_Informacion_Meses_Reporte_Promedios
            
            report8.DiscardSavedData
    
            report8.Database.Tables(1).ConnectionProperties("Password") = "USURPT"
            report8.Database.Tables(1).SetDataSource rs
            
            report8.SQLQueryString = "select placod,nombre,descripcion," & _
            "Mes1,Mes2,Mes3,Mes4,Mes5,Mes6 from promedios_Maestra2 " & _
            "order by placod "
            Screen.MousePointer = vbHourglass
            CRViewer91.ReportSource = report8
            CRViewer91.ViewReport
            Screen.MousePointer = vbDefault
    End Select
End Sub
Private Sub Form_Load()
    Me.Top = 1000: Me.Left = 1000
End Sub
Private Sub Form_Resize()
    CRViewer91.Top = 0
    CRViewer91.Left = 0
    CRViewer91.Height = ScaleHeight
    CRViewer91.Width = ScaleWidth
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Select Case i_Direccion_Reportes
        Case 1: Call Elimina_Informacion_Reporte_Asistencia
        Case 2: Call Eliminar_Tabla_Lista_Trabajadores
        Case 3: Call Elimina_Informacion_Reporte_Aportes
        Case 4: Call Elimina_Informacion_Reporte_Aportes
        Case 5: Call Elimina_Informacion_Reporte_Aportes
        Case 6: Call Elimina_Informacion_Reporte_Aportes
    End Select
End Sub
Sub Llena_Informacion_Empresa()
    Call Recupera_Informacion_Empresa(wcia)
    Set rs_rptasistencias = Reportes_Centrales.rs_RptCentrales_pub
    Set Reportes_Centrales.rs_RptCentrales_pub = Nothing
    Select Case i_Direccion_Reportes
        Case 1
            Report.Text20.SetText (rs_rptasistencias!razsoc)
            Report.Text21.SetText (rs_rptasistencias!RUC)
        Case 2
            report2.Text13.SetText (rs_rptasistencias!razsoc)
        Case 7
            report7.Text7.SetText (rs_rptasistencias!razsoc)
            report7.Text9.SetText (rs_rptasistencias!RUC)
    End Select
    Set rs_rptasistencias = Nothing
End Sub
Sub Llena_Datos_Fecha(Optional pFecha As String)
    If Len(Trim(pFecha)) > 0 Then
        Report.Text30.SetText (Day(CDate(pFecha))): Report.Text31.SetText (Month(CDate(pFecha)))
        Report.Text32.SetText (Year(CDate(pFecha)))
    Else
    Report.Text30.SetText (Day(Date)): Report.Text31.SetText (Month(Date))
    Report.Text32.SetText (Year(Date))
    End If
End Sub
Sub Llena_Informacion_Adicional()
    Call Recupera_Informacion_RegistroExistencia(wcia)
    Set rs_rptasistencias = Reportes_Centrales.rs_RptCentrales_pub
    Set Reportes_Centrales.rs_RptCentrales_pub = Nothing
    Report.Text25.SetText (rs_rptasistencias!Pagina)
    Report.Text22.SetText (rs_rptasistencias!Horario)
    Set rs_rptasistencias = Nothing
End Sub
Sub Titulo_General()
    Select Case i_Direccion_Reportes
        Case 1: Report.Text17.SetText ("LIBRO CONTROL DE ASISTENCIA " & FrmRptSeleccion.CmbTip.Text & "S")
        Case 2: report2.Text11.SetText ("DATOS GENERALES " & FrmRptseleccion2.CmbTip.Text & "S")
    End Select
End Sub
Sub Recupera_tipo_trabajador()
    Select Case FrmRptseleccion2.CmbTip.Text
        Case "EMPLEADO": s_Tipo_Trabajador = "01"
        Case "OBRERO": s_Tipo_Trabajador = "02"
    End Select
End Sub
Sub Enviar_Informacion_Proceso()
    Select Case i_Direccion_Reportes
        Case 3
            report3.Text6.SetText ("" & Mid(FrmRptAportes.s_MesSeleccionRpt, 1, 1) & "")
            report3.Text5.SetText ("" & Mid(FrmRptAportes.s_MesSeleccionRpt, 2, 1) & "")
            report3.Text4.SetText ("" & Mid(FrmRptAportes.TxtAño, 1, 1) & "")
            report3.Text3.SetText ("" & Mid(FrmRptAportes.TxtAño, 2, 1) & "")
            report3.Text2.SetText ("" & Mid(FrmRptAportes.TxtAño, 3, 1) & "")
            report3.Text1.SetText ("" & Mid(FrmRptAportes.TxtAño, 4, 1) & "")
        Case 4
            report4.Text17.SetText ("" & Mid(FrmRptAportes.s_MesSeleccionRpt, 1, 1) & "")
            report4.Text16.SetText ("" & Mid(FrmRptAportes.s_MesSeleccionRpt, 2, 1) & "")
            report4.Text15.SetText ("" & Mid(FrmRptAportes.TxtAño, 1, 1) & "")
            report4.Text14.SetText ("" & Mid(FrmRptAportes.TxtAño, 2, 1) & "")
            report4.Text13.SetText ("" & Mid(FrmRptAportes.TxtAño, 3, 1) & "")
            report4.Text12.SetText ("" & Mid(FrmRptAportes.TxtAño, 4, 1) & "")
    End Select
End Sub
Sub Recupera_Ruc()
    Call Recupera_Ruc_Empresa_Reporte
    Set rs_rptasistencias = Reportes_Centrales.rs_RptCentrales_pub
    Set Reportes_Centrales.rs_RptCentrales_pub = Nothing
    s_RucEmpresa = rs_rptasistencias!RUC
    Set rs_rptasistencias = Nothing
End Sub
Sub Enviar_Ruc_Empresa_Reporte()
    report4.Text1.SetText ("" & Mid(s_RucEmpresa, 1, 1) & "")
    report4.Text2.SetText ("" & Mid(s_RucEmpresa, 2, 1) & "")
    report4.Text3.SetText ("" & Mid(s_RucEmpresa, 3, 1) & "")
    report4.Text4.SetText ("" & Mid(s_RucEmpresa, 4, 1) & "")
    report4.Text5.SetText ("" & Mid(s_RucEmpresa, 5, 1) & "")
    report4.Text6.SetText ("" & Mid(s_RucEmpresa, 6, 1) & "")
    report4.Text7.SetText ("" & Mid(s_RucEmpresa, 7, 1) & "")
    report4.Text8.SetText ("" & Mid(s_RucEmpresa, 8, 1) & "")
    report4.Text9.SetText ("" & Mid(s_RucEmpresa, 9, 1) & "")
    report4.Text10.SetText ("" & Mid(s_RucEmpresa, 10, 1) & "")
    report4.Text11.SetText ("" & Mid(s_RucEmpresa, 11, 1) & "")
End Sub
Sub Enviar_Totales_Segundo_Nivel()
    Select Case i_Direccion_Reportes
        Case 3
            report3.Text8.SetText ("" & Format(FrmRptAportes.i_TotalAporteObligatorio, "0.00") & "")
            report3.Text9.SetText ("" & Format(FrmRptAportes.i_TotalRetencionRetribucion, "0.00") & "")
            report3.Text10.SetText ("X")
        Case 4
            report4.Text19.SetText ("" & Format(FrmRptAportes.i_TotalAporteObligatorio, "0.00") & "")
            report4.Text20.SetText ("" & Format(FrmRptAportes.i_TotalRetencionRetribucion, "0.00") & "")
            report4.Text21.SetText ("X")
        Case 5
            report5.Text1.SetText ("" & Format(FrmRptAportes.i_TotalAporteObligatorio, "0.00") & "")
            report5.Text2.SetText ("" & Format(FrmRptAportes.i_TotalRetencionRetribucion, "0.00") & "")
            report5.Text3.SetText ("X")
        Case 6
            report6.Text2.SetText ("" & Format(FrmRptAportes.i_TotalAporteObligatorio, "0.00") & "")
            report6.Text3.SetText ("" & Format(FrmRptAportes.i_TotalRetencionRetribucion, "0.00") & "")
            report6.Text1.SetText ("X")
    End Select
End Sub
Sub Envia_Informacion_Comprobante_Retencion()
    report7.Text19.SetText ("" & FrmRptRetenciones.s_Jornal & "")
    report7.Text20.SetText ("" & FrmRptRetenciones.s_Incremento & "")
    report7.Text21.SetText ("" & FrmRptRetenciones.s_Gratificacion & "")
    report7.Text23.SetText ("" & FrmRptRetenciones.s_Vacaciones & "")
    report7.Text5.SetText ("" & FrmRptRetenciones.TxtAño & "")
    report7.Text24.SetText ("" & FrmRptRetenciones.i_Total_General1 & "")
    report7.Text12.SetText ("" & FrmRptRetenciones.s_NombreCompleto & "")
    report7.Text29.SetText ("" & FrmRptRetenciones.i_TotRetRetribucion_Afp & "")
    report7.Text31.SetText ("" & FrmRptRetenciones.i_TotRetRetribucion_Onp & "")
    report7.Text33.SetText ("" & FrmRptRetenciones.i_TotalRetenciones & "")
    report7.Text28.SetText ("" & FrmRptRetenciones.i_TotalRetenciones & "")
    report7.Text27.SetText ("" & FrmRptRetenciones.s_NombreAfp & "")
    report7.Text14.SetText ("" & FrmRptRetenciones.i_TotalRetenciones & "")
End Sub
Sub Envia_Informacion_Meses_Reporte_Promedios()
    Dim i_Mes_Contador As Integer
    Dim i_vueltas_Contador As Integer
    Dim s_Mes_Actual As String
    If FrmRptPromedios.CmbMes.ListIndex = 11 Then
    i_Mes_Contador = FrmRptPromedios.CmbMes.ListIndex + 1
    Else
    i_Mes_Contador = FrmRptPromedios.CmbMes.ListIndex
    End If
    
    
    For i_vueltas_Contador = 1 To 6
        Select Case i_Mes_Contador
            Case 1: s_Mes_Actual = "Enero"
            Case 2: s_Mes_Actual = "Febrero"
            Case 3: s_Mes_Actual = "Marzo"
            Case 4: s_Mes_Actual = "Abril"
            Case 5: s_Mes_Actual = "Mayo"
            Case 6: s_Mes_Actual = "Junio"
            Case 7: s_Mes_Actual = "Julio"
            Case 8: s_Mes_Actual = "Agosto"
            Case 9: s_Mes_Actual = "Setiembre"
            Case 10: s_Mes_Actual = "Octubre"
            Case 11: s_Mes_Actual = "Noviembre"
            Case 12: s_Mes_Actual = "Diciembre"
        End Select
        Select Case i_vueltas_Contador
            Case 1: report8.Text8.SetText (s_Mes_Actual)
            Case 2: report8.Text7.SetText (s_Mes_Actual)
            Case 3: report8.Text6.SetText (s_Mes_Actual)
            Case 4: report8.Text5.SetText (s_Mes_Actual)
            Case 5: report8.Text4.SetText (s_Mes_Actual)
            Case 6: report8.Text1.SetText (s_Mes_Actual)
        End Select
        i_Mes_Contador = i_Mes_Contador - 1
        If i_Mes_Contador = 0 Then
            i_Mes_Contador = 12
        End If
    Next i_vueltas_Contador
End Sub
