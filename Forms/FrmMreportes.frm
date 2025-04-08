VERSION 5.00
Begin VB.Form FrmMreportes 
   BackColor       =   &H80000012&
   Caption         =   "REPORTES CONTABLES"
   ClientHeight    =   3840
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9030
   LinkTopic       =   "Form1"
   ScaleHeight     =   3840
   ScaleWidth      =   9030
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "FRMMREPORTES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim crapp As New CRAXDRT.Application
Dim crreport As New CRAXDRT.Report
Dim crapp2 As New CRAXDRT.Application
Dim crreport2 As New CRAXDRT.Report
Dim rs_reportes As ADODB.Recordset
Private Sub Form_Activate()
    Call Procesa_Reporte_Detallado(FrmMGenera.s_Año_ProcesoReport, FrmMGenera.s_Mes_ProcesoReport, _
    FrmMGenera.s_Tip_TrabajadorReport, FrmMGenera.s_Tip_BoletaReport, FrmMGenera.i_TipoReporte)
End Sub
Private Sub Form_Resize()
'    CRViewer1.Top = 0: CRViewer1.Left = 0
 '   CRViewer1.Height = ScaleHeight: CRViewer1.Width = ScaleWidth
End Sub
Sub Procesa_Reporte_Detallado(Año_Proceso As String, Mes_Proceso As String, Tipo_Trabajador As String, _
Tipo_Boleta As String, Opcion As Integer)
    Dim s_NombreTip_Trabajador As String
    Select Case Tipo_Trabajador
        Case 0: s_NombreTip_Trabajador = "OBREROS"
        Case 1: s_NombreTip_Trabajador = "EMPLEADOS"
    End Select
    Select Case Opcion
        Case 0
            Set crreport = crapp.OpenReport(App.Path & "\asientos1.rpt")
            crreport.RecordSelectionFormula = "{Asientos_Pla.Pla_Año}='" & Año_Proceso & "' " & _
            "and {Asientos_Pla.Pla_Mes}='" & Mes_Proceso & "' and " & _
            "{Asientos_Pla.Pla_TipTrabajador}='" & Tipo_Trabajador & "' and " & _
            "{Asientos_Pla.Pla_Boleta}='" & Tipo_Boleta & "' and ({Asientos_Pla.Pla_Debe}>0 or " & _
            "{Asientos_Pla.Pla_Haber}>0)"
            crreport.FormulaFields.GetItemByName("Titulo").Text = "'ASIENTOS CONTABLES / DETALLES POR BOLETA'"
            crreport.FormulaFields.GetItemByName("Titulo2").Text = "'" & s_NombreTip_Trabajador & " DE " & FrmMGenera.CmbMes.Text & " DEL " & FrmMGenera.TxtAño & " '"
            'Stop
 '           Call Recupera_Informacion_Reporte
'            crreport.Database.SetDataSource rs_reportes
            'Screen.MousePointer = vbHourglass
            'CRViewer1.ReportSource = Report
'            crreport.SQLQueryString = "select Asientos_Pla.Pla_CodTrabajador,Asientos_Pla.Pla_Año," & _
 '           "Asientos_Pla.Pla_Mes,Asientos_Pla.Pla_Semana,Asientos_Pla.Pla_Boleta,Asientos_Pla.Pla_Voucher," & _
  '          "Asientos_Pla.Pla_Cgcod,Asientos_Pla.Pla_Haber,Asientos_Pla.Pla_Debe,Asientos_Pla.Pla_TipTrabajador " & _
   '         "From Asientos_Pla Where Asientos_Pla.Pla_Mes = '07' and Asientos_Pla.Pla_TipTrabajador='0' and " & _
    '        "Asientos_Pla.Pla_Año = '2008' AND (Asientos_Pla.Pla_Debe > 0. OR Asientos_Pla.Pla_Haber > 0.) and " & _
     '       "Asientos_Pla.Pla_Boleta= '01' Order By Asientos_Pla.Pla_Voucher "
'            crreport.Database.SetDataSource "select Asientos_Pla.Pla_CodTrabajador,Asientos_Pla.Pla_Año," & _
 '           "Asientos_Pla.Pla_Mes,Asientos_Pla.Pla_Semana,Asientos_Pla.Pla_Boleta,Asientos_Pla.Pla_Voucher," & _
  '          "Asientos_Pla.Pla_Cgcod,Asientos_Pla.Pla_Haber,Asientos_Pla.Pla_Debe,Asientos_Pla.Pla_TipTrabajador " & _
   '         "From Asientos_Pla Where Asientos_Pla.Pla_Mes = '07' and Asientos_Pla.Pla_TipTrabajador='0' and " & _
    '        "Asientos_Pla.Pla_Año = '2008' AND (Asientos_Pla.Pla_Debe > 0. OR Asientos_Pla.Pla_Haber > 0.) and " & _
     '       "Asientos_Pla.Pla_Boleta= '01' Order By Asientos_Pla.Pla_Voucher "
            CRViewer1.ReportSource = crreport
            'CRViewer
            CRViewer1.ViewReport
        Case 1
            Set crreport2 = crapp2.OpenReport(App.Path & "\asientos_Generales.rpt")
 '           crreport2.RecordSelectionFormula = "{Asientos_Pla.Pla_Año}='" & Año_Proceso & "' " & _
  '          "and {Asientos_Pla.Pla_Mes}='" & Mes_Proceso & "' and {Asientos_Pla.Pla_TipTrabajador}='" & Tipo_Trabajador & "' and " & _
   '         "({Asientos_Pla.Pla_Haber}>0 or {Asientos_Pla.Pla_Debe}>0) and {Asientos_Pla.Pla_Boleta}='" & Tipo_Boleta & "'  "
            crreport2.FormulaFields.GetItemByName("Titulo").Text = "'ASIENTOS CONTABLES / TOTAL GENERALES'"
            crreport2.FormulaFields.GetItemByName("Titulo_General2").Text = "'" & s_NombreTip_Trabajador & " DE " & FrmMGenera.CmbMes.Text & " DEL " & FrmMGenera.TxtAño & " '"
            crreport2.FormulaFields.GetItemByName("AÑO").Text = "'2007'"
            CRViewer1.ReportSource = crreport2
'            CRViewer1.Refresh
            CRViewer1.ViewReport
    End Select
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set crreport = Nothing: Set crapp = Nothing
    Set crreport2 = Nothing: Set crapp2 = Nothing
End Sub
Sub Recupera_Informacion_Reporte()
    Dim s_ReportesCont As String
    s_ReportesCont = "select Asientos_Pla.Pla_CodTrabajador,Asientos_Pla.Pla_Año," & _
    "Asientos_Pla.Pla_Mes,Asientos_Pla.Pla_Semana,Asientos_Pla.Pla_Boleta,Asientos_Pla.Pla_Voucher," & _
    "Asientos_Pla.Pla_Cgcod,Asientos_Pla.Pla_Haber,Asientos_Pla.Pla_Debe,Asientos_Pla.Pla_TipTrabajador " & _
    "From Asientos_Pla Where Asientos_Pla.Pla_Mes = '07' and Asientos_Pla.Pla_TipTrabajador='0' and " & _
    "Asientos_Pla.Pla_Año = '2008' AND (Asientos_Pla.Pla_Debe > 0. OR Asientos_Pla.Pla_Haber > 0.) and " & _
    "Asientos_Pla.Pla_Boleta= '01' Order By Asientos_Pla.Pla_Voucher "
    Set rs_reportes = New ADODB.Recordset
    rs_reportes.Open s_ReportesCont, cn, adOpenKeyset, adLockOptimistic
    rs_reportes.Requery
End Sub
