VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form RptDetalles 
   Caption         =   "Sistema de Planillas"
   ClientHeight    =   5700
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10815
   LinkTopic       =   "Form1"
   ScaleHeight     =   5700
   ScaleWidth      =   10815
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin CRVIEWER9LibCtl.CRViewer9 CRViewer91 
      Height          =   7005
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10365
      lastProp        =   500
      _cx             =   18283
      _cy             =   12356
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   0   'False
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
Attribute VB_Name = "RptDetalles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Report As New CrystalReport1

Private Sub Form_Activate()
    Dim rs As New ADODB.Recordset
    CRViewer91.ReportSource = Report
   
    Report.DiscardSavedData
    If wcia = "06" Then
        Report.Database.Tables(1).ConnectionProperties("User ID") = "sa"
        Report.Database.Tables(1).ConnectionProperties("Password") = "SistemasRoda@"
    Else
        Report.Database.Tables(1).ConnectionProperties("Data Source") = wserver
        Report.Database.Tables(1).ConnectionProperties("Initial Catalog") = WDatabase
        
        Report.Database.Tables(1).ConnectionProperties("User ID") = "USURPT"
        Report.Database.Tables(1).ConnectionProperties("Password") = "USURPT"
    End If
    Report.Database.Tables(1).SetDataSource rs
        
    Report.SQLQueryString = "select Pla_Voucher,Pla_CodTrabajador,Pla_CC,Pla_Cgcod,Pla_Tipo,Pla_Debe," & _
    "Pla_Haber from Asientos_Pla where Pla_Año=" & FrmMGenera.s_Año_ProcesoReport & " and Pla_Mes=" & _
    "" & FrmMGenera.s_Mes_ProcesoReport & " and Pla_TipTrabajador='" & FrmMGenera.s_Tip_TrabajadorReport & "' and " & _
    "Pla_Boleta='" & FrmMGenera.s_Tip_BoletaReport & "' and pla_cia='" & wcia & "' " & _
    "and (Pla_Debe<>0 or Pla_Haber<>0) order by pla_codtrabajador,pla_voucher,pla_cgcod"
    
    Screen.MousePointer = vbHourglass
    
    CRViewer91.ViewReport
    CRViewer91.Refresh
    
    Report.Text16.SetText ("REPORTE DE ASIENTOS CONTABLES DETALLADO " & FrmMGenera.CmbTrabTipo.Text & "S ")
    Report.Text17.SetText ("AL MES DE " & FrmMGenera.CmbMes.Text & " ")
           
    Screen.MousePointer = vbDefault
    
End Sub
Private Sub Form_Resize()
    CRViewer91.Top = 0
    CRViewer91.Left = 0
    CRViewer91.Height = ScaleHeight
    CRViewer91.Width = ScaleWidth
End Sub
