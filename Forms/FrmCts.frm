VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form FrmCts 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Depositos de CTS"
   ClientHeight    =   8910
   ClientLeft      =   4395
   ClientTop       =   2400
   ClientWidth     =   8025
   Icon            =   "FrmCts.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8910
   ScaleWidth      =   8025
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FramTipoCamb 
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   3960
      TabIndex        =   24
      Top             =   6360
      Visible         =   0   'False
      Width           =   4095
      Begin VB.TextBox TxtTc 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2280
         TabIndex        =   26
         Top             =   840
         Width           =   855
      End
      Begin MSComCtl2.DTPicker Cbofecha 
         Height          =   285
         Left            =   2280
         TabIndex        =   27
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Format          =   200802305
         CurrentDate     =   37265
      End
      Begin MSForms.CommandButton CommandButton1 
         Height          =   375
         Left            =   120
         TabIndex        =   30
         Top             =   1320
         Width           =   1305
         ForeColor       =   16711680
         Caption         =   "  Salir"
         PicturePosition =   327683
         Size            =   "2302;661"
         Picture         =   "FrmCts.frx":030A
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton CommandButton2 
         Height          =   375
         Left            =   2160
         TabIndex        =   29
         Top             =   1320
         Width           =   1665
         ForeColor       =   16711680
         Caption         =   "Aceptar"
         PicturePosition =   327683
         Size            =   "2937;661"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000C&
         Caption         =   "Tipo de Cambio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   28
         Top             =   880
         Width           =   1335
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000C&
         Caption         =   "Fecha de Deposito"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   25
         Top             =   360
         Width           =   1620
      End
   End
   Begin VB.Frame FrameDepo 
      BackColor       =   &H8000000B&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   3840
      TabIndex        =   21
      Top             =   8200
      Visible         =   0   'False
      Width           =   4095
      Begin VB.CommandButton Command5 
         Caption         =   "Tipo de Cambio Fecha Deposito"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2400
         TabIndex        =   23
         Top             =   80
         Width           =   1515
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Reporte por Bancos Actualizar Planilla"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   22
         Top             =   80
         Width           =   1875
      End
   End
   Begin VB.CommandButton CmdProvision 
      Caption         =   "Generar Asientos Provisión"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Top             =   8280
      Width           =   3615
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Calcular"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   14
      Top             =   7800
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Certificados"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   15
      Top             =   7800
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Reporte CTS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   13
      Top             =   7800
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   0
      TabIndex        =   5
      Top             =   570
      Width           =   8055
      Begin VB.ComboBox Cmbtipo 
         Height          =   315
         ItemData        =   "FrmCts.frx":0624
         Left            =   5400
         List            =   "FrmCts.frx":0631
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   60
         Width           =   2175
      End
      Begin VB.ComboBox Cmbmes 
         Height          =   315
         ItemData        =   "FrmCts.frx":064E
         Left            =   840
         List            =   "FrmCts.frx":0676
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   60
         Width           =   2055
      End
      Begin VB.TextBox Txtano 
         Height          =   315
         Left            =   2910
         TabIndex        =   7
         Top             =   60
         Width           =   615
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Trabajador"
         Height          =   195
         Left            =   4080
         TabIndex        =   11
         Top             =   60
         Width           =   1125
      End
      Begin MSForms.SpinButton SpinButton1 
         Height          =   315
         Left            =   3510
         TabIndex        =   9
         Top             =   60
         Width           =   255
         Size            =   "450;556"
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Periodo"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   60
         Width           =   540
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   6855
      Left            =   0
      TabIndex        =   3
      Top             =   825
      Width           =   7935
      Begin Threed.SSPanel panel 
         Height          =   735
         Left            =   630
         TabIndex        =   16
         Top             =   3240
         Visible         =   0   'False
         Width           =   6735
         _Version        =   65536
         _ExtentX        =   11880
         _ExtentY        =   1296
         _StockProps     =   15
         ForeColor       =   0
         BackColor       =   13160660
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FloodColor      =   0
         Alignment       =   6
         Begin MSComctlLib.ProgressBar pBar 
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   240
            Width           =   6495
            _ExtentX        =   11456
            _ExtentY        =   450
            _Version        =   393216
            Appearance      =   0
         End
      End
      Begin MSDataGridLib.DataGrid Dgrdcabeza 
         Bindings        =   "FrmCts.frx":06DE
         Height          =   6255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   11033
         _Version        =   393216
         Appearance      =   0
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   6
         BeginProperty Column00 
            DataField       =   "nombre"
            Caption         =   "Nombre"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "monto_total"
            Caption         =   "Rem. Afecta"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "monto_total"
            Caption         =   "Deposito"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "placod"
            Caption         =   "Codigo"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "fingreso"
            Caption         =   "fingreso"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "excluye"
            Caption         =   "Excluir"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   4649.953
            EndProperty
            BeginProperty Column01 
               Alignment       =   1
               ColumnWidth     =   1065.26
            EndProperty
            BeginProperty Column02 
               Alignment       =   1
               Object.Visible         =   -1  'True
               ColumnWidth     =   824.882
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
               Object.Visible         =   0   'False
               ColumnWidth     =   900.284
            EndProperty
            BeginProperty Column04 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column05 
               Alignment       =   2
               ColumnWidth     =   659.906
            EndProperty
         EndProperty
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         Caption         =   "Doble Click para Excluir o Incluir a un trabajador del deposito de CTS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   120
         TabIndex        =   20
         Top             =   6555
         Width           =   7800
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H00808080&
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8055
      Begin VB.ComboBox Cmbcia 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   120
         Width           =   6495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Compañia"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   210
         Left            =   120
         TabIndex        =   1
         Top             =   165
         Width           =   825
      End
   End
   Begin MSAdodcLib.Adodc AdoCabeza 
      Height          =   1410
      Left            =   2520
      Top             =   4080
      Visible         =   0   'False
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   2487
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Deposito"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   210
      Left            =   5160
      TabIndex        =   19
      Top             =   7860
      Width           =   1185
   End
   Begin VB.Label Lbltotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   6480
      TabIndex        =   18
      Top             =   7800
      Width           =   1215
   End
End
Attribute VB_Name = "FrmCts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mperiodo As String
Dim Sql As String
Dim mpag As Integer
Dim RUTA As String
Dim mlinea As Integer
Dim pla_area As String
'***************codigo nuevo giovanni 17082007*************************
Dim s_MesSeleccion As String
Dim rs_Liquidacion As ADODB.Recordset
Dim rs_Liquidacion2 As ADODB.Recordset
Dim rsConsultaCTS As ADODB.Recordset
Dim s_CodEmpresa_Starsoft As String
Dim i_Numero_Voucher As Integer
Dim i_Numero_VoucherG As String
Dim s_TipoTrabajador_Cts As String
Dim s_Dia_Proceso_Prov As String
'**********************************************************************

Const COL_PRIMERMES = 5
Const COL_SEGUNDOMES = 11
Dim ArrReporte() As Variant

Const MAXCOL = 19
Const CLM_CODIGO = 1
Const CLM_TRABAJADOR = 2
Const CLM_FECINGRESO = 3
Const CLM_FECCESE = 4
Const CLM_TIEMPOSERV = 5
Const CLM_JORNAL = 6
Const CLM_AFP3 = 7
Const CLM_BONOCOSTO = 8
Const CLM_BONOTSERV = 9
Const CLM_ASIGFAM = 10
Const CLM_PROMGRATI = 11
Const CLM_PROMHEXTRAS = 12
Const CLM_PROMOTROSPAG = 13
Const CLM_PROMHVERANO = 14
Const CLM_PROMTURNO = 15
Const CLM_PROMPRODUCC = 16
Const CLM_PROMBONOPROD = 17
Const CLM_PROMREMDI = 18
Const CLM_JORNALINDEN = 19
Const CLM_TIEMPOSERVACT = 3
Const CLM_MONTOINDEN1 = 4
Const CLM_MONTOINDEN2 = 5
Const CLM_MONTOINDENSINTP = 6
Const CLM_MONTOTOTINDEN = 7
Const CLM_PROVANOANT = 8
Const CLM_AJUSTEPROVI = 9
Const CLM_PROVIANOACTUAL = 10
Const CLM_PROVMESACTUAL = 11
Const CLM_SALDOPEND = 12

'Cambio hecho para la impresion de los certificados de CTS
Public Id_Trab             As String
Public Fecha_dDeposito     As Date
Public T_dCambio           As Double
Public Interes_Moratorio   As Double

'Dim xlApp2 As Excel.Application
'Dim xlApp1  As Excel.Application
'Dim xlBook As Excel.Workbook
'Dim xlSheet As Excel.Worksheet

Dim xlApp2 As Object
Dim xlApp1  As Object
Dim xlBook As Object
Dim xlSheet As Object



Dim xlApp  As Object
Private Sub CmbMes_Click()
Carga_Cts
If Cmbmes.ListIndex + 1 = 4 Or Cmbmes.ListIndex + 1 = 10 Then
   FrameDepo.Visible = True
Else
   FrameDepo.Visible = False
End If
End Sub
Private Sub CmbTipo_Click()
    Call Carga_Cts
End Sub
Private Sub CmdProvision_Click()
If Txtano.Text < 2013 Then
   MsgBox "En este sistema se pueden trabajar provisiones desde Setiembre del 2013", vbInformation
   Exit Sub
End If
If Txtano.Text = 2013 And Cmbmes.ListIndex + 1 < 9 Then
   MsgBox "En este sistema se pueden trabajar provisiones desde Setiembre del 2013", vbInformation
   Exit Sub
End If

Dim strPeriodo As String, strTipotrabajador As String
Dim strTitulo As String, strLote As String, strVoucher As String, strSubDiario As String
strLote = "":   strVoucher = "": strTitulo = "": strSubDiario = ""

strPeriodo = Txtano.Text + Format(Cmbmes.ListIndex + 1, "00")
strTipotrabajador = Format(Cmbtipo.ListIndex + 1, "00")

Sql = "spPlaSeteoAsientoLotes " & "'04','" & strTipotrabajador & "'"
If (fAbrRst(rs, Sql)) Then
    strLote = rs(0).Value
    strVoucher = rs(1).Value
    strTitulo = rs(2).Value
    strSubDiario = rs(3).Value
End If
rs.Close: Set rs = Nothing
Call GenerarAsientoProvision(strPeriodo, "04", strTipotrabajador, strLote, strVoucher, strTitulo, Cmbmes.Text, strSubDiario)
    
End Sub
Private Sub Command1_Click()
'Reporte_Cts
'If wGrupoPla = "01" Then
    Dim lExcluir As String
    lExcluir = ""
    'If Cmbmes.ListIndex + 1 = 4 Or Cmbmes.ListIndex + 1 = 10 Then
       If MsgBox("Desea Reporte Sin Exluidos  ? " & Chr(13) & "Sin Excluidos para deposito de CTS", vbQuestion + vbYesNo + vbDefaultButton1, "Reporte de Cts") = vbYes Then lExcluir = "S"
    'End If
    ReporteCts (lExcluir)
'Else
'    ReporteCtsRODA
'End If
End Sub

Private Sub Command2_Click()
'If wGrupoPla <> "01" Then
    Reporte_Cts_Bancos
    Call Procesa_Archivo_Banco_Excel(Txtano.Text, Cmbmes.ListIndex + 1)
'End If
End Sub

Private Sub Command3_Click()
    Procesa_Certifica_Cts
End Sub

Private Sub Command4_Click()
If Txtano.Text < 2013 Then
   MsgBox "En este sistema se pueden trabajar provisiones desde Setiembre del 2013", vbInformation
   Exit Sub
End If
If Txtano.Text = 2013 And Cmbmes.ListIndex + 1 < 9 Then
   MsgBox "En este sistema se pueden trabajar provisiones desde Setiembre del 2013", vbInformation
   Exit Sub
End If

'**********codigo agregado giovanni 10092007**************************
If Cmbtipo.Text = "" Then
    MsgBox "Debe Seleccionar Tipo Trabajador", vbInformation
Else
    PROVICIONES_CTS
End If
'*********************************************************************
    'PROVICIONES_CTS
End Sub

Private Sub Command5_Click()
Sql = "SELECT * FROM Pla_Tc_Cts Where cia='" & wcia & "' and ano=" & Txtano.Text & " and mes=" & Cmbmes.ListIndex + 1 & " and status<>'*'"
If (fAbrRst(rs, Sql)) Then
   TxtTc.Text = rs!Tipo_Camb
   Cbofecha.Value = rs!FecDepo
Else
   TxtTc.Text = ""
   Cbofecha.Value = Date
End If
rs.Close: Set rs = Nothing


FramTipoCamb.Visible = True
End Sub

Private Sub CommandButton1_Click()
FramTipoCamb.Visible = False
End Sub

Private Sub CommandButton2_Click()
If Not IsNumeric(TxtTc.Text) Then
   MsgBox "Ingrese Correctamente Tipo de Cambio", vbInformation
   Exit Sub
End If
If CCur(TxtTc.Text) > 4 Then
   MsgBox "Ingrese Correctamente Tipo de Cambio", vbInformation
   Exit Sub
End If
If CCur(TxtTc.Text) < 2 Then
   MsgBox "Ingrese Correctamente Tipo de Cambio", vbInformation
   Exit Sub
End If


Sql$ = "Update Pla_Tc_Cts Set Status='*' Where cia='" & wcia & "' and ano=" & Txtano.Text & " and mes=" & Cmbmes.ListIndex + 1 & " and status<>'*'"
cn.Execute Sql$

Sql$ = "Insert Into Pla_Tc_Cts Values('" & wcia & "'," & Txtano.Text & "," & Cmbmes.ListIndex + 1 & "," & CCur(TxtTc.Text) & ",'',Getdate(),'" & wuser & "','" & Format(Cbofecha.Value, "mm/dd/yyyy") & "')"
cn.Execute Sql$

Sql$ = "update plaprovcts set T_dCambio=" & CCur(TxtTc.Text) & ",Fecha_dDeposito='" & Format(Cbofecha.Value, "mm/dd/yyyy") & "' where cia='" & wcia & "' and YEAR(fechaproceso)=" & Txtano.Text & " and MONTH(fechaproceso)=" & Cmbmes.ListIndex + 1 & " and status<>'*'"
cn.Execute Sql$

FramTipoCamb.Visible = False
End Sub

Private Sub Dgrdcabeza_DblClick()
On Error GoTo Pasa

'If CmbMes.ListIndex + 1 <> 4 And CmbMes.ListIndex + 1 <> 10 Then Exit Sub

Dim lCod As String
Dim lNom As String
Dim lExcluye As String
lCod = Trim(Dgrdcabeza.Columns(3) & "")
lNom = Trim(Dgrdcabeza.Columns(0) & "")
lExcluye = Trim(Dgrdcabeza.Columns(5) & "")

If lExcluye <> "S" Then
   If MsgBox("Desea Excluir al Trabajador  ? " & Chr(13) & lCod & " - " & lNom, vbQuestion + vbYesNo + vbDefaultButton1, "Excluir Trabjadores") = vbYes Then Call Excluye_Trabajador(lCod, lExcluye)
Else
   If MsgBox("Desea Incluir al Trabajador  ? " & Chr(13) & lCod & " - " & lNom, vbQuestion + vbYesNo + vbDefaultButton1, "Excluir Trabjadores") = vbYes Then Call Excluye_Trabajador(lCod, lExcluye)
End If

Pasa:
End Sub
Private Sub Excluye_Trabajador(lCod As String, lExcluye As String)
If lExcluye = "S" Then lExcluye = "" Else lExcluye = "S"
Sql$ = "Usp_Pla_Cts_Excluye '" & wcia & "'," & Txtano.Text & "," & Cmbmes.ListIndex + 1 & ",'" & lCod & "','" & lExcluye & "','" & wuser & "'"
cn.Execute Sql$
Call Carga_Cts
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
'Me.Height = 8065
'Me.Width = 7575
Call rCarCbo(Cmbcia, Carga_Cia, "C", "00")
Call rUbiIndCmbBox(Cmbcia, wcia, "00")
If Month(Date) = 1 Then
   Txtano.Text = Format(Year(Date) - 1, "0000")
   Cmbmes.ListIndex = 11
Else
  Txtano.Text = Format(Year(Date), "0000")
  Cmbmes.ListIndex = Month(Date) - 2
End If
   
End Sub

Private Sub Calculo_CTS()
Dim mcad As String
Dim mcadIns As String
Dim mCadVal As String
Dim mCadFields As String
Dim mFactor As Currency
Dim mTotal As Currency
Dim mDepo As Currency
Dim rs2 As ADODB.Recordset
Dim rsAfectos As ADODB.Recordset
Dim I As Integer, x As Integer
Dim INSCAD As String, VALUESCAD As String
Dim NRO As Integer, cad As String
Dim VALORES(50) As Integer

Screen.MousePointer = vbArrowHourglass

mperiodo = Txtano.Text & Format(Cmbmes.ListIndex + 1, "00")
Sql = "select factorcts from cia where cod_cia='" & wcia & "' and status<>'*'"
If (fAbrRst(rs, Sql)) Then
   If IsNull(rs!factorcts) Then
      MsgBox "No se ha registrado el Factor para calculo de CTS", vbInformation, "Calculo de CTS"
      Screen.MousePointer = vbDefault
      Exit Sub
   End If
   If rs!factorcts = 0 Then
      MsgBox "No se ha registrado el Factor para calculo de CTS", vbInformation, "Calculo de CTS"
      Screen.MousePointer = vbDefault
      Exit Sub
   End If
Else
   MsgBox "No se ha registrado el Factor para calculo de CTS", vbInformation, "Calculo de CTS"
   Screen.MousePointer = vbDefault
   Exit Sub
End If

mFactor = rs!factorcts
rs.Close

Sql = "select cod_remu from plaafectos where cia='" & wcia & "' and tipo='C' and status<>'*' order by cod_remu"
If (fAbrRst(rsAfectos, Sql)) Then
   rsAfectos.MoveFirst
Else
  MsgBox "No Se Registran Remuneraciones Afectas para el Calculo", vbCritical, "Calculo de CTS"
  Screen.MousePointer = vbDefault
  Exit Sub
End If

mcad = ""
mCadFields = ""

NRO = 0
Do While Not rsAfectos.EOF
   NRO = NRO + 1
   mcad = mcad & "SUM(I" & Format(rsAfectos!cod_remu, "00") & "),"
   mCadFields = mCadFields & "I" & Format(rsAfectos!cod_remu, "00") & ","
   rsAfectos.MoveNext
Loop

mcad = Mid(mcad, 1, Len(Trim(mcad)) - 1)
Sql = "select " & mcad
Sql = Sql & ",placod from plahistorico where cia='" & _
wcia & "' and year(fechaproceso)=" & Trim(Txtano.Text) & _
" and month(fechaproceso)=" & Cmbmes.ListIndex + 1 & _
" and status<>'*' Group by placod"

  If (fAbrRst(rs, Sql)) Then
     rs.MoveFirst
  Else
     MsgBox "No Se Registran Boletas para el Calculo", vbCritical, "Calculo de CTS"
     Screen.MousePointer = vbDefault
     Exit Sub
  End If

Sql = wInicioTrans
cn.Execute Sql

Sql = "update platserdep set status='*' where cia='" & wcia & "' and periodo='" & mperiodo & "' and status<>'*'"
cn.Execute Sql

Do While Not rs.EOF
   Sql = "select fingreso,cargo,ctsbanco,ctstipcta,ctsmoneda,ctsnumcta from planillas where cia='" & wcia & "' and placod='" & Trim(rs!PlaCod) & "' and status<>'*'"
   If (fAbrRst(rs2, Sql)) Then rs2.MoveFirst
   rsAfectos.MoveFirst
   mcadIns = ""
   I = 0
   mTotal = 0: mDepo = 0
   Do While Not rsAfectos.EOF
      mCadVal = "0"
      mCadVal = rs(I)
      
      If Not IsNull(rs(I)) Then mTotal = mTotal + rs(I)
      VALORES(Val(rsAfectos(0))) = rs(I)
      mcadIns = mcadIns & mCadVal & ","
      rsAfectos.MoveNext
      I = I + 1
   Loop
   mDepo = Round((mTotal * mFactor) / 100, 2)
  
   INSCAD = "i05,i07,i08," & _
   "i16,i17,i18,i19,i20," & _
   "i21,i22,i23,i24,i25,i26,i27,i28,i29,i30,i31," & _
   "i32,i33,i34,i35,i36,i37,i38,i39,i40,i41,i42," & _
   "i43,i44,i45,i46,i47,i48,i49,i50"
   
   INSCAD = ""
   VALUESCAD = ""
   For x = 1 To 50
       INSCAD = INSCAD & "I" & Format(x, "00") & ","
       VALUESCAD = VALUESCAD & VALORES(x) & ","
   Next
   
   'VALUESCAD = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
  
   mcadIns = mcadIns & Str(mTotal) & ","
   Sql = "SET DATEFORMAT " & Coneccion.FormatFechaSql & "  "
   Sql = Sql & "insert into platserdep(cia,placod,fechaingreso,cargo," & INSCAD
   Sql = Sql & "total,factor,fecha,periodo,banco,cta,moneda,nro_cta,status,interes,TOTLIQUID) "
   Sql = Sql & "values('" & wcia & "','" & Trim(rs!PlaCod) & "','" & Format(rs2!fIngreso, FormatFecha) & "','" & Trim(rs2!Cargo) & "',"
   Sql = Sql & VALUESCAD & mTotal & ","
   Sql = Sql & "" & mFactor & "," & FechaSys & ",'" & mperiodo & "','" & rs2!ctsbanco & "','" & rs2!ctstipcta & "',"
   Sql = Sql & "'" & rs2!ctsmoneda & "','" & _
   Trim(rs2!ctsnumcta) & "','',0,0" & ")"
   

   
   cn.Execute Sql
   If rs2.State = 1 Then rs2.Close
   rs.MoveNext
Loop

Sql = wFinTrans
cn.Execute Sql

If rs.State = 1 Then rs.Close
If rsAfectos.State = 1 Then rsAfectos.Close
Screen.MousePointer = vbDefault
Carga_Cts
End Sub

Private Sub SpinButton1_SpinDown()
If Trim(Txtano.Text) = "" Then
   Txtano.Text = Format(Year(Date), "0000")
Else
   Txtano.Text = Txtano.Text - 1
End If
End Sub

Private Sub SpinButton1_SpinUp()
If Trim(Txtano.Text) = "" Then
   Txtano.Text = Format(Year(Date), "0000")
Else
   Txtano.Text = Txtano.Text + 1
End If
End Sub

Private Sub Txtano_Change()
    Carga_Cts
End Sub

Private Sub Txtano_KeyPress(KeyAscii As Integer)
Txtano.Text = Txtano.Text + fc_ValNumeros(KeyAscii)
End Sub
Private Sub Carga_Cts()

If Cmbtipo.Text = "" Then
    Exit Sub
End If

Call Captura_Tipo_Trabajador

If Trim(Txtano.Text) = "" Then Exit Sub
If Cmbmes.ListIndex < 0 Then Exit Sub
mperiodo = Trim(Txtano.Text) & Format(Cmbmes.ListIndex + 1, "00")
Sql = nombre()

Select Case s_TipoTrabajador_Cts
    Case "01"
        Sql = Sql & "a.placod,a.monto_total,a.monto_total,b.fingreso, " _
        & "(Select excluye from pla_cts_excluye where cia=a.cia and ayo=year(a.fechaproceso) and mes=month(a.fechaproceso) and placod=a.placod and status<>'*') as Excluye " _
        & "from plaprovcts a,planillas b " _
        & "where a.cia='" & wcia & "' and month(a.fechaproceso)=" & Cmbmes.ListIndex + 1 & " and year(fechaproceso)=" & Txtano.Text & " and a.status<>'*' " _
        & "and a.placod=b.placod and a.cia=b.cia and b.tipotrabajador='01' and b.status<>'*' order by nombre"
    Case "02"
        Sql = Sql & "a.placod,a.monto_total,a.monto_total,b.fingreso, " _
        & "(Select excluye from pla_cts_excluye where cia=a.cia and ayo=year(a.fechaproceso) and mes=month(a.fechaproceso) and placod=a.placod and status<>'*') as Excluye " _
        & "from plaprovcts a,planillas b " _
        & "where a.cia='" & wcia & "' and month(a.fechaproceso)=" & Cmbmes.ListIndex + 1 & " and year(fechaproceso)=" & Txtano.Text & " and a.status<>'*' " _
        & "and a.placod=b.placod and a.cia=b.cia and b.tipotrabajador='02' and b.status<>'*' order by nombre"
    Case "03"
        Sql = Sql & "a.placod,a.monto_total,a.monto_total,b.fingreso, " _
        & "(Select excluye from pla_cts_excluye where cia=a.cia and ayo=year(a.fechaproceso) and mes=month(a.fechaproceso) and placod=a.placod and status<>'*') as Excluye " _
        & "from plaprovcts a,planillas b " _
        & "where a.cia='" & wcia & "' and month(a.fechaproceso)=" & Cmbmes.ListIndex + 1 & " and year(fechaproceso)=" & Txtano.Text & " and a.status<>'*' " _
        & "and a.placod=b.placod and a.cia=b.cia and b.status<>'*' order by nombre"
End Select

'sql = sql & "a.placod,a.monto_total,a.provision_actual,b.fingreso " _
'& "from plaprovcts a,planillas b " _
'& "where a.cia='" & wcia & "' and month(a.fechaproceso)=" & Cmbmes.ListIndex + 1 & " and year(fechaproceso)=" & Txtano.Text & " and a.status<>'*' " _
'& "and a.placod=b.placod and a.cia=b.cia and b.status<>'*' order by nombre"

cn.CursorLocation = adUseClient
Set AdoCabeza.Recordset = cn.Execute(Sql$, 64)
If AdoCabeza.Recordset.RecordCount > 0 Then
   AdoCabeza.Recordset.MoveFirst
   Command4.Enabled = False
   Select Case s_TipoTrabajador_Cts
      Case "01": Sql = "select SUM(monto_total) from plaprovcts where cia='" & wcia & "' and month(fechaproceso)=" & Cmbmes.ListIndex + 1 & " and year(fechaproceso)=" & Txtano.Text & " and tipotrab='01' and status<>'*'"
      Case "02": Sql = "select SUM(monto_total) from plaprovcts where cia='" & wcia & "' and month(fechaproceso)=" & Cmbmes.ListIndex + 1 & " and year(fechaproceso)=" & Txtano.Text & " and tipotrab='02' and status<>'*'"
      Case "03": Sql = "select SUM(monto_total) from plaprovcts where cia='" & wcia & "' and month(fechaproceso)=" & Cmbmes.ListIndex + 1 & " and year(fechaproceso)=" & Txtano.Text & " and status<>'*'"
   End Select
   If (fAbrRst(rs, Sql)) Then Lbltotal.Caption = Format(rs(0), "###,###.00")
   rs.Close
Else
   Command4.Enabled = True
   Lbltotal.Caption = "0.00"
End If
Dgrdcabeza.Refresh
End Sub
Public Sub Elimina_Cts()

If Txtano.Text < 2013 Then
   MsgBox "En este sistema se pueden trabajar provisiones desde Setiembre del 2013", vbInformation
   Exit Sub
End If
If Txtano.Text = 2013 And Cmbmes.ListIndex + 1 < 9 Then
   MsgBox "En este sistema se pueden trabajar provisiones desde Setiembre del 2013", vbInformation
   Exit Sub
End If

Dim NroTrans As Integer
On Error GoTo ErrorTrans
mperiodo = Txtano.Text & Format(Cmbmes.ListIndex + 1, "00")
NroTrans = 0
If MsgBox("Desea Eliminar Calculo de Cts ? ", vbQuestion + vbYesNo + vbDefaultButton1, "Seteo de Usuario") = vbNo Then
    Exit Sub
Else
    cn.BeginTrans
    NroTrans = 1
    Sql = "update plaprovcts set status='*' where cia='" & wcia & "' and month(fechaproceso)=" & Cmbmes.ListIndex + 1 & " and year(fechaproceso)=" & Txtano.Text & " and status<>'*'"
    cn.Execute Sql
    
    cn.CommitTrans
    Carga_Cts
End If

Screen.MousePointer = vbDefault

Exit Sub

ErrorTrans:
    If NroTrans = 1 Then
        cn.RollbackTrans
    End If
    MsgBox Err.Description, vbCritical, Me.Caption
    Screen.MousePointer = vbDefault
    
End Sub
Private Sub Reporte_Cts()
Dim mcad As String
Dim mtotd As Currency
Dim mtott As Currency
mpag = 0
If AdoCabeza.Recordset.RecordCount <= 0 Then
   MsgBox "No Hay Depositos", vbInformation, "Depositos de CTS": Exit Sub
   Exit Sub
End If
RUTA = App.Path & "\REPORTS\" & "RepCts.txt"
Open RUTA For Output As #1
Cabeza_Lista_CTS
mtotd = 0: mtott = 0
AdoCabeza.Recordset.MoveFirst
Do While Not AdoCabeza.Recordset.EOF
   If mlinea > 55 Then Print #1, Chr(12) + Chr(13): Cabeza_Lista_CTS
   mcad = lentexto(55, Left(AdoCabeza.Recordset!nombre, 55))
   mcad = AdoCabeza.Recordset!PlaCod & "   " & mcad & Space(5) & fCadNum(AdoCabeza.Recordset!Total, "##,###,##0.00") & Space(5) & fCadNum(AdoCabeza.Recordset!totliquid, "##,###,##0.00")
   Print #1, Space(2) & mcad
   mtotd = mtotd + AdoCabeza.Recordset!totliquid
   mtott = mtott + AdoCabeza.Recordset!Total
   AdoCabeza.Recordset.MoveNext
Loop
Print #1,
Print #1, Space(35) & "TOTAL :                           " & fCadNum(mtott, "###,###,##0.00") & Space(4) & fCadNum(mtotd, "###,###,##0.00")
Close #1
Call Imprime_Txt("RepCts.txt", RUTA)
End Sub
Private Sub Cabeza_Lista_CTS()
mpag = mpag + 1
Print #1, Chr(18) & Space(2) & Trim(Cmbcia.Text) & Space(25) & "Pag. " & fCadNum(mpag, "###")
Print #1,
Print #1, Space(22) & "DEPOSITO POR TIEMPO DE SERVICIO"
Print #1, Space(23) & "PERIODO : " & Cmbmes.Text & " - " & Format(Txtano.Text, "0000") & Chr(15)
Print #1, Space(2) & String(107, "-")
Print #1, Space(2) & "CODIGO            NOMBRE                                               REMUNERACION          MONTO A "
Print #1, Space(2) & "                                                                        COMPUTABLE          DEPOSITAR"
Print #1, Space(2) & String(107, "-")
mlinea = 10
End Sub
Private Sub Reporte_Cts_Bancos()
Dim mcad As String
Dim mcod As String
Dim mItem As Integer
Dim mTotDep As Currency
Dim rs2 As ADODB.Recordset
Dim wciamae As String
Dim mtotsbco As Currency
Dim mtotnbco As Integer
Dim mtotsbcoTS As Currency
Dim mtotnbcoTS As Integer
Dim mtotsbcoTD As Currency
Dim mtotnbcoTD As Integer
Dim lExcluye As String

If Trim(Txtano.Text) = "" Then Exit Sub
If Cmbmes.ListIndex < 0 Then Exit Sub
mperiodo = Txtano.Text & Format(Cmbmes.ListIndex + 1, "00")
Sql$ = nombre

Dim lTipoTrab As String
lTipoTrab = s_TipoTrabajador_Cts
If lTipoTrab = "03" Then lTipoTrab = "**"
If Cmbtipo.Text = "" Then lTipoTrab = "**"

lExcluye = " and a.placod not in(select placod from pla_cts_excluye where cia='" & wcia & "' and ayo=" & Trim(Txtano.Text) & " and mes='" & Format(Cmbmes.ListIndex + 1, "00") & "' and excluye='S' and status<>'*') "

Sql = "select ap_pat,ap_mat,rtrim(nom_1)+' '+rtrim(nom_2) AS nombre,"
Sql = Sql & "a.placod,a.totalremun AS total,a.monto_total AS totliquid,a.ctsbanco AS banco,a.ctsnumcta as nro_cta,a.ctsmoneda as moneda,b.fingreso " _
& "from plaprovcts a,planillas b " _
& "where a.cia='" & wcia & "' and MONTH(a.FECHAPROCESO)='" & Format(Cmbmes.ListIndex + 1, "00") & "' and a.status<>'*' " _
& " and year(a.FECHAPROCESO)=" & Trim(Txtano.Text) & ""
If lTipoTrab <> "**" Then
   Sql = Sql & " and a.tipotrab='" & lTipoTrab & "' "
End If
Sql = Sql & " and a.placod=b.placod and a.cia=b.cia and b.status<>'*' "
Sql = Sql & lExcluye
Sql = Sql & " order by banco,B.moneda,nombre"

If (fAbrRst(rs, Sql)) Then
   rs.MoveFirst
Else
   MsgBox "No Hay Depositos", vbInformation, "Depositos de CTS": Exit Sub
   Exit Sub
End If

RUTA = App.Path & "\REPORTS\" & "BcoCts.txt"
Open RUTA For Output As #1
Call Cabeza_Banco_CTS(rs!banco, rs!moneda)
mcod = rs!banco & rs!moneda
mItem = 0: mTotDep = 0
Do While Not rs.EOF
   If mcod <> rs!banco & rs!moneda Then
      Print #1, Space(12) & String(20, "_") & Space(5) & String(71, "_") & Space(2) & String(13, "_")
      Print #1, Space(37) & "TOTAL INSCRITOS < " & fCadNum(mItem, "#####") & " >                            TOTAL ACUM.       < " & fCadNum(mTotDep, "##,###,##0.00") & " >"
      mItem = 0: mTotDep = 0
     Print #1, Chr(12) + Chr(13): Call Cabeza_Banco_CTS(rs!banco, rs!moneda)
     mcod = rs!banco & rs!moneda
   End If
   If mlinea > 55 Then Print #1, Chr(12) + Chr(13): Call Cabeza_Banco_CTS(rs!banco, rs!moneda)
   mcad = lentexto(24, Left(rs!nro_cta, 24)) & Space(1)
   mcad = mcad & lentexto(20, Left(rs!ap_pat, 20)) & Space(2) & lentexto(20, Left(rs!ap_mat, 20)) & Space(2) & Space(2) & lentexto(25, Left(rs!nombre, 25))
   mcad = mcad & Space(2) & fCadNum(rs!totliquid, "##,###,##0.00")
   Print #1, Space(12) & mcad
   mlinea = mlinea + 1
   mItem = mItem + 1
   mTotDep = mTotDep + rs!totliquid
   rs.MoveNext
Loop
Print #1, Space(12) & String(20, "_") & Space(5) & String(71, "_") & Space(2) & String(13, "_")
Print #1, Space(37) & "TOTAL INSCRITOS < " & fCadNum(mItem, "#####") & " >                            TOTAL ACUM.       < " & fCadNum(mTotDep, "##,###,##0.00") & " >"
If rs.State = 1 Then rs.Close


Sql = "select distinct(banco) from platserdep where cia='" & wcia & "' and periodo='" & mperiodo & "' and  status<>'*'"

Sql = "select distinct(b.ctsbanco) as banco from plaprovcts a inner join planillas b ON (b.cia=a.cia and b.placod=a.placod and b.status!='*' )"
Sql = Sql & " where a.cia='" & wcia & "' and month(a.fechaproceso)=" & Cmbmes.ListIndex + 1 & " and year(a.FECHAPROCESO)=" & Trim(Txtano.Text) & " and  a.status<>'*' "
Sql = Sql & lExcluye
If lTipoTrab <> "**" Then Sql = Sql & " and a.tipotrab='" & lTipoTrab & "' "


If (fAbrRst(rs, Sql)) Then rs.MoveFirst
Print #1, Chr(12) + Chr(13): Cabeza_Banco_Resumen
mtotnbcoTS = 0: mtotsbcoTS = 0
mtotnbcoTD = 0: mtotsbcoTD = 0
Do While Not rs.EOF
   mcad = ""
   wciamae = Determina_Maestro("01007")
   mtotnbco = 0: mtotsbco = 0
   Sql$ = "Select descrip from maestros_2 where status<>'*' and cod_maestro2='" & rs!banco & "'" & Space(5)
   Sql$ = Sql$ & wciamae
   If (fAbrRst(rs2, Sql$)) Then mcad = lentexto(33, Left(rs2!DESCRIP, 23)) Else mcad = Space(33)
   If rs2.State = 1 Then rs2.Close
   
   'Dolares
   Sql = "select count(a.placod) as num from plaprovcts a inner join planillas b ON (b.cia=a.cia and b.placod=a.placod and b.status!='*' )"
   Sql = Sql & " where a.cia='" & wcia & "' and month(a.fechaproceso)=" & Cmbmes.ListIndex + 1 & " and year(a.FECHAPROCESO)=" & Trim(Txtano.Text) & " and  b.ctsbanco='" & rs!banco & "' and b.ctsmoneda<>'" & wmoncont & "' and a.status<>'*' "
   Sql = Sql & lExcluye
   If lTipoTrab <> "**" Then Sql = Sql & " and a.tipotrab='" & lTipoTrab & "' "
   
   If (fAbrRst(rs2, Sql)) Then
      If rs2!Num = 0 Then
         mcad = mcad & "     0"
      Else
         mcad = mcad & " " & fCadNum(rs2!Num, "#####")
      End If
      mtotnbco = mtotnbco + rs2!Num
      mtotnbcoTD = mtotnbcoTD + rs2!Num
      rs2.Close
    
      Sql = "select sum(a.monto_total) as depo from plaprovcts a inner join planillas b ON (b.cia=a.cia and b.placod=a.placod and b.status!='*' )"
      Sql = Sql & " where a.cia='" & wcia & "' and month(a.fechaproceso)=" & Cmbmes.ListIndex + 1 & " and year(a.FECHAPROCESO)=" & Trim(Txtano.Text) & " and  b.ctsbanco='" & rs!banco & "' and b.ctsmoneda<>'" & wmoncont & "' and a.status<>'*' "
      Sql = Sql & lExcluye
      If lTipoTrab <> "**" Then Sql = Sql & " and a.tipotrab='" & lTipoTrab & "' "
      
      
      If (fAbrRst(rs2, Sql)) Then
         If IsNull(rs2!depo) Then
            mcad = mcad & " " & fCadNum(0, "#,###,##0.00")
         Else
            mcad = mcad & " " & fCadNum(rs2!depo, "#,###,##0.00")
            mtotsbco = mtotsbco + rs2!depo
            mtotsbcoTD = mtotsbcoTD + rs2!depo
         End If
      End If
      rs2.Close
      mcad = mcad & Space(2)
   End If
   
   'Soles
    Sql = "select count(a.placod) as num from plAprovcts a inner join planillas b ON (b.cia=a.cia and b.placod=a.placod and b.status!='*' )"
    Sql = Sql & " where a.cia='" & wcia & "' and month(a.fechaproceso)=" & Cmbmes.ListIndex + 1 & " and year(a.FECHAPROCESO)=" & Trim(Txtano.Text) & " and  b.ctsbanco='" & rs!banco & "' and b.ctsmoneda='" & wmoncont & "' and a.status<>'*' "
    Sql = Sql & lExcluye
    If lTipoTrab <> "**" Then Sql = Sql & " and a.tipotrab='" & lTipoTrab & "' "


   
   If (fAbrRst(rs2, Sql)) Then
      mcad = mcad & Space(8)
      If rs2!Num = 0 Then
         mcad = mcad & "     0"
      Else
         mcad = mcad & " " & fCadNum(rs2!Num, "#####")
      End If
      mtotnbco = mtotnbco + rs2!Num
      mtotnbcoTS = mtotnbcoTS + rs2!Num
      rs2.Close
            
      Sql = "select sum(monto_total) as depo from plaprovcts a inner join planillas b ON (b.cia=a.cia and b.placod=a.placod and b.status!='*' )"
        Sql = Sql & " where a.cia='" & wcia & "' and month(a.fechaproceso)=" & Cmbmes.ListIndex + 1 & " and year(a.FECHAPROCESO)=" & Trim(Txtano.Text) & " and  b.ctsbanco='" & rs!banco & "' and b.ctsmoneda='" & wmoncont & "' and a.status<>'*' "
        Sql = Sql & lExcluye
        If lTipoTrab <> "**" Then Sql = Sql & " and a.tipotrab='" & lTipoTrab & "' "
      
      If (fAbrRst(rs2, Sql)) Then
         If IsNull(rs2!depo) Then
            mcad = mcad & " " & fCadNum(0, "#,###,##0.00")
         Else
            mcad = mcad & " " & fCadNum(rs2!depo, "#,###,##0.00")
            mtotsbco = mtotsbco + rs2!depo
            mtotsbcoTS = mtotsbcoTS + rs2!depo
         End If
      End If
      rs2.Close
      'Total
      mcad = mcad & Space(10) & " " & fCadNum(mtotnbco, "#####") & " " & fCadNum(mtotsbco, "#,###,##0.00")
      'mcad = mcad & Space(2)
      'Se cambia por correo enviado del bbva (Emily Luyo Mendoza) el 12/05/014, la fecha de deposito debe ser la fecha del dia que se envia el archivo.
      mcad = mcad & Space(1)
   End If
   Print #1, Space(12) & mcad
   rs.MoveNext
Loop
Print #1, Space(12) & String(30, "_") & Space(5) & String(19, "_") & Space(10) & String(19, "_") & Space(10) & String(19, "_")
'Total General
mcad = "  ** T O T A L ** " & Space(16) & fCadNum(mtotnbcoTD, "#####") & " " & fCadNum(mtotsbcoTD, "#,###,##0.00")
mcad = mcad & Space(11) & fCadNum(mtotnbcoTS, "#####") & " " & fCadNum(mtotsbcoTS, "#,###,##0.00")
mcad = mcad & Space(11) & fCadNum(mtotnbcoTS + mtotnbcoTD, "#####") & " " & fCadNum(mtotsbcoTS + mtotsbcoTD, "#,###,##0.00")
Print #1, Space(12) & mcad
Close #1
Call Imprime_Txt("BcoCts.txt", RUTA)

'If MsgBox("Desea Actualizar Planilla", vbInformation + vbYesNo, "CTS") = vbYes Then
'
'    Call Insertar_PlanillaCTS(wcia, Trim(Txtano.Text), Str(Cmbmes.ListIndex + 1))
'
'End If

End Sub
Private Sub Cabeza_Banco_CTS(banco, moneda)
Dim wciamae As String
Dim rs2 As ADODB.Recordset
Dim mBanc As String
Dim MMON As String
If Trim(banco) = "" Then
   mBanc = "SIN BANCO"
Else
   mBanc = ""
   wciamae = Determina_Maestro("01007")
   Sql$ = "Select descrip from maestros_2 where status<>'*' and cod_maestro2='" & banco & "'" & Space(5)
   Sql$ = Sql$ & wciamae
   If (fAbrRst(rs2, Sql$)) Then mBanc = lentexto(30, Left(rs2!DESCRIP, 30))
   If rs2.State = 1 Then rs2.Close
End If
If Trim(moneda) = "" Then
   MMON = ""
ElseIf moneda = wmoncont Then
   MMON = "Moneda Nacional (" & moneda & ")"
Else
   MMON = "Moneda Extranjera (" & moneda & ")"
End If
Print #1, Chr(18) & Space(10) & Trim(Cmbcia.Text) & " - " & Trim(Cmbtipo.Text & "")
Print #1,
Print #1, Space(27) & mBanc
Print #1,
Print #1, Space(27) & MMON & Chr(15)
Print #1,
Print #1, Space(16) & "CUENTA                              APELLIDOS Y NOMBRES                                              MONTO"
Print #1, Space(12) & String(20, "_") & Space(5) & String(71, "_") & Space(2) & String(13, "_")
Print #1,
mlinea = 9
End Sub
Private Sub Cabeza_Banco_Resumen()
Print #1, Chr(18) & Space(10) & Trim(Cmbcia.Text)
Print #1,
Print #1, Space(27) & "R E S U M E N" & Chr(15)
Print #1,
Print #1, Space(12) & String(30, "_") & Space(5) & String(19, "_") & Space(10) & String(19, "_") & Space(10) & String(19, "_")
Print #1,
Print #1, Space(46) & "DEP.MON.EXT (DOLAR)              DEP. SOLES                   T O T A L"
Print #1, Space(16) & "BANCOS                        -------------------          -------------------          -------------------"
Print #1, Space(46) & "# TRAB    MONTO S/.          # TRAB    MONTO S/.          # TRAB    MONTO S/."
Print #1, Space(12) & String(30, "_") & Space(5) & String(19, "_") & Space(10) & String(19, "_") & Space(10) & String(19, "_")
Print #1,
mlinea = 9
End Sub
Public Sub Procesa_Certifica_Cts()
Dim mItem As Integer
Dim mcad As String
Dim mnombre As String
Dim cadnombre As String
Dim mcadrem As String
Dim mtot As Currency
Dim mFactor As Currency
Dim mdir As String
Dim mmoneda As String
Dim cadefectivo As String
Dim mdpto As String
Dim rs2 As ADODB.Recordset
Dim rsRem As ADODB.Recordset
Dim wciamae As String
Dim mempleado As String
Dim mobrero As String
Dim mfec As String
Dim I As Integer
Dim mctsfalta As String
Dim mctsdiassubs As String
Dim rsTmp As ADODB.Recordset

Fecha_dDeposito = Format(Date, "dd/MM/yyyy")
T_dCambio = 0
Interes_Moratorio = 0
Cadena = "SP_DATOS_ANEXOS_CERTIFICADO_CTS '" & wcia & "', " & CInt(Txtano.Text) & ", " & Cmbmes.ListIndex + 1
Set rsTmp = OpenRecordset(Cadena, cn)
If rsTmp!Tipo_Camb <> 0 Then
    Fecha_dDeposito = rsTmp!FecDepo
    T_dCambio = rsTmp!Tipo_Camb
   If Not Salva_Certificado_CTS(Fecha_dDeposito, T_dCambio) Then MsgBox "Error interno al generar los Certificados de CTS.": Exit Sub
Else
    MsgBox "No se encuentra la información necesaria." & _
    vbCrLf & _
    "Debe de ingresar la información necesaria para la emisión de los Certificados de CTS.", vbInformation + vbOKOnly, "Sistema"
    If Not Salva_Certificado_CTS(Fecha_dDeposito, T_dCambio) Then MsgBox "No se ha establecido los datos necesarios para los Certificados de CTS.": Exit Sub
End If
rsTmp.Close
Set rsTmp = Nothing

mfec = fMaxDay(Cmbmes.ListIndex + 1, Val(Txtano)) & "/" & Format(Cmbmes.ListIndex + 1, "00") & "/" & Txtano
cadnombre = nombre()
mperiodo = Txtano.Text & Format(Cmbmes.ListIndex + 1, "00")
mdir = ""
mdpto = ""
mFactor = 6

'Sql$ = "select distinct(direcc),nro,u.dpto  from cia c,ubigeos u where c.cod_cia='" & wcia & "' and c.status<>'*' and left(u.cod_ubi,5)=left(c.cod_ubi,5)"

Sql$ = " select distinct(direcc), c.*, " & _
      " (select nombre from sunat_departamento where id_dpto=left(c.cod_ubi,2)) as dpto, " & _
      " (select nombre from sunat_provincia where id_prov=left(c.cod_ubi,4)) as prov, " & _
      "  u.nombre  as dist, 'PERU' as pais from cia c, sunat_ubigeo u " & _
      "  WHERE cod_cia='" & wcia & "' AND u.id_ubigeo=*c.cod_ubi"


Sql$ = "SELECT A.*,v.flag1 AS via,ISNULL(z.dpto,'') AS ZONA FROM CIA A LEFT OUTER JOIN ubigeos z ON (Left(z.cod_ubi,5)=left(a.cod_ubi,5))"
Sql$ = Sql$ & "    LEFT OUTER JOIN maestros_2 v ON ( v.ciamaestro='01036'  and v.cod_maestro2 =  a.cod_via) WHERE a.cod_cia='" & wcia & "'"

If (fAbrRst(rs, Sql$)) Then mdir = Left(Trim(rs!via) & " " & rs!direcc, 56) & " No " & Trim(rs!NRO) & " - " & Trim(rs!DPTO): mdpto = Trim(rs!zona)

rs.Close

'Sql = "SELECT p.ctsbanco,p.ctsmoneda,p.ctstipcta,p.ctsnumcta,ppc.* FROM PLAPROVCTS ppc INNER JOIN planillas p ON (p.cia=ppc.cia and p.placod=ppc.placod and p.status!='*') "

Sql = "SELECT ppc.ctsbanco,ppc.ctsmoneda,ppc.ctstipcta,ppc.ctsnumcta,p.placodpresentacion,ppc.* FROM PLAPROVCTS ppc INNER JOIN planillas p ON (p.cia=ppc.cia and p.placod=ppc.placod and p.status!='*') "
Sql = Sql & "where ppc.cia='" & wcia & "' and MONTH(ppc.fechaproceso)='" & Format(Cmbmes.ListIndex + 1, "00") & "' AND YEAR(ppc.fechaproceso) = " & Txtano.Text & " and ppc.status<>'*' AND p.fcese is null "
Sql = Sql & " and ppc.placod in " & Id_Trab
Sql = Sql & " order by PPC.placod"

If Not (fAbrRst(rs, Sql$)) Then Else rs.MoveFirst

RUTA = App.Path & "\REPORTS\" & "CertCts.txt"
Open RUTA For Output As #1

Do While Not rs.EOF
    If IsNull(rs!Fecha_dDeposito) Or IsNull(rs!T_dCambio) Then MsgBox "No se encuentra la información necesaria.", vbExclamation + vbOKOnly, Me.Caption: Close #1: Exit Sub
    Fecha_dDeposito = Format(rs!Fecha_dDeposito, "dd/MM/yyyy")
    T_dCambio = rs!T_dCambio
    Interes_Moratorio = 0

   wciamae = Determina_Maestro("01007")
   Sql$ = "Select descrip from maestros_2 where status<>'*' and cod_maestro2='" & rs!ctsbanco & "'" & Space(5)
   Sql$ = Sql$ & wciamae
   If (fAbrRst(rs2, Sql$)) Then
      If rs!ctsbanco = "01" Or rs!ctsbanco = "02" Or rs!ctsbanco = "08" Or rs!ctsbanco = "29" Or rs!ctsbanco = "15" Or rs!ctsbanco = "68" Then
         mcad = lentexto(48, "BANCO " & Left(rs2!DESCRIP, 42))
      Else
         mcad = lentexto(48, Left(rs2!DESCRIP, 48))
      End If
   Else
      mcad = Space(48)
   End If
   
   If rs2.State = 1 Then rs2.Close

   If rs!ctsmoneda = wmoncont Then
      mmoneda = "     CUENTA A PLAZO/MONEDA NACIONAL (SOLES)       "
          
   Else
   
    ' Fecha Modificacion : 05/11/2008
    ' Modificado por     : Ricardo Hinostroza
    ' Se Reemplazo Texto
    ' mmoneda = "     CUENTA A PLAZO/EXTRANJERA NACIONAL (DOLARES) "
   
      mmoneda = "     CUENTA A PLAZO/EXTRANJERA (DOLARES)          "
      
   End If
   
   'Sql$ = cadnombre & "placod,tipotrabajador,fingreso,fcese from planillas where cia='" & wcia & "' and placod='" & rs!PLACOD & "' and status<>'*'"
   
   Sql$ = cadnombre & "placod,placodpresentacion,tipotrabajador,fingreso,fcese from planillas where cia='" & wcia & "' and placod='" & rs!PlaCod & "' and status<>'*'"
   
   If (fAbrRst(rs2, Sql$)) Then mnombre = lentexto(41, Left(rs2!nombre, 41)) Else mnombre = Space(40)
   If rs2!TipoTrabajador = "01" Then
      mempleado = " EMPLEADO [X] "
      mobrero = " OBRERO   [ ] "
   Else
      mempleado = " EMPLEADO [ ] "
      mobrero = " OBRERO   [X] "
   End If
   
    Print #1, Chr(18) & "Liquidacion de Compensacion por Tiempo de Servicios(CTS)" & Chr(15)
    
    Print #1, "Ley de Compensacion por Tiempo de Servicio TUO del D.Leg 650 (DS-1-97-TR de 27-02-97 Y DS 4-97-TR)"
    
    Print #1, Chr(218) & String(87, Chr(196)) & Chr(194) & String(44, Chr(196)) & Chr(191)
    
    Print #1, Chr(179) & "Nombre o Razon Social del Empleador :  " & lentexto(48, Left(Cmbcia.Text, 48)) & Chr(179) & "Ciudad y Fecha  " & lentexto(16, Left(mdpto, 16)) & Format(Fecha_dDeposito, "dd/MM/yyyy") & "  " & Chr(179)
    
    Print #1, Chr(195) & String(77, Chr(196)) & Chr(194) & String(9, Chr(196)) & Chr(193) & String(16, Chr(196)) & Chr(194) & String(27, Chr(196)) & Chr(180)
    
    Print #1, Chr(179) & "Direccion del Empleador :   " & lentexto(49, Left(mdir, 49)) & Chr(179) & "  Deposito  [X]           " & Chr(179) & "     Pago Directo  [ ]     " & Chr(179)
    
    Print #1, Chr(195) & String(53, Chr(196)) & Chr(194) & String(23, Chr(196)) & Chr(193) & String(26, Chr(196)) & Chr(197) & String(27, Chr(196)) & Chr(180)
    
    Print #1, Chr(179) & Space(15) & "ENTIDAD DEPOSITARIA" & Space(19) & Chr(179) & "Tipo de Cuenta  " & Space(34) & Chr(179) & "No de Cuenta" & Space(15) & Chr(179)
    
    Print #1, Chr(179) & Space(5) & mcad & Chr(179) & mmoneda & Chr(179) & Space(1) & lentexto(26, Left(Trim(rs!ctsnumcta & ""), 26)) & Chr(179)
    
    Print #1, Chr(195) & String(53, Chr(196)) & Chr(193) & String(23, Chr(196)) & Chr(194) & String(26, Chr(196)) & Chr(193) & String(27, Chr(196)) & Chr(180)
    
    Print #1, Chr(179) & Space(28) & "DATOS DEL TRABAJADOR" & Space(29) & Chr(179) & "    CONCEPTOS REMUNERATIVOS PARA DETERMINAR LA        " & Chr(179)
    
    Print #1, Chr(179) & Space(77) & Chr(179) & "               REMUNERACION COMPUTABLE                " & Chr(179)
    
    Print #1, Chr(195) & String(17, Chr(196)) & Chr(194) & String(59, Chr(196)) & Chr(197) & String(26, Chr(196)) & Chr(194) & String(27, Chr(196)) & Chr(180)
    
    Print #1, Chr(179) & "CODIGO: " & lentexto(9, Left(rs!PlaCodpresentacion, 9)) & Chr(179) & "Apell.y Nombres : " & mnombre & Chr(179) & "         CONCEPTO         " & Chr(179) & "            MONTO          " & Chr(179)
    
    Print #1, Chr(195) & String(17, Chr(196)) & Chr(197) & String(14, Chr(196)) & Chr(194) & String(17, Chr(196)) & Chr(194) & String(26, Chr(196)) & Chr(197) & String(26, Chr(196)) & Chr(197) & String(27, Chr(196)) & Chr(180)
    
    Print #1, Chr(179) & "Fecha de Ingreso " & Chr(179) & mempleado & Chr(179) & "  Fecha de Cese  " & Chr(179) & "      Motivo de Cese      " & Chr(179) & "                          " & Chr(179) & "                           " & Chr(179)
    
    mcad = ""
    
    mItem = 0
    For I = 10 To 109
        If rs(I) <> 0 Then
           mcadrem = Chr(179)
           Sql = "select descripcion from placonstante where cia='" & wcia & "' and tipomovimiento='02' and status<>'*' and codinterno='" & Right(rs(I).Name, 2) & "'"
           If (fAbrRst(rsRem, Sql$)) Then
                Select Case Trim(rsRem!Descripcion)
                Case Is = "BASICO": mcadrem = mcadrem & lentexto(26, Left("SUELDO BASICO", 26))
                Case Is = "GRATIFICACION": mcadrem = mcadrem & lentexto(26, Left("PROM.GRATIFICACION", 26))
                Case Else
                    mcadrem = mcadrem & lentexto(26, Left(rsRem!Descripcion, 26))
                End Select
           Else
                mcadrem = mcadrem & Space(26)
           End If
           mcadrem = mcadrem & Chr(179) & Space(12) & fCadNum(rs(I), "##,###,##0.00") & Space(2) & Chr(179)
           rsRem.Close
           mItem = mItem + 1
           Select Case mItem
                  Case Is = 1
                       mcad = Chr(179) & "   " & Format(rs2!fIngreso, "dd/mm/yyyy") & "    " & Chr(179) & mobrero & Chr(179) & Space(3)
                       If IsNull(rs2!fcese) Then mcad = mcad & Space(10) Else mcad = mcad & Format(rs2!fcese, "dd/mm/yyyy")
                       mcad = mcad & Space(4) & Chr(179) & Space(26) & mcadrem
                       Print #1, mcad
                       Print #1, Chr(195) & String(3, Chr(196)) & Chr(194) & String(13, Chr(196)) & Chr(193) & String(14, Chr(196)) & Chr(193) & String(17, Chr(196)) & Chr(193) & String(26, Chr(196)) & Chr(197) & String(26, Chr(196)) & Chr(197) & String(27, Chr(196)) & Chr(180)
                       Print #1, Chr(179) & "   " & Chr(179) & Space(19) & "PERIODO DE SERVICIOS QUE SE CANCELA" & Space(19) & Chr(179) & Space(26) & Chr(179) & Space(27) & Chr(179)
                  Case Is = 2
                       If Year(rs2!fIngreso) = Val(Txtano.Text) And (Month(rs2!fIngreso) = Cmbmes.ListIndex + 1 Or (DateAdd("m", -6, DateAdd("d", 1, mfec))) < ((rs2!fIngreso))) Then
                          mcad = "Del : " & Format(rs2!fIngreso, "dd/mm/yyyy") & "  Al : " & mfec & Space(3) & "No De Meses : " & Mid(rs!recordact, 4, 2) & " No de Dias : " & Mid(rs!recordact, 6, 3) & Space(3) & mcadrem
                          'cadefectivo = Chr(179) & "   POR" & Space(3) & Mid(rs!recordact, 4, 2) & " MES " & Mid(rs!recordact, 6, 3) & " Dias" & Space(8) & Chr(179)
                            If Trim(rs!recordiasfalto) = "" Then
                                cadefectivo = Chr(179) & "   POR" & Space(3) & Mid(rs!recordact, 4, 2) & " MES " & Mid(rs!recordact, 6, 3) & " DIAS" & Space(8) & Chr(179)
                            Else
                                cadefectivo = Chr(179) & "   POR" & Space(3) & Mid(rs!recordiasfalto, 4, 2) & " MES " & Mid(rs!recordiasfalto, 6, 3) & " DIAS" & Space(8) & Chr(179)
                            End If
                       Else
                           mcad = "Del : " & Format(DateAdd("m", -6, DateAdd("d", 1, mfec)), "dd/mm/yyyy") & "  Al : " & mfec & Space(3) & "No De Meses : " & Mid(rs!recordact, 4, 2) & "  No de Dias : " & Mid(rs!recordact, 6, 3) & Space(3) & mcadrem
                            If Trim(rs!recordiasfalto) = "" Then
                                cadefectivo = Chr(179) & "   POR" & Space(3) & Mid(rs!recordact, 4, 2) & " MES " & Mid(rs!recordact, 6, 3) & " DIAS" & Space(8) & Chr(179)
                            Else
                                cadefectivo = Chr(179) & "   POR" & Space(3) & Mid(rs!recordiasfalto, 4, 2) & " MES " & Mid(rs!recordiasfalto, 6, 3) & " DIAS" & Space(8) & Chr(179)
                            End If
                          'cadefectivo = Chr(179) & "   POR" & Space(3) & Mid(rs!recordact, 4, 2) & " MES " & Mid(rs!recordact, 6, 3) & " Dias" & Space(8) & Chr(179)
                       End If
                       mcad = Chr(179) & " 1 " & Chr(179) & mcad
                       Print #1, mcad
                       Print #1, Chr(195) & String(3, Chr(196)) & Chr(197) & String(73, Chr(196)) & Chr(197) & String(26, Chr(196)) & Chr(197) & String(27, Chr(196)) & Chr(180)
                  Case Is = 3
                       mcad = Chr(179) & " 2 " & Chr(179) & "Periodo no Comp.(Faltas y Subs.)  " & Chr(179) & " " + Trim(RTrim(rs!diasctsfalto + rs!dias_subs)) & " Dias" & Space(32 - Len(Trim(RTrim(rs!diasctsfalto + rs!dias_subs)))) & mcadrem
                       Print #1, mcad
                       Print #1, Chr(195) & String(3, Chr(196)) & Chr(193) & String(73, Chr(196)) & Chr(197) & String(26, Chr(196)) & Chr(197) & String(27, Chr(196)) & Chr(180)
                  Case Is = 4
                       mcad = Chr(179) & "TIEMPO EFECTIVO A LIQUIDAR (1-2)           " & cadefectivo & Space(0) & mcadrem
                       Print #1, mcad
                       Print #1, Chr(195) & String(77, Chr(196)) & Chr(197) & String(26, Chr(196)) & Chr(197) & String(27, Chr(196)) & Chr(180)
                  Case Is = 5
                       mcad = Chr(179) & "REM.COMPUT.MENSUAL PARA LIQUID.EL PER." & Chr(179) & Space(18) & fCadNum(rs!totalremun, "##,###,##0.00") & Space(20 - Len(fCadNum(rs!totalremun, "##,###,##0.00"))) & mcadrem
                       Print #1, mcad
                       Print #1, Chr(195) & String(77, Chr(196)) & Chr(197) & String(26, Chr(196)) & Chr(197) & String(27, Chr(196)) & Chr(180)
                  Case Is = 6
                       mcad = Chr(179) & "Un dozavo de la Rem.compt.mensual     " & Chr(179) & Space(18) & fCadNum((rs!totalremun / 12), "##,###,##0.00") & Space(20 - Len(fCadNum((rs!totalremun / 12), "##,###,##0.00"))) & mcadrem
                       Print #1, mcad
                       Print #1, Chr(195) & String(77, Chr(196)) & Chr(197) & String(26, Chr(196)) & Chr(197) & String(27, Chr(196)) & Chr(180)
                  Case Is = 7
                       mcad = Chr(179) & "Un 30avo del 12avo de la Rem. Compt.  " & Chr(179) & Space(18) & fCadNum(((rs!totalremun / 12) / 30), "##,###,##0.00") & Space(20 - Len(fCadNum(((rs!totalremun / 12) / 30), "##,###,##0.00"))) & mcadrem
                       Print #1, mcad
                       Print #1, Chr(195) & String(77, Chr(196)) & Chr(197) & String(26, Chr(196)) & Chr(197) & String(27, Chr(196)) & Chr(180)
                  Case Is = 8
                       mcad = Chr(179) & Space(77) & mcadrem
                       Print #1, mcad
                       Print #1, Chr(195) & String(77, Chr(196)) & Chr(197) & String(26, Chr(196)) & Chr(197) & String(27, Chr(196)) & Chr(180)
                  Case Is = 9
                       mcad = Chr(179) & Space(77) & mcadrem
                       Print #1, mcad
                       Print #1, Chr(195) & String(77, Chr(196)) & Chr(197) & String(26, Chr(196)) & Chr(197) & String(27, Chr(196)) & Chr(180)
                  Case Is = 10
                       mcad = Chr(179) & Space(77) & mcadrem
                       Print #1, mcad
                       Print #1, Chr(195) & String(77, Chr(196)) & Chr(197) & String(26, Chr(196)) & Chr(197) & String(27, Chr(196)) & Chr(180)
           End Select
        End If
    Next
    For I = mItem To 10
        mItem = mItem + 1
        mcadrem = Chr(179) & Space(26) & Chr(179) & Space(27) & Chr(179)
        Select Case mItem
               Case Is = 1
                    mcad = Chr(179) & "   " & Format(rs2!fIngreso, "dd/mm/yyyy") & "    " & Chr(179) & mobrero & Chr(179) & Space(3)
                    If IsNull(rs2!fcese) Then mcad = mcad & Space(10) Else mcad = mcad & Format(rs2!fcese, "dd/mm/yyyy")
                    mcad = mcad & Space(4) & Chr(179) & Space(26) & mcadrem
                    Print #1, mcad
                    Print #1, Chr(195) & String(3, Chr(196)) & Chr(194) & String(13, Chr(196)) & Chr(193) & String(14, Chr(196)) & Chr(193) & String(17, Chr(196)) & Chr(193) & String(26, Chr(196)) & Chr(197) & String(26, Chr(196)) & Chr(197) & String(27, Chr(196)) & Chr(180)
                    Print #1, Chr(179) & "   " & Chr(179) & Space(19) & "PERIODO DE SERVICIOS QUE SE CANCELA" & Space(19) & Chr(179) & Space(26) & Chr(179) & Space(27) & Chr(179)
               Case Is = 2
'                    If Month(rs2!fingreso) = Cmbmes.ListIndex + 1 And Year(rs2!fingreso) = Val(Txtano.Text) Then
'                       mcad = "Del : " & Format(rs2!fingreso, "dd/mm/yyyy") & "  Al : " & mfec & Space(3) & "No De Meses : " & Mid(rs!recordact, 4, 2) & " No de Dias : " & Mid(rs!recordact, 6, 3) & Space(3) & mcadrem
'                          cadefectivo = Chr(179) & "   POR" & Space(3) & Mid(rs!recordact, 4, 2) & " MES " & Mid(rs!recordact, 6, 3) & " Dias" & Space(8) & Chr(179)
'                    Else
'                       mcad = "Del : " & Format(DateAdd("m", -6, DateAdd("d", 1, mfec)), "dd/mm/yyyy") & "  Al : " & mfec & Space(3) & "No De Meses : " & Mid(rs!recordact, 4, 2) & "  No de Dias : " & Mid(rs!recordact, 6, 3) & Space(3) & mcadrem
'                       cadefectivo = Chr(179) & "   POR" & Space(3) & Mid(rs!recordact, 4, 2) & " MES " & Mid(rs!recordact, 6, 3) & " Dias" & Space(8) & Chr(179)
'                    End If
                     If Year(rs2!fIngreso) = Val(Txtano.Text) And (Month(rs2!fIngreso) = Cmbmes.ListIndex + 1 Or (DateAdd("m", -6, DateAdd("d", 1, mfec))) < ((rs2!fIngreso))) Then
                  '  If Month(rs2!fIngreso) = Cmbmes.ListIndex + 1 And Year(rs2!fIngreso) = Val(Txtano.Text) Then
                       mcad = "Del : " & Format(rs2!fIngreso, "dd/mm/yyyy") & "  Al : " & mfec & Space(3) & "No De Meses : " & Mid(rs!recordact, 4, 2) & " No de Dias : " & Mid(rs!recordact, 6, 3) & Space(3) & mcadrem
                       
                       If Trim(rs!recordiasfalto) = "" Then
                          cadefectivo = Chr(179) & "   POR" & Space(3) & Mid(rs!recordact, 4, 2) & " MES " & Mid(rs!recordact, 6, 3) & " DIAS" & Space(8) & Chr(179)
                       Else
                          cadefectivo = Chr(179) & "   POR" & Space(3) & Mid(rs!recordiasfalto, 4, 2) & " MES " & Mid(rs!recordiasfalto, 6, 3) & " DIAS" & Space(8) & Chr(179)
                       End If
                    Else
                       mcad = "Del : " & Format(DateAdd("m", -6, DateAdd("d", 1, mfec)), "dd/mm/yyyy") & "  Al : " & mfec & Space(3) & "No De Meses : " & Mid(rs!recordact, 4, 2) & "  No de Dias : " & Mid(rs!recordact, 6, 3) & Space(3) & mcadrem
                       If Trim(rs!recordiasfalto) = "" Then
                        cadefectivo = Chr(179) & "   POR" & Space(3) & Mid(rs!recordact, 4, 2) & " MES " & Mid(rs!recordact, 6, 3) & " DIAS" & Space(8) & Chr(179)
                       Else
                        cadefectivo = Chr(179) & "   POR" & Space(3) & Mid(rs!recordiasfalto, 4, 2) & " MES " & Mid(rs!recordiasfalto, 6, 3) & " DIAS" & Space(8) & Chr(179)
                       End If
                    End If
                    mcad = Chr(179) & " 1 " & Chr(179) & mcad
                    Print #1, mcad
                    Print #1, Chr(195) & String(3, Chr(196)) & Chr(197) & String(73, Chr(196)) & Chr(197) & String(26, Chr(196)) & Chr(197) & String(27, Chr(196)) & Chr(180)
               Case Is = 3
                    mcad = Chr(179) & " 2 " & Chr(179) & "Periodo no Comp.(Faltas y Subs.)" & Space(41) & mcadrem
                    Print #1, mcad
                    Print #1, Chr(195) & String(3, Chr(196)) & Chr(193) & String(73, Chr(196)) & Chr(197) & String(26, Chr(196)) & Chr(197) & String(27, Chr(196)) & Chr(180)
               Case Is = 4
                    mcad = Chr(179) & "TIEMPO EFECTIVO A LIQUIDAR (1-2)" & "          " & cadefectivo & Space(45 - 44) & mcadrem
                    Print #1, mcad
                    Print #1, Chr(195) & String(77, Chr(196)) & Chr(197) & String(26, Chr(196)) & Chr(197) & String(27, Chr(196)) & Chr(180)
               Case Is = 5
                        mcad = Chr(179) & "REM.COMPUT.MENSUAL PARA LIQUID.EL PER." & Chr(179) & Space(18) & fCadNum(rs!totalremun, "##,###,##0.00") & Space(20 - Len(fCadNum(rs!totalremun, "##,###,##0.00"))) & mcadrem
                       Print #1, mcad
                       Print #1, Chr(195) & String(77, Chr(196)) & Chr(197) & String(26, Chr(196)) & Chr(197) & String(27, Chr(196)) & Chr(180)
               Case Is = 6
                        mcad = Chr(179) & "Un dozavo de la Rem.compt.mensual     " & Chr(179) & Space(18) & fCadNum((rs!totalremun / 12), "##,###,##0.00") & Space(20 - Len(fCadNum((rs!totalremun / 12), "##,###,##0.00"))) & mcadrem
                       Print #1, mcad
                       Print #1, Chr(195) & String(77, Chr(196)) & Chr(197) & String(26, Chr(196)) & Chr(197) & String(27, Chr(196)) & Chr(180)
               Case Is = 7
                       mcad = Chr(179) & "Un 30avo del 12avo de la Rem. Compt.  " & Chr(179) & Space(18) & fCadNum(((rs!totalremun / 12) / 30), "##,###,##0.00") & Space(20 - Len(fCadNum(((rs!totalremun / 12) / 30), "##,###,##0.00"))) & mcadrem
                       Print #1, mcad
                       Print #1, Chr(195) & String(77, Chr(196)) & Chr(197) & String(26, Chr(196)) & Chr(197) & String(27, Chr(196)) & Chr(180)
               Case Is = 8
                    mcad = Chr(179) & Space(77) & mcadrem
                    Print #1, mcad
                    Print #1, Chr(195) & String(77, Chr(196)) & Chr(197) & String(26, Chr(196)) & Chr(197) & String(27, Chr(196)) & Chr(180)
               Case Is = 9
                    mcad = Chr(179) & Space(77) & mcadrem
                    Print #1, mcad
                    Print #1, Chr(195) & String(77, Chr(196)) & Chr(197) & String(26, Chr(196)) & Chr(197) & String(27, Chr(196)) & Chr(180)
               Case Is = 10
                    mcad = Chr(179) & Space(77) & mcadrem
                    Print #1, mcad
                    Print #1, Chr(195) & String(77, Chr(196)) & Chr(197) & String(26, Chr(196)) & Chr(197) & String(27, Chr(196)) & Chr(180)
        End Select
    Next
    If rs!diasctsfalto > 0 Then
        mctsfalta = "DIAS FALTA: " & Space(4 - Len(Trim(rs!diasctsfalto))) & rs!diasctsfalto
    Else
        mctsfalta = "                "
    End If
    If rs!dias_subs > 0 Then
        mctsdiassubs = "DIAS SUBS.: " & Space(4 - Len(Trim(rs!dias_subs))) & rs!dias_subs
    Else
        mctsdiassubs = "                "
    End If
    
    
    mcadrem = Chr(179)
    Sql = "select descripcion from placonstante where cia='" & wcia & "' and tipomovimiento='02' and status<>'*' and codinterno='" & Right(rs(I).Name, 2) & "'"
    mcadrem = mcadrem & "TOTAL" & Space(21)
    mcadrem = mcadrem & Chr(179) & Space(12) & fCadNum(rs!totalremun, "##,###,##0.00") & Space(2) & Chr(179)
    mcad = Chr(179) & Space(77) & mcadrem
    Print #1, mcad
    Print #1, Chr(195) & String(77, Chr(196)) & Chr(193) & String(26, Chr(196)) & Chr(193) & String(27, Chr(196)) & Chr(180)
    Print #1, Chr(179) & mctsfalta & " " & mctsdiassubs & Space(12) & "LIQUIDACION DE LAS CTS CON EFECTO CANCELATORIO" & Space(41) & Chr(179)
    Print #1, Chr(195) & String(32, Chr(196)) & Chr(194) & String(71, Chr(196)) & Chr(194) & String(27, Chr(196)) & Chr(180)
    Print #1, Chr(179) & Space(3) & "TIEMPO EFECTIVO A LIQUIDAR" & Space(3) & Chr(179) & Space(27) & "CALCULO DE LA CTS" & Space(27) & Chr(179) & "            MONTO          " & Chr(179)
    Print #1, Chr(195) & String(32, Chr(196)) & Chr(197) & String(71, Chr(196)) & Chr(197) & String(27, Chr(196)) & Chr(180)
    Print #1, cadefectivo & Space(20) & "TOTAL CTS DEPOSITADA O PAGADA :" & Space(20) & Chr(179) & Space(12) & fCadNum(rs!monto_total, "##,###,##0.00") & Space(2) & Chr(179)
    
    
    'MODIFICACION
    Print #1, Chr(195) & String(32, Chr(196)) & Chr(197) & String(71, Chr(196)) & Chr(197) & String(27, Chr(196)) & Chr(180)
    Cadena = "(T.CAMBIO -> " & Format(T_dCambio, "##,###,##0.000") & ") :"
    If Trim(rs!ctsmoneda) <> wmoncont Then
        Cadena = Chr(179) & Space(32) & Chr(179) & Space(10) & "MONTO DEL DEPOSITO EN $ " & Cadena & Space(37 - Len(Cadena)) & Chr(179) & Space(10) & fCadNum(rs!monto_total / T_dCambio, "$ ##,###,##0.00") & Space(2) & Chr(179)
    Else
        Cadena = Chr(179) & Space(32) & Chr(179) & Space(71) & Chr(179) & Space(27) & Chr(179)
    End If
    Print #1, Cadena
    Print #1, Chr(195) & String(32, Chr(196)) & Chr(197) & String(71, Chr(196)) & Chr(197) & String(27, Chr(196)) & Chr(180)
    Cadena = "Fecha de Deposito : " & Format(Fecha_dDeposito, "dd/MM/yyyy")
    Cadena = Chr(179) & Space(1) & Cadena & Space(1) & Chr(179) & Space(26) & "INTERES MORATORIO :" & Space(26) & Chr(179) & Space(27) & Chr(179)
    Print #1, Cadena
    Print #1, Chr(195) & String(32, Chr(196)) & Chr(197) & String(71, Chr(196)) & Chr(197) & String(27, Chr(196)) & Chr(180)
    If Trim(rs!ctsmoneda) <> wmoncont Then
        Cadena = Chr(179) & Space(32) & Chr(179) & Space(20) & "TOTAL CTS DEPOSITADA O PAGADA :" & Space(20) & Chr(179) & Space(10) & fCadNum(rs!monto_total / T_dCambio, "$ ##,###,##0.00") & Space(2) & Chr(179)
    Else
        Cadena = Chr(179) & Space(32) & Chr(179) & Space(20) & "TOTAL CTS DEPOSITADA O PAGADA :" & Space(20) & Chr(179) & Space(12) & fCadNum(rs!monto_total, "##,###,##0.00") & Space(2) & Chr(179)
    End If
    Print #1, Cadena
    Print #1, Chr(195) & String(32, Chr(196)) & Chr(197) & String(71, Chr(196)) & Chr(197) & String(27, Chr(196)) & Chr(180)
    'FIM DE MODIFICACION
    
    
'    Print #1, Chr(195) & String(32, Chr(196)) & Chr(197) & String(71, Chr(196)) & Chr(197) & String(27, Chr(196)) & Chr(180)
'    mcad = Chr(179) & Space(32) & Chr(179) & Space(20) & Space(27) & Space(24) & Chr(179) & Space(27) & Chr(179)
'    Print #1, mcad
'    Print #1, Chr(195) & String(32, Chr(196)) & Chr(193) & String(27, Chr(196)) & Chr(194) & String(43, Chr(196)) & Chr(193) & String(27, Chr(196)) & Chr(180)
    Print #1, Chr(179) & "OBSERVACIONES :" & Space(45) & Chr(179) & " CONSTANCIA DE RECEPCION :" & Space(45) & Chr(179)
    Print #1, Chr(179) & Space(60) & Chr(179) & " 1)De la Presente Liquidacion" & Space(42) & Chr(179)
    Print #1, Chr(179) & Space(60) & Chr(179) & " 2)Del Documento que Acredita el Deposito de la CTS antes detalleda" & Space(4) & Chr(179)
    Print #1, Chr(195) & String(11, Chr(196)) & Chr(194) & String(48, Chr(196)) & Chr(180) & Space(71) & Chr(179)
    Print #1, Chr(179) & "VoB." & Space(7) & Chr(179) & Space(48) & Chr(179) & Space(71) & Chr(179)
    Print #1, Chr(179) & Space(11) & Chr(179) & Space(48) & Chr(179) & Space(71) & Chr(179)
    Print #1, Chr(179) & Space(11) & Chr(179) & Space(48) & Chr(179) & Space(71) & Chr(179)
    Print #1, Chr(179) & Space(11) & Chr(179) & String(48, "-") & Chr(179) & String(71, "-") & Chr(179)
    Print #1, Chr(179) & Space(11) & Chr(179) & "Nombres,Apellidos y Cargo del Representante del " & Chr(179) & Space(17) & "Firma del Trabajador (Huella Digital)" & Space(17) & Chr(179)
    Print #1, Chr(179) & Space(11) & Chr(179) & "            Empleador(Sello y Firma)            " & Chr(179) & Space(23) & mnombre & Space(7) & Chr(179)
    Print #1, Chr(192) & String(11, Chr(196)) & Chr(193) & String(48, Chr(196)) & Chr(193) & String(71, Chr(196)) & Chr(217)
    
    Print #1, SaltaPag
   rs.MoveNext
Loop
Close #1
Call Imprime_Txt("CertCts.txt", RUTA$)

End Sub

Private Function Salva_Certificado_CTS(ByRef mFecha_dDeposito As Date, ByRef mT_dCambio As Double) As Boolean
On Error GoTo MyErr
FrmCTS_Info_Anexa.txt_T_Cambio = mT_dCambio
FrmCTS_Info_Anexa.dtp_Fecha = mFecha_dDeposito

FrmCTS_Info_Anexa.iYear = CInt(Txtano.Text): FrmCTS_Info_Anexa.iMonth = Cmbmes.ListIndex + 1
FrmCTS_Info_Anexa.Show 1
Salva_Certificado_CTS = FrmCTS_Info_Anexa.bBoolean
MyErr:
End Function
Private Sub PROVICIONES_CTS()
Dim sSQL As String
Dim MAXROW As Long, MAXCOL As Integer, MaxColInicial As Integer
Dim rs As ADODB.Recordset, rsAux As ADODB.Recordset
Dim CantMes As String, Campo As String
Dim CantMesCtsFalto As String
Dim FecIni As String, FecFin As String, FecProceso As String
Dim I As Integer, MaxColTemp As Integer
Dim dblFactor As Currency, Cadena As String
Dim Factor_EsSalud As Currency, totaportes As Currency
Dim sCol As Integer, curfactor As Currency
Dim sSQLI As String, sSQLP As String
Dim diasctsfalto As Integer
Dim diasctsSubsidio As Integer
Dim nVecesPer As Integer
Dim Fec_Auxiliar As String

Dim mFec_Inicial As String
Dim mFec_Final As String
Dim mdia As String
Dim mFecProceso As String
Dim NroTrans As Integer
'On Error GoTo ErrorTrans
Dim Reint As Currency
Reint = 0#
NroTrans = 0

Const COL_CODIGO = 0
Const COL_FECING = 1
Const COL_AREA = 2

Call Captura_Tipo_Trabajador
DoEvents
panel.Caption = "Preparando para generar Provisión..."

panel.Visible = True
panel.ZOrder 0
Me.Refresh
pBar.Min = 0
pBar.Value = 0

MAXCOL = 2
I = MAXCOL + 1
FecProceso = Format(DateAdd("d", -1, Format(DateAdd("m", 1, "01/" & Cmbmes.ListIndex + 1 & "/" & Txtano.Text), "dd/mm/yyyy")), "dd/mm/yyyy")

'IMPLEMENTACION GALLOS

nVecesPer = IIf(wGrupoPla = "01", IIf(Format(FecProceso, "yyyymm") > "201110", "3", "1"), "3")

Select Case Month(FecProceso)
    Case 1: FecIni = "01/11/" & Val(Txtano.Text) - 1
    Case 2: FecIni = "01/11/" & Val(Txtano.Text) - 1
    Case 3: FecIni = "01/11/" & Val(Txtano.Text) - 1
    Case 4: FecIni = "01/11/" & Val(Txtano.Text) - 1
    Case 5: FecIni = "01/05/" & Val(Txtano.Text)
    Case 6: FecIni = "01/05/" & Val(Txtano.Text)
    Case 7: FecIni = "01/05/" & Val(Txtano.Text)
    Case 8: FecIni = "01/05/" & Val(Txtano.Text)
    Case 9: FecIni = "01/05/" & Val(Txtano.Text)
    Case 10: FecIni = "01/05/" & Val(Txtano.Text)
    Case 11: FecIni = "01/11/" & Val(Txtano.Text)
    Case 12: FecIni = "01/11/" & Val(Txtano.Text)
End Select

Fec_Auxiliar = FecIni
FecFin = FecProceso
Erase ArrReporte

sSQL = "select distinct factor from platasaanexo where status!='*' and tipomovimiento='02' and basecalculo=16 and cia='" & wcia & "'"
If (fAbrRst(rs, sSQL)) Then
    dblFactor = rs(0)
    rs.Close
End If

sSQL = "SELECT 'I'+campo as CampUnion,sn_promedio,0 AS factor,sn_carga,formula FROM estructura_provisiones WHERE tipo='C' and "
sSQL = sSQL & " campo NOT IN (SELECT codinterno FROM platasaanexo WHERE cia='" & wcia & "' and tipomovimiento='02' and basecalculo='16')"
sSQL = sSQL & " UNION ALL SELECT 'P'+campo as CampUnion,sn_promedio,b.factor,sn_carga,formula  FROM estructura_provisiones a INNER JOIN platasaanexo b ON"
sSQL = sSQL & " (b.cia='" & wcia & "' and b.tipomovimiento='02' and b.basecalculo='16' and b.status <> '*' and b.codinterno=a.campo) WHERE a.tipo='C'"

If (fAbrRst(rs, sSQL)) Then
   pBar.Max = rs.RecordCount
   Do While Not rs.EOF
      MAXCOL = MAXCOL + 1
      rs.MoveNext
   Loop

   rs.MoveFirst
   MaxColTemp = MAXCOL + 1
   MAXCOL = MAXCOL + 9
    
   ReDim Preserve ArrReporte(0 To MAXCOL, 0 To MAXROW)
   Cadena = ""
    
   Dim FiltroFecIng As String
    
   Do While Not rs.EOF
      pBar.Value = rs.AbsolutePosition
      ArrReporte(I, MAXROW) = Trim(rs!CAMPUNION)
        
      FiltroFecIng = IIf(wGrupoPla = "01", " AND LEFT(CONVERT(VARCHAR,ph.fechaproceso,112),8) >= LEFT(CONVERT(VARCHAR,a.fingreso,112),8)", "")

      If CInt(rs!sn_carga) <> 0 Then
         If Mid(rs!CAMPUNION, 2, 2) = "29" Then
            Cadena = Cadena & "CASE WHEN ( SELECT COUNT(ISNULL(PH.I29,0)) AS CONTEO " & _
            " FROM PLAHISTORICO PH" & _
            " WHERE PH.CIA ='" & wcia & "'" & _
            " AND ph.status!='*'" & _
            " AND ph.FECHAPROCESO>= dbo.SP_VAL_FECHA_INGRESO_INICIO('" & wcia & "', '" & FecIni & "', a.placod)" & _
            " AND ph.FECHAPROCESO<='" & FecFin & "'" & _
            " AND ph.placod=A.PLACOD" & _
            " AND PH.PROCESO='01'" & _
            " AND PH.I29<>0 ) >=3 THEN " & _
            "ISNULL((SELECT SUM(COALESCE(I29,0)) FROM plahistorico ph WHERE ph.cia='" & wcia & "' and ph.status!='*'  and  ph.FECHAPROCESO>= dbo.SP_VAL_FECHA_INGRESO_INICIO('" & wcia & "', '" & FecIni & "', a.placod) AND ph.FECHAPROCESO<='" & FecFin & "' and ph.placod=a.placod AND PH.PROCESO='01')/6,0) ELSE 0 END as P29"
         Else
            Dim FECINICIO As String
            If wGrupoPla = "01" Then
               FECINICIO = Format(DateAdd("d", 0, Format(DateAdd("m", -5, "01/" & Cmbmes.ListIndex + 1 & "/" & Txtano.Text), "dd/mm/yyyy")), "dd/mm/yyyy")
                    
               If Trim(rs!Formula) & "" = "" Then
                  Cadena = Cadena & IIf(CInt(rs!sn_promedio) = 0, "I", "P") & Mid(rs!CAMPUNION, 2, 2) & "=CASE WHEN dbo.fc_validapagosProvision('" & wcia & "',a.placod,'" & FecIni & "','" & FecFin & "','" & "B" & "')>=" & nVecesPer & " THEN (SELECT ISNULL(SUM(COALESCE(I" & Trim(Mid(rs!CAMPUNION, 2, 2)) & ",0)),0) FROM plahistorico ph WHERE ph.cia='" & wcia & "' and ph.status!='*'  and  ph.FECHAPROCESO>='" & FecIni & "' " & _
                  "AND ph.FECHAPROCESO<='" & FecFin & "'" & FiltroFecIng & " AND ph.placod=a.placod AND PH.PROCESO='01')"
               Else
                  Cadena = Cadena & IIf(CInt(rs!sn_promedio) = 0, "I", "P") & Mid(rs!CAMPUNION, 2, 2) & "=CASE WHEN dbo.fc_validapagosProvision('" & wcia & "',a.placod,'" & FecIni & "','" & FecFin & "','" & "E" & "')>=" & nVecesPer & " THEN (SELECT ISNULL(SUM(" & rs!Formula & "),0) FROM plahistorico ph WHERE ph.cia='" & wcia & "' and ph.status!='*'  and  ph.FECHAPROCESO>='" & FecIni & "' " & _
                  "AND ph.FECHAPROCESO<='" & FecFin & "'" & FiltroFecIng & " AND ph.placod=a.placod AND PH.PROCESO='01')"
               End If
                
            Else
               If Trim(Mid(rs!CAMPUNION, 1, 1)) = "P" Or CBool(rs!sn_promedio) = True Then
                  mFecProceso = DateAdd("m", 1, CDate(FecProceso))
                  mFec_Inicial = Fecha_Promedios(CInt(dblFactor), mFecProceso)
                  If Val(Mid(mFecProceso, 4, 2)) = 1 Then
                     mdia = Ultimo_Dia(12, Val(Mid(mFecProceso, 7, 4)) - 1)
                     mFec_Final = Format(mdia, "00") & "/12/" & Format(Val(Mid(mFecProceso, 7, 4)) - 1, "0000")
                  Else
                     mdia = Ultimo_Dia(Val(Mid(mFecProceso, 4, 2) - 1), Val(Mid(mFecProceso, 7, 4)))
                     mFec_Final = Format(mdia, "00") & "/" & Format(Val(Mid(mFecProceso, 4, 2) - 1), "00") & "/" & Mid(mFecProceso, 7, 4)
                  End If
                        
                  Cadena = Cadena & Chr(13) & IIf(CInt(rs!sn_promedio) = 0, "I", "P") & Mid(rs!CAMPUNION, 2, 2) & "=CASE WHEN dbo.FC_PAGOS_VALIDOS('" & wcia & "',dbo.SP_VAL_FECHA_INGRESO_INICIO('" & wcia & "', '" & mFec_Inicial & "', a.placod),'" & mFec_Final & "',a.placod, 'i" & Trim(Mid(rs!CAMPUNION, 2, 2)) & "')>=3 THEN (SELECT SUM(COALESCE(I" & Trim(Mid(rs!CAMPUNION, 2, 2)) & ",0)) FROM plahistorico ph WHERE ph.cia='" & wcia & "' and ph.status!='*'  and  ph.FECHAPROCESO>=dbo.SP_VAL_FECHA_INGRESO_INICIO('" & wcia & "', '" & mFec_Inicial & "', a.placod) " & _
                  "AND ph.FECHAPROCESO<='" & mFec_Final & "' and ph.placod=a.placod AND PH.PROCESO='01')"
               Else
                  FECINICIO = Format(DateAdd("d", 0, Format(DateAdd("m", -5, "01/" & Cmbmes.ListIndex + 1 & "/" & Txtano.Text), "dd/mm/yyyy")), "dd/mm/yyyy")
                  Cadena = Cadena & Chr(13) & IIf(CInt(rs!sn_promedio) = 0, "I", "P") & Mid(rs!CAMPUNION, 2, 2) & "=CASE WHEN dbo.FC_PAGOS_VALIDOS('" & wcia & "',dbo.SP_VAL_FECHA_INGRESO_INICIO('" & wcia & "', '" & FecIni & "', a.placod),'" & FecFin & "',a.placod, 'i" & Trim(Mid(rs!CAMPUNION, 2, 2)) & "')>=3 THEN (SELECT SUM(COALESCE(I" & Trim(Mid(rs!CAMPUNION, 2, 2)) & ",0)) FROM plahistorico ph WHERE ph.cia='" & wcia & "' and ph.status!='*'  and  ph.FECHAPROCESO>=dbo.SP_VAL_FECHA_INGRESO_INICIO('" & wcia & "', '" & FecIni & "', a.placod) " & _
                  "AND ph.FECHAPROCESO<='" & FecFin & "' and ph.placod=a.placod AND PH.PROCESO='01')"
               End If
                
            End If
         End If
            
         If CInt(rs!sn_promedio) = -1 And Trim(Mid(rs!CAMPUNION, 1, 1)) = "A" Then
            Cadena = Cadena & "/(SELECT aportacion FROM PLACONSTANTE WHERE tipomovimiento='03' AND cia='" & wcia & "' AND status!='*' AND codinterno='" & Mid(rs!CAMPUNION, 2, 2) & "'),"
         ElseIf CInt(rs!sn_promedio) = -1 And (Trim(Mid(rs!CAMPUNION, 1, 1)) = "I" Or Trim(Mid(rs!CAMPUNION, 1, 1)) = "P") Then
            If Mid(rs!CAMPUNION, 2, 2) = "29" Then
               Cadena = Cadena & ","
            Else
               If wGrupoPla = "01" Then
                  If Val(rs!factor) <> 0 Then
                     Cadena = Cadena & "/" & dblFactor & " ELSE 0 END,"
                  Else
                     Cadena = Cadena & "/" & dblFactor & " ELSE 0 END,"
                  End If
               Else
                  If Trim(Mid(rs!CAMPUNION, 1, 1)) = "P" Or CBool(rs!sn_promedio) = True Then
                     mFecProceso = DateAdd("m", 1, CDate(FecProceso))
                     mFec_Inicial = Fecha_Promedios(CInt(dblFactor), mFecProceso)
                     If Val(Mid(mFecProceso, 4, 2)) = 1 Then
                        mdia = Ultimo_Dia(12, Val(Mid(mFecProceso, 7, 4)) - 1)
                        mFec_Final = Format(mdia, "00") & "/12/" & Format(Val(Mid(mFecProceso, 7, 4)) - 1, "0000")
                     Else
                        mdia = Ultimo_Dia(Val(Mid(mFecProceso, 4, 2) - 1), Val(Mid(mFecProceso, 7, 4)))
                        mFec_Final = Format(mdia, "00") & "/" & Format(Val(Mid(mFecProceso, 4, 2) - 1), "00") & "/" & Mid(mFecProceso, 7, 4)
                     End If
                     'Cadena = Cadena & "/ dbo.SP_VAL_FACTOR_MESES('" & wcia & "', '" & mFec_Inicial & "', '" & mFec_Final & "', a.placod) ELSE 0 END,"
                     Cadena = Cadena & "/" & dblFactor & " ELSE 0 END,"
                  Else
                     Cadena = Cadena & "/ dbo.SP_VAL_FACTOR_MESES('" & wcia & "', '" & FecIni & "', '" & FecFin & "', a.placod) ELSE 0 END,"
                  End If
               End If
            End If
         Else
            Cadena = Cadena & ","
         End If
      Else
         If Trim(rs!CAMPUNION & "") = "I02" Then 'Asignacion Familiar si esta por semanas entonces (AsigFam*4)/240
            Cadena = Cadena & "COALESCE((SELECT top 1 Case factor_horas When 48 then (prb.importe*4)/(240/" & hORAS_X_DIA & ") else prb.importe/(factor_horas/" & hORAS_X_DIA & " ) End FROM plaremunbase prb WHERE cia='" & wcia & "' and placod=a.placod AND status!='*' and concepto='" & Trim(Mid(rs!CAMPUNION, 2, 2)) & "'),0) as '" & Trim(rs!CAMPUNION) & "',"
         Else
            Cadena = Cadena & "COALESCE((SELECT top 1 prb.importe/(factor_horas/" & hORAS_X_DIA & ") FROM plaremunbase prb WHERE cia='" & wcia & "' and placod=a.placod AND status!='*' and concepto='" & Trim(Mid(rs!CAMPUNION, 2, 2)) & "'),0) as '" & Trim(rs!CAMPUNION) & "',"
         End If
         
      End If
      I = I + 1
      rs.MoveNext
   Loop
   Cadena = Mid(Cadena, 1, Len(Trim(Cadena)) - 1)
   rs.Close
End If

sSQL = "SET DATEFORMAT DMY SELECT"
sSQL = sSQL & " a.placod,LTRIM(RTRIM(a.ap_pat))+' '+LTRIM(RTRIM(a.ap_mat))+' '+LTRIM(RTRIM(a.ap_cas))+' '+LTRIM(RTRIM(a.nom_1))+' '+LTRIM(RTRIM(a.nom_2)),"
sSQL = sSQL & " fingreso , fcese, a.area,a.tipotrabajador,a.cod_area,"
sSQL = sSQL & " a.ctsbanco, a.ctsmoneda, a.ctstipcta, a.ctsnumcta,"
sSQL = sSQL & Cadena
sSQL = sSQL & " FROM planillas a LEFT OUTER JOIN plahistorico ph ON (ph.cia='" & wcia & "' and ph.status!='*' "

'If wGrupoPla = "01" Then
    sSQL = sSQL & " and ph.FECHAPROCESO>='" & FecIni & "' AND ph.FECHAPROCESO<='" & FecFin & "' and ph.placod=a.placod) "
'Else
'    sSQL = sSQL & " and ph.FECHAPROCESO>=dbo.SP_VAL_FECHA_INGRESO_INICIO('" & wcia & "', '" & FecIni & "', a.placod) AND ph.FECHAPROCESO<='" & FecFin & "' and ph.placod=a.placod) "
'End If

Select Case Format(Cmbmes.ListIndex + 2, "00")
    Case "02": s_Dia_Proceso_Prov = "28"
    Case Else: s_Dia_Proceso_Prov = "30"
End Select

Dim Fecha_CI As Date
Dim FechaCF As Date
Fecha_CI = DateSerial(Txtano.Text, (Cmbmes.ListIndex + 1) + 0, 1)
FechaCF = DateSerial(Txtano.Text, (Cmbmes.ListIndex + 1) + 1, 0)


'sSQL = sSQL & " WHERE a.cat_trab<>'04' and a.status!='*' and a.cia='" & wcia & "' and (a.fcese >='" & DateAdd("d", -1, DateAdd("m", 1, "01" & "/" & Format(CmbMes.ListIndex + 1, "00") & "/" & Txtano.Text)) & "' or a.fcese is null)"
'sSQL = sSQL & " WHERE a.cat_trab<>'04' and a.status!='*' and a.cia='" & wcia & "' and (a.fcese >='" & DateAdd("d", -1, DateAdd("m", 1, "01" & "/" & Format(Cmbmes.ListIndex + 1, "00") & "/" & Txtano.Text)) & "' AND DATEDIFF(DAY, A.FINGRESO, A.FCESE) >= 30 or a.fcese is null)"

'sSQL = sSQL & " WHERE a.placod not in (select placod from dbo.trab_Suspencion) and a.cat_trab<>'04' and a.status!='*' and a.cia='" & wcia & "' and (a.fcese >'" & DateAdd("d", -1, DateAdd("m", 1, "01" & "/" & Format(Cmbmes.ListIndex + 1, "00") & "/" & Txtano.Text)) & "' or a.fcese is null )"
'sSQL = sSQL & " WHERE a.cat_trab<>'04' and a.status!='*' and a.cia='" & wcia & "' and (a.fcese >'" & DateAdd("d", -1, DateAdd("m", 1, "01" & "/" & Format(Cmbmes.ListIndex + 1, "00") & "/" & Txtano.Text)) & "' or a.fcese is null )"
sSQL = sSQL & " WHERE a.cat_trab<>'04' and a.status!='*' and a.cia='" & wcia & "' and (a.fcese >'" & DateAdd("d", -1, DateAdd("m", 1, "01" & "/" & Format(Cmbmes.ListIndex + 1, "00") & "/" & Txtano.Text)) & "' or a.fcese is null )"

Select Case s_TipoTrabajador_Cts
    Case "01": sSQL = sSQL & " and a.tipotrabajador='01'"
    Case "02": sSQL = sSQL & " and a.tipotrabajador='02'"
End Select
sSQL = sSQL & " GROUP BY a.PLACOD , a.ap_pat, a.ap_mat, a.ap_cas, a.nom_1, a.nom_2, a.fingreso, a.fcese,a.area,a.tipotrabajador"
sSQL = sSQL & ",a.ctsbanco, a.ctsmoneda, a.ctstipcta, a.ctsnumcta,a.cod_area"

Dim Banco_CTS As String
Dim tMoneda_CTS As String
Dim tCta_CTS As String
Dim NumCta_CTS As String

Dim ltmpdiasfaltas As Integer
Dim ltmpmesesfaltas As Integer

MAXROW = MAXROW + 1
Dim PROMEDIO_HRS As Integer
DoEvents
panel.Caption = "Generando Provisión ..."
cn.BeginTrans
NroTrans = 1

If (fAbrRst(rs, sSQL)) Then
   pBar.Min = 0
   pBar.Value = 0
   pBar.Max = rs.RecordCount
    
   rs.MoveFirst
   Do While Not rs.EOF
      FiltroFecIng = " AND LEFT(CONVERT(VARCHAR,fechaproceso,112),8) >= '" & Format(rs!fIngreso, "yyyymmdd") & "'"
      Banco_CTS = Empty: tMoneda_CTS = Empty: tCta_CTS = Empty: NumCta_CTS = Empty
      Banco_CTS = rs!ctsbanco
      tMoneda_CTS = rs!ctsmoneda
      tCta_CTS = rs!ctstipcta
      NumCta_CTS = rs!ctsnumcta
      pla_area = rs!cod_area
      pBar.Value = rs.AbsolutePosition
      Debug.Print rs.AbsolutePosition & " " & Trim(rs!PlaCod)
      PROMEDIO_HRS = 0
    
      If Year(rs!fIngreso) < Txtano Then GoTo procesa_ProV_CTS_Nueva
      If Month(rs!fIngreso) <= Cmbmes.ListIndex + 1 Then
procesa_ProV_CTS_Nueva:
          
         ReDim Preserve ArrReporte(0 To MAXCOL, 0 To MAXROW)
         CantMes = ""
         'ACP
         'If Trim(rs!PlaCod) = "O6158" Then Stop
         CantMes = CalcularMeses1(FecIni, FecFin, Format(Day(rs!fIngreso), "00") & "/" & Format(Month(rs!fIngreso), "00") & "/" & Format(Year(rs!fIngreso), "0000"), rs!PlaCod)
         'diasctsfalto = CalcularDiasCtsFalto(FecIni, FecFin, rs!PlaCod)
         'Cambios Junio 2020
        
         Dim cad As String
 
         diasctsfalto = CalcularDiasCtsFalto(Txtano.Text, Cmbmes.ListIndex + 1, rs!PlaCod, "F")
         
         diasctsSubsidio = CalcularDiasCtsFalto(Txtano.Text, Cmbmes.ListIndex + 1, rs!PlaCod, "S")
         'CantMesCtsFalto = CalcularMesesCtsFalto(CantMes, diasctsfalto)
          CantMesCtsFalto = CantMes
         ArrReporte(COL_CODIGO, MAXROW) = Trim(rs!PlaCod)
         ArrReporte(COL_FECING, MAXROW) = rs!fIngreso
         ArrReporte(COL_AREA, MAXROW) = rs!Area
         For I = 6 To rs.Fields.count - 1
            sCol = BuscaColumna(rs.Fields(I).Name, MAXCOL)
            If sCol > 0 Then
               If Trim(rs!TipoTrabajador) = "01" Then
                  Select Case Mid(rs.Fields(I).Name, 1, 1)
                     Case "I"
                        ArrReporte(sCol, MAXROW) = Round(rs.Fields(I).Value * DIAS_TRABAJO, 2)
                     Case "P"
                        If PROMEDIO_HRS = 0 Then
                           If Mid(rs.Fields(11).Name, 1, 3) = "P29" Then
                              ArrReporte(sCol, MAXROW) = Round(rs.Fields(11).Value, 2)
                              PROMEDIO_HRS = PROMEDIO_HRS + 1
                           Else
                              ArrReporte(sCol, MAXROW) = Round(rs.Fields(I).Value, 2)
                           End If
                        Else
                           ArrReporte(sCol, MAXROW) = Round(rs.Fields(I).Value, 2)
                        End If
                  End Select
               Else
                  Select Case Mid(rs.Fields(I).Name, 1, 1)
                     Case "I"
                        ArrReporte(sCol, MAXROW) = Round(rs.Fields(I).Value, 2)
                     Case "P"
                        ArrReporte(sCol, MAXROW) = Round(rs.Fields(I).Value / DIAS_TRABAJO, 2)
                        If Mid(rs.Fields(I).Name, 1, 3) = "P11" Then
                        'If I = 11 Then
                            Reint = 0#
                             ArrReporte(sCol, MAXROW) = Round(ArrReporte(sCol, MAXROW) + Reint, 2)
                        End If
                        If Mid(rs.Fields(I).Name, 1, 3) = "P37" Then
                        'If I = 37 Then
                            Reint = 0#
 
                            ArrReporte(sCol, MAXROW) = Round(ArrReporte(sCol, MAXROW) + Reint, 2)
                        End If
                  End Select
               End If
               totaportes = totaportes + IIf(IsNull(ArrReporte(sCol, MAXROW)) = True, 0, ArrReporte(sCol, MAXROW))
            End If
         Next

         I = MaxColTemp
        
         ArrReporte(I, MAXROW) = CantMes
        
         ' PROMEDIO GRATIFICACION
         I = I + 1
         ArrReporte(I, 0) = "P15"
            
         If Cmbmes.ListIndex + 1 > 6 Then
            Sql = "select (ISNULL(totaling,0)-ISNULL(i30,0)) from plahistorico where cia='" & wcia & "' and proceso='03' and placod='" & Trim(rs!PlaCod) & "' and status!='*' and month(fechaproceso)=7 and YEAR(fechaproceso)=" & Txtano.Text & FiltroFecIng & " "
         Else
            Sql = "select (ISNULL(totaling,0)-ISNULL(i30,0)) from plahistorico where cia='" & wcia & "' and proceso='03' and placod='" & Trim(rs!PlaCod) & "' and status!='*' and month(fechaproceso)=12 and YEAR(fechaproceso)=" & Txtano.Text - 1 & FiltroFecIng & " "
         End If
            Reint = 0#
            ArrReporte(I, MAXROW) = 0#
            
         If Reint = 0# Then
            If (fAbrRst(rsAux, Sql)) Then
               If Trim(rs!TipoTrabajador) = "01" Then
                  ArrReporte(I, MAXROW) = Round((rsAux(0) / 6), 2)
               Else
                  ArrReporte(I, MAXROW) = Round((rsAux(0) / 180), 2)
               End If
               rsAux.Close
            Else
               ArrReporte(I, MAXROW) = 0
            End If
         Else
            ArrReporte(I, MAXROW) = Round(Reint, 2)
         End If
         totaportes = totaportes + ArrReporte(I, MAXROW)

         'TOTAL DE JORNAL
         I = I + 1
         ArrReporte(I, MAXROW) = totaportes
        
         'Monto por dias falta y subsidios
         Dim lMonto_Faltas As Double
         Dim lMonto_Subsido As Double
         lMonto_Subsido = 0: lMonto_Faltas = 0
         
         If rs!TipoTrabajador = "01" Then
            If diasctsfalto > 0 Then lMonto_Faltas = Round((totaportes / 360) * diasctsfalto, 2)
            If diasctsSubsidio > 0 Then lMonto_Subsido = Round((totaportes / 360) * diasctsSubsidio, 2)
         Else
            If diasctsfalto > 0 Then lMonto_Faltas = Round(((totaportes * 30) / 360) * diasctsfalto, 2)
            If diasctsSubsidio > 0 Then lMonto_Subsido = Round(((totaportes * 30) / 360) * diasctsSubsidio, 2)
         End If
        
         'TOTAL DE INDENMISATORIO
         I = I + 1
         'ArrReporte(I, MAXROW) = IIf(wGrupoPla = "01", MontoIndecnizado(CantMes, totaportes, rs!TipoTrabajador), MontoIndecnizado(CantMesCtsFalto, totaportes, rs!TipoTrabajador))
         ArrReporte(I, MAXROW) = MontoIndecnizado(CantMesCtsFalto, totaportes, rs!TipoTrabajador, lMonto_Faltas, lMonto_Subsido)
         
         'PROVISION DEL AÑO PASADO
         I = I + 1
         If Cmbmes.ListIndex + 1 < COL_PRIMERMES Then
            Sql = "select monto_total from plaprovcts where cia='" & wcia & "' and month(fechaproceso)=12 and year(fechaproceso)=" & Txtano.Text - 1 & " and placod='" & Trim(rs!PlaCod) & "' and status<>'*'" & FiltroFecIng & " "
            If (fAbrRst(rsAux, Sql)) Then
               ArrReporte(I, MAXROW) = rsAux(0)
               rsAux.Close
            Else
               ArrReporte(I, MAXROW) = 0
            End If
         Else
            ArrReporte(I, MAXROW) = 0
         End If
      
         'AJUSTE DE PROVISION
         I = I + 1
         ArrReporte(I, MAXROW) = ArrReporte(I - 2, MAXROW) - ArrReporte(I - 1, MAXROW)
        
         'PROVISION DE ESTE AÑO
         I = I + 1
         If Cmbmes.ListIndex = 0 Then
            Sql = "select 0"
         Else
            If Cmbmes.ListIndex + 1 < COL_PRIMERMES Then
               Sql = "select COALESCE(SUM(COALESCE(provision_actual,0)),0) from plaprovcts where cia='" & wcia & "' and fechaproceso>='01/01/" & Txtano.Text & "' AND fechaproceso<'01/" & Cmbmes.ListIndex + 1 & "/" & Txtano.Text & "' and placod='" & Trim(rs!PlaCod) & "' and status<>'*'" & FiltroFecIng & " "
            ElseIf Cmbmes.ListIndex + 1 = COL_PRIMERMES Or Cmbmes.ListIndex + 1 = COL_SEGUNDOMES Then
               Sql = "select 0"
            Else
               If Cmbmes.ListIndex + 1 >= COL_PRIMERMES And Cmbmes.ListIndex + 1 < COL_SEGUNDOMES Then
                  If wGrupoPla = "01" Then
                     Sql = "select COALESCE(SUM(COALESCE(provision_actual,0)),0) from plaprovcts where cia='" & wcia & "' and fechaproceso>='01/" & COL_PRIMERMES & "/" & Txtano.Text & "' AND fechaproceso<'01/" & Cmbmes.ListIndex + 1 & "/" & Txtano.Text & "' and placod='" & Trim(rs!PlaCod) & "' and status<>'*'" & FiltroFecIng & " "
                  Else
                     Sql = "select COALESCE(SUM(COALESCE(provision_actual,0)),0) from plaprovcts where cia='" & wcia & "' and fechaproceso>= '" & Val_Fec_Ingreso_Inicio("01/" & COL_PRIMERMES & "/" & Txtano.Text, rs!fIngreso) & "' AND fechaproceso<'01/" & Cmbmes.ListIndex + 1 & "/" & Txtano.Text & "' and placod='" & Trim(rs!PlaCod) & "' and status<>'*'"
                  End If
               Else
                  If wGrupoPla = "01" Then
                     Sql = "select COALESCE(SUM(COALESCE(provision_actual,0)),0) from plaprovcts where cia='" & wcia & "' and fechaproceso>='01/" & COL_SEGUNDOMES & "/" & Txtano.Text & "' AND fechaproceso<'01/" & Cmbmes.ListIndex + 1 & "/" & Txtano.Text & "' and placod='" & Trim(rs!PlaCod) & "' and status<>'*'" & FiltroFecIng & " "
                  Else
                     Sql = "select COALESCE(SUM(COALESCE(provision_actual,0)),0) from plaprovcts where cia='" & wcia & "' and fechaproceso>= '" & Val_Fec_Ingreso_Inicio("01/" & COL_SEGUNDOMES & "/" & Txtano.Text, rs!fIngreso) & "' AND fechaproceso<'01/" & Cmbmes.ListIndex + 1 & "/" & Txtano.Text & "' and placod='" & Trim(rs!PlaCod) & "' and status<>'*'"
                  End If
               End If
            End If
         End If
      
         If (fAbrRst(rsAux, Sql)) Then
            ArrReporte(I, MAXROW) = rsAux(0)
            rsAux.Close
         Else
            ArrReporte(I, MAXROW) = 0
         End If
        
         'PROVISION DEL MES
         I = I + 1
            
         If ArrReporte(I - 2, MAXROW) < ArrReporte(I - 1, MAXROW) Then
            ArrReporte(I, MAXROW) = 0
         Else
            ArrReporte(I, MAXROW) = ArrReporte(I - 2, MAXROW) - ArrReporte(I - 1, MAXROW)
         End If
            
         MAXROW = MAXROW + 1
       
         sSQLI = ""
         For I = 1 To 50
            Campo = "I" & Format(I, "00")
            sCol = BuscaColumna(Campo, MAXCOL)
            If sCol > 0 Then
               sSQLI = sSQLI & IIf(Len(Trim(ArrReporte(sCol, MAXROW - 1))) = 0, "0", ArrReporte(sCol, MAXROW - 1)) & ","
            Else
               sSQLI = sSQLI & "0,"
            End If
         Next

         sSQLP = ""
         For I = 1 To 50
            Campo = "P" & Format(I, "00")
            sCol = BuscaColumna(Campo, MAXCOL)
            If sCol > 0 Then
               sSQLP = sSQLP & ArrReporte(sCol, MAXROW - 1) & ","
            Else
               sSQLP = sSQLP & "0,"
            End If
         Next
         If ArrReporte(COL_CODIGO, MAXROW - 1) = "O6706" Then
            Print ArrReporte(COL_CODIGO, MAXROW - 1)
         End If
         
         If diasctsfalto + diasctsSubsidio > 0 Then
            ltmpdiasfaltas = Val(Mid(CantMesCtsFalto, 7, 2))
            ltmpmesesfaltas = Val(Mid(CantMesCtsFalto, 4, 2))
                       
            If ltmpdiasfaltas >= diasctsfalto + diasctsSubsidio Then
            
               CantMesCtsFalto = Mid(CantMesCtsFalto, 1, 6) & Space(2 - Len(Trim(ltmpdiasfaltas - (diasctsfalto + diasctsSubsidio)))) & ltmpdiasfaltas - (diasctsfalto + diasctsSubsidio)
               
            ElseIf ltmpmesesfaltas > 0 Then
                
                If diasctsfalto + diasctsSubsidio > 30 Then
                   ltmpmesesfaltas = ltmpmesesfaltas - Fix((diasctsfalto + diasctsSubsidio) / 30)
                   ltmpdiasfaltas = 30 - ((diasctsfalto + diasctsSubsidio) - ((Fix((diasctsfalto + diasctsSubsidio) / 30) * 30)))
                Else
                   'mgirao
                   ltmpdiasfaltas = 30 - (diasctsfalto + diasctsSubsidio) + ltmpdiasfaltas
                   'original
                   'ltmpdiasfaltas = 30 - (diasctsfalto + diasctsSubsidio)
                End If
                ltmpmesesfaltas = ltmpmesesfaltas - 1
                
                
                CantMesCtsFalto = Mid(CantMesCtsFalto, 1, 3) & " " & Space(2 - Len(ltmpmesesfaltas)) & ltmpmesesfaltas & " " & Space(2 - Len(Trim(Str(ltmpdiasfaltas)))) & ltmpdiasfaltas
            End If
         End If
         
         sSQL = ""
         sSQL = "INSERT plaprovcts VALUES ('" & wcia & "','" & ArrReporte(COL_CODIGO, MAXROW - 1) & "','" & Trim(rs!TipoTrabajador) & "','" & ArrReporte(MaxColTemp, MAXROW - 1) & "','" & ArrReporte(COL_AREA, MAXROW - 1) & "',"
         sSQL = sSQL & sSQLI & sSQLP & ArrReporte(MaxColTemp + 2, MAXROW - 1) & ",0,0,0," & ArrReporte(MaxColTemp + 3, MAXROW - 1) & "," & ArrReporte(MaxColTemp + 4, MAXROW - 1) & "," & ArrReporte(MaxColTemp + 5, MAXROW - 1) & ","
         sSQL = sSQL & ArrReporte(MaxColTemp + 6, MAXROW - 1) & "," & ArrReporte(MaxColTemp + 7, MAXROW - 1) & ",0,'" & Format(FecProceso, "DD/MM/YYYY") & "',''," & diasctsfalto & ",'" & CantMesCtsFalto & "'" ')"
         sSQL = sSQL & ",'" & Banco_CTS & "', '" & tMoneda_CTS & "', '" & tCta_CTS & "', '" & NumCta_CTS & "',null,null,'" & Format(rs!fIngreso, "DD/MM/YYYY") & "'," & lMonto_Faltas & "," & diasctsSubsidio & "," & lMonto_Subsido & ",'" & pla_area & "')"
      
         cn.Execute (sSQL)
         totaportes = 0
      End If
      rs.MoveNext
   Loop
   'Actializa centros de costos
   sSQL = "Update plaprovcts set area= "
   sSQL = sSQL & "isnull((select Top 1 ccosto from planilla_ccosto where cia=plaprovcts.cia and placod=plaprovcts.placod and status<>'*' order by porc desc),'') "
   sSQL = sSQL & "where cia='" & wcia & "' and year(fechaproceso)=" & Txtano.Text & " and month(fechaproceso)=" & Cmbmes.ListIndex + 1 & " and status<>'*'"
   cn.Execute (sSQL)
   
   'Actualiza Bancos y Cuentas de CTS
   
   DoEvents
   panel.Caption = "Provisión : Proceso Terminado"
End If

Carga_Cts
panel.Visible = False
If NroTrans = 1 Then cn.CommitTrans

Exit Sub

ErrorTrans:
    If NroTrans = 1 Then
        cn.RollbackTrans
    End If
    
    MsgBox Err.Description, vbCritical, Me.Caption
End Sub

Private Function CalcularMesesCtsFalto(ByVal DiasMes As String, diasctsfalto As Integer) As String
   
Dim ano As Integer
Dim Mes As Integer
Dim Dia As Integer

Dim ano1 As Integer
Dim Mes1 As Integer
Dim dia1 As Integer

Dim anof As Integer
Dim mesf As Integer
Dim diaf As Integer

ano = Val(Mid(DiasMes, 1, 3))
Mes = Val(Mid(DiasMes, 3, 3))
Dia = Val(Mid(DiasMes, 6, 9))

If diasctsfalto >= 30 Then
    Mes1 = Int(diasctsfalto / 30)
    dia1 = diasctsfalto Mod 30
Else
    dia1 = diasctsfalto
End If

'anof = ano - ano1
mesf = Mes - Mes1



If Dia < dia1 Then
 diaf = 30 - dia1 + Dia
 mesf = mesf - 1
Else
  diaf = Dia - dia1
End If

CalcularMesesCtsFalto = Space(2 - Len(Trim(anof))) & ano & " " & Space(2 - Len(Trim(mesf))) & mesf & " " & Space(2 - Len(Trim(diaf))) & diaf & " "

End Function
'Private Function CalcularDiasCtsFalto(fechainicio As String, fechafin As String, PlaCod As String) As Integer
Private Function CalcularDiasCtsFalto(ayo As Integer, Mes As Integer, PlaCod As String, lTipo As String) As Integer
   Dim Dias As Integer
    Dim rs As Recordset
    Set rs = New Recordset
    
'    Sql$ = "select count(fecha_falta) diasfalto " & _
'           " From cts_falta " & _
'           " Where fecha_falta" & _
'           " between convert(datetime, '" & fechainicio & "',103) and convert(datetime,'" & fechafin & "',103)" & _
'           " and rtrim(placod) ='" & Trim(PlaCod) & _
'           "' and rtrim(cia) ='" & Trim(wcia) & _
'           "'and status <>'*'"

     Sql$ = "select dias as diasfalto from pla_cts_faltas where cia='" & wcia & "' and tipo='" & lTipo & "' and ayo=" & ayo & " and Mes=" & Mes & " and placod='" & PlaCod & "' and status<>'*'"
     
     Set rs = cn.Execute(Sql$)
    
    If rs.RecordCount > 0 Then
        Dias = rs!diasfalto
    Else
        Dias = 0
    End If
    
  CalcularDiasCtsFalto = Dias
End Function

'Private Function CalcularMeses(ByVal pFecIngreso As String) As String
''Dim mesestmp As String
''Dim año As String, mes As String, Dia As String
''Dim FecIngTmp As String
''Dim pFecproc As String
''Dim pFecValida As String
''
''
''año = 0: mes = 0: Dia = 0
''If Year(CDate(pFecIngreso)) > Val(Txtano.Text) Then GoTo Salir
''
''    If Cmbmes.ListIndex + 1 >= COL_SEGUNDOMES Or Cmbmes.ListIndex + 1 < COL_PRIMERMES Then
''        pFecValida = DateAdd("m", 1, "01/" & Cmbmes.ListIndex + 1 & "/" & Txtano.Text)
''        pFecValida = DateAdd("d", -1, "01/" & Month(pFecValida) & "/" & Year(pFecValida))
''        FecIngTmp = DateAdd("d", 0, "01/" & COL_SEGUNDOMES & "/" & Year(pFecValida) - 1)
''
''        If Year(CDate(pFecIngreso)) < Year(CDate(pFecValida)) Then
''            If Month(CDate(pFecValida)) >= 1 And Month(CDate(pFecValida)) < COL_PRIMERMES Then
''            'ESTE CAMBIO ES DE EMERGENCIA SE DEBERA DE MODIFICAR DESPUES
''              ' If Day(CDate(pFecIngreso)) > 1 And Month(CDate(pFecIngreso)) >= COL_SEGUNDOMES And Year(CDate(pFecIngreso)) = IIf(Month(pFecValida) = 4, Year(pFecValida) - 1, Year(pFecValida)) Then
''              If CDate(pFecIngreso) > CDate(FecIngTmp) Then
''                    'mes = Abs(DateDiff("m", "01/" & COL_SEGUNDOMES + 1 & "/" & Txtano.Text - 1, "01/" & CmbMes.ListIndex + 1 & "/" & Txtano.Text)) + 1
''                    mes = Abs(DateDiff("m", pFecIngreso, pFecValida))
''                    If Day(pFecIngreso) = 1 Then
''                        Dia = 0
''                        mes = mes + 1
''                        pFecproc = DateAdd("d", -1, DateAdd("m", 1, pFecIngreso))
''                    Else
''                        pFecproc = DateAdd("d", -1, "01/" & Month(DateAdd("m", 1, pFecIngreso)) & "/" & Year(DateAdd("m", 1, pFecIngreso)))
''                        Dia = DateDiff("d", pFecIngreso, pFecproc)
''                        pFecproc = DateAdd("d", Val(Dia), pFecIngreso)
''                    End If
''                Else
''                    mes = Abs(DateDiff("m", "01/" & COL_SEGUNDOMES & "/" & Txtano.Text - 1, "01/" & Cmbmes.ListIndex + 1 & "/" & Txtano.Text)) + 1
''               End If
''
''                'mes = Abs(DateDiff("m", CDate(pFecIngreso), "01/" & CmbMes.ListIndex + 1 & "/" & Txtano.Text)) + 1
''            Else
''                mes = Abs(DateDiff("m", "01/" & COL_SEGUNDOMES & "/" & Txtano.Text, "01/" & Cmbmes.ListIndex + 1 & "/" & Txtano.Text)) + 1
''            End If
''
''        Else
''            año = 0
''            'Dia = Abs(DateDiff("d", DateAdd("m", mes, pFecIngreso), DateAdd("d", -1, DateAdd("m", 1, "01/" & CmbMes.ListIndex + 1 & "/" & txtano.Text))))
''
'''            pFecproc = DateAdd("d", -1, "01/" & Month(DateAdd("m", 1, pFecIngreso)) & "/" & Year(DateAdd("m", 1, pFecIngreso)))
'''            Dia = DateDiff("d", pFecIngreso, pFecproc)
'''            pFecproc = DateAdd("d", Val(Dia), pFecIngreso)
'''            mes = DateDiff("m", pFecproc, DateAdd("d", -1, DateAdd("m", 1, "01/" & Cmbmes.ListIndex + 1 & "/" & Txtano.Text)))
''
''            If Day(pFecIngreso) = 1 Then
''                Dia = 0
''                pFecproc = DateAdd("d", -1, DateAdd("m", 1, pFecIngreso))
''            Else
''                pFecproc = DateAdd("d", -1, "01/" & Month(DateAdd("m", 1, pFecIngreso)) & "/" & Year(DateAdd("m", 1, pFecIngreso)))
''                Dia = DateDiff("d", pFecIngreso, pFecproc)
''                pFecproc = DateAdd("d", Val(Dia), pFecIngreso)
''            End If
''
''            mes = DateDiff("m", pFecproc, DateAdd("d", -1, DateAdd("m", 1, "01/" & Cmbmes.ListIndex + 1 & "/" & Txtano.Text)))
''
''            If Day(pFecIngreso) = 1 Then mes = CStr(Val(mes) + 1)
''
''        End If
''    Else
''
''        '************codigo modificado giovanni 23082007******************************
''        'If pFecIngreso < CDate("01/" & COL_PRIMERMES & "/" & txtano.Text) Or pFecIngreso < CDate("01/" & COL_SEGUNDOMES & "/" & txtano.Text) Then
''        If pFecIngreso < CDate("01/" & COL_PRIMERMES & "/" & Txtano.Text) Then
''        '*****************************************************************************
''
''            mes = Abs(DateDiff("m", "01/" & COL_PRIMERMES & "/" & Txtano.Text, "01/" & Cmbmes.ListIndex + 1 & "/" & Txtano.Text)) + 1
''
''        Else
'''            año = 0
'''            mes = Abs(DateDiff("m", pFecIngreso, DateAdd("d", -1, DateAdd("m", 1, "01/" & CmbMes.ListIndex + 1 & "/" & txtano.Text))))
'''            Dia = Abs(DateDiff("d", DateAdd("m", mes, pFecIngreso), DateAdd("d", -2, DateAdd("m", 1, "01/" & CmbMes.ListIndex + 1 & "/" & txtano.Text))))
''
''            año = 0
''            If Day(pFecIngreso) = 1 Then
''                Dia = 0
''                pFecproc = DateAdd("d", -1, DateAdd("m", 1, pFecIngreso))
''            Else
''                pFecproc = DateAdd("d", -1, "01/" & Month(DateAdd("m", 1, pFecIngreso)) & "/" & Year(DateAdd("m", 1, pFecIngreso)))
''                Dia = DateDiff("d", pFecIngreso, pFecproc)
''                pFecproc = DateAdd("d", Val(Dia), pFecIngreso)
''            End If
''
''            mes = DateDiff("m", pFecproc, DateAdd("d", -1, DateAdd("m", 1, "01/" & Cmbmes.ListIndex + 1 & "/" & Txtano.Text)))
''
''            If Day(pFecIngreso) = 1 Then mes = CStr(Val(mes) + 1)
''        End If
''
''    End If
''
''Salir:
'''If mes > 6 Then mes = mes - 6
''CalcularMeses = Space(2 - Len(Trim(año))) & año & " " & Space(2 - Len(Trim(mes))) & mes & " " & Space(2 - Len(Trim(Dia))) & Dia & " "
'Dim mesestmp As String
'Dim año As String, mes As String, Dia As String
'Dim FecIngTmp As String
'Dim pFecproc As String
'Dim pFecValida As String
'
'
'año = 0: mes = 0: Dia = 0
'If Year(CDate(pFecIngreso)) > Val(Txtano.Text) Then GoTo Salir
'
'    If Cmbmes.ListIndex + 1 >= COL_SEGUNDOMES Or Cmbmes.ListIndex + 1 < COL_PRIMERMES Then
'        pFecValida = DateAdd("m", 1, "01/" & Cmbmes.ListIndex + 1 & "/" & Txtano.Text)
'        pFecValida = DateAdd("d", -1, "01/" & Month(pFecValida) & "/" & Year(pFecValida))
'        FecIngTmp = DateAdd("d", 0, "01/" & COL_SEGUNDOMES & "/" & Year(pFecValida) - 1)
'
'        If Year(CDate(pFecIngreso)) < Year(CDate(pFecValida)) Then
'            If Month(CDate(pFecValida)) >= 1 And Month(CDate(pFecValida)) < COL_PRIMERMES Then
'            'ESTE CAMBIO ES DE EMERGENCIA SE DEBERA DE MODIFICAR DESPUES
'              ' If Day(CDate(pFecIngreso)) > 1 And Month(CDate(pFecIngreso)) >= COL_SEGUNDOMES And Year(CDate(pFecIngreso)) = IIf(Month(pFecValida) = 4, Year(pFecValida) - 1, Year(pFecValida)) Then
'              If CDate(pFecIngreso) > CDate(FecIngTmp) Then
'                    'mes = Abs(DateDiff("m", "01/" & COL_SEGUNDOMES + 1 & "/" & Txtano.Text - 1, "01/" & CmbMes.ListIndex + 1 & "/" & Txtano.Text)) + 1
'                    mes = Abs(DateDiff("m", pFecIngreso, pFecValida))
'                    If Day(pFecIngreso) = 1 Then
'                        Dia = 0
'                        mes = mes + 1
'                        pFecproc = DateAdd("d", -1, DateAdd("m", 1, pFecIngreso))
'                    Else
'                        pFecproc = DateAdd("d", -1, "01/" & Month(DateAdd("m", 1, pFecIngreso)) & "/" & Year(DateAdd("m", 1, pFecIngreso)))
'                        Dia = DateDiff("d", pFecIngreso, pFecproc)
'                        pFecproc = DateAdd("d", Val(Dia), pFecIngreso)
'                    End If
'                Else
'                    mes = Abs(DateDiff("m", "01/" & COL_SEGUNDOMES & "/" & Txtano.Text - 1, "01/" & Cmbmes.ListIndex + 1 & "/" & Txtano.Text)) + 1
'               End If
'
'                'mes = Abs(DateDiff("m", CDate(pFecIngreso), "01/" & CmbMes.ListIndex + 1 & "/" & Txtano.Text)) + 1
'            Else
'                mes = Abs(DateDiff("m", "01/" & COL_SEGUNDOMES & "/" & Txtano.Text, "01/" & Cmbmes.ListIndex + 1 & "/" & Txtano.Text)) + 1
'            End If
'
'        Else
'            año = 0
'            'Dia = Abs(DateDiff("d", DateAdd("m", mes, pFecIngreso), DateAdd("d", -1, DateAdd("m", 1, "01/" & CmbMes.ListIndex + 1 & "/" & txtano.Text))))
'
''            pFecproc = DateAdd("d", -1, "01/" & Month(DateAdd("m", 1, pFecIngreso)) & "/" & Year(DateAdd("m", 1, pFecIngreso)))
''            Dia = DateDiff("d", pFecIngreso, pFecproc)
''            pFecproc = DateAdd("d", Val(Dia), pFecIngreso)
''            mes = DateDiff("m", pFecproc, DateAdd("d", -1, DateAdd("m", 1, "01/" & Cmbmes.ListIndex + 1 & "/" & Txtano.Text)))
'
'            If Day(pFecIngreso) = 1 Then
'                Dia = 0
'                pFecproc = DateAdd("d", -1, DateAdd("m", 1, pFecIngreso))
'            Else
'                pFecproc = DateAdd("d", -1, "01/" & Month(DateAdd("m", 1, pFecIngreso)) & "/" & Year(DateAdd("m", 1, pFecIngreso)))
'                If Trim(Str(Year(pFecIngreso))) = Trim(Txtano.Text) And Month(pFecIngreso) < Cmbmes.ListIndex + 1 Then
'                    If Month(pFecIngreso) > Month(FecIngTmp) Then
'                        Dia = DateDiff("d", pFecIngreso, pFecproc)
'                    Else
'                        Dia = DateDiff("d", pFecValida, pFecValida)
'
'                    End If
'
'                    '
'                Else
'                    Dia = DateDiff("d", pFecIngreso, pFecproc)
'                End If
'                pFecproc = DateAdd("d", Val(Dia), pFecIngreso)
'            End If
'
'            If Trim(Str(Year(pFecIngreso))) = Trim(Txtano.Text) And Month(pFecIngreso) < Cmbmes.ListIndex + 1 Then
'                If Month(pFecIngreso) > Month(FecIngTmp) Then
'                    mes = DateDiff("m", "01/" & Cmbmes.ListIndex + 1 & "/" & Txtano.Text, DateAdd("d", -1, DateAdd("m", 1, "01/" & Cmbmes.ListIndex + 1 & "/" & Txtano.Text))) + 1
'                Else
'                    mes = Abs(DateDiff("m", "01/" & COL_SEGUNDOMES & "/" & Txtano.Text, "01/" & Cmbmes.ListIndex + 1 & "/" & Txtano.Text)) + 1
'                End If
'
'
'            Else
'                mes = DateDiff("m", pFecproc, DateAdd("d", -1, DateAdd("m", 1, "01/" & Cmbmes.ListIndex + 1 & "/" & Txtano.Text)))
'            End If
'
'          '  mes = DateDiff("m", pFecproc, DateAdd("d", -1, DateAdd("m", 1, "01/" & CmbMes.ListIndex + 1 & "/" & txtano.Text)))
'
'            If Day(pFecIngreso) = 1 Then mes = CStr(Val(mes) + 1)
'
'        End If
'    Else
'
'        '************codigo modificado giovanni 23082007******************************
'        'If pFecIngreso < CDate("01/" & COL_PRIMERMES & "/" & txtano.Text) Or pFecIngreso < CDate("01/" & COL_SEGUNDOMES & "/" & txtano.Text) Then
'        If pFecIngreso < CDate("01/" & COL_PRIMERMES & "/" & Txtano.Text) Then
'        '*****************************************************************************
'
'            mes = Abs(DateDiff("m", "01/" & COL_PRIMERMES & "/" & Txtano.Text, "01/" & Cmbmes.ListIndex + 1 & "/" & Txtano.Text)) + 1
'
'        Else
''            año = 0
''            mes = Abs(DateDiff("m", pFecIngreso, DateAdd("d", -1, DateAdd("m", 1, "01/" & CmbMes.ListIndex + 1 & "/" & txtano.Text))))
''            Dia = Abs(DateDiff("d", DateAdd("m", mes, pFecIngreso), DateAdd("d", -2, DateAdd("m", 1, "01/" & CmbMes.ListIndex + 1 & "/" & txtano.Text))))
'
'            año = 0
'            If Day(pFecIngreso) = 1 Then
'                Dia = 0
'                pFecproc = DateAdd("d", -1, DateAdd("m", 1, pFecIngreso))
'            Else
'                pFecproc = DateAdd("d", -1, "01/" & Month(DateAdd("m", 1, pFecIngreso)) & "/" & Year(DateAdd("m", 1, pFecIngreso)))
'                Dia = DateDiff("d", pFecIngreso, pFecproc)
'                pFecproc = DateAdd("d", Val(Dia), pFecIngreso)
'            End If
'
'            If Trim(Str(Year(pFecIngreso))) = Trim(Txtano.Text) And Month(pFecIngreso) < Cmbmes.ListIndex + 1 Then
'                mes = DateDiff("m", pFecproc, DateAdd("d", -1, DateAdd("m", 1, "01/" & Cmbmes.ListIndex + 1 & "/" & Txtano.Text)))
'            Else
'                mes = DateDiff("m", pFecproc, DateAdd("d", -1, DateAdd("m", 1, "01/" & Cmbmes.ListIndex + 1 & "/" & Txtano.Text)))
'            End If
'
'            If Day(pFecIngreso) = 1 Then mes = CStr(Val(mes) + 1)
'        End If
'
'    End If
'
'Salir:
''If mes > 6 Then mes = mes - 6
'CalcularMeses = Space(2 - Len(Trim(año))) & año & " " & Space(2 - Len(Trim(mes))) & mes & " " & Space(2 - Len(Trim(Dia))) & Dia & " "
'End Function

Private Function CalcularMeses(ByVal pFecIngreso As String) As String
Dim mesestmp As String
Dim año As String, Mes As String, Dia As String
Dim FecIngTmp As String
Dim pFecproc As String
Dim pFecValida As String


año = 0: Mes = 0: Dia = 0
If Year(CDate(pFecIngreso)) > Val(Txtano.Text) Then GoTo Salir

    If Cmbmes.ListIndex + 1 >= COL_SEGUNDOMES Or Cmbmes.ListIndex + 1 < COL_PRIMERMES Then
        pFecValida = DateAdd("m", 1, "01/" & Cmbmes.ListIndex + 1 & "/" & Txtano.Text)
        pFecValida = DateAdd("d", -1, "01/" & Month(pFecValida) & "/" & Year(pFecValida))
        FecIngTmp = DateAdd("d", 0, "01/" & COL_SEGUNDOMES & "/" & Year(pFecValida) - 1)
        
        If Year(CDate(pFecIngreso)) < Year(CDate(pFecValida)) Then
            If Month(CDate(pFecValida)) >= 1 And Month(CDate(pFecValida)) < COL_PRIMERMES Then
            'ESTE CAMBIO ES DE EMERGENCIA SE DEBERA DE MODIFICAR DESPUES
              ' If Day(CDate(pFecIngreso)) > 1 And Month(CDate(pFecIngreso)) >= COL_SEGUNDOMES And Year(CDate(pFecIngreso)) = IIf(Month(pFecValida) = 4, Year(pFecValida) - 1, Year(pFecValida)) Then
              If CDate(pFecIngreso) > CDate(FecIngTmp) Then
                    'mes = Abs(DateDiff("m", "01/" & COL_SEGUNDOMES + 1 & "/" & Txtano.Text - 1, "01/" & CmbMes.ListIndex + 1 & "/" & Txtano.Text)) + 1
                    Mes = Abs(DateDiff("m", pFecIngreso, pFecValida))
                    If Day(pFecIngreso) = 1 Then
                        Dia = 0
                        Mes = Mes + 1
                        pFecproc = DateAdd("d", -1, DateAdd("m", 1, pFecIngreso))
                    Else
                        pFecproc = DateAdd("d", -1, "01/" & Month(DateAdd("m", 1, pFecIngreso)) & "/" & Year(DateAdd("m", 1, pFecIngreso)))
                        Dia = DateDiff("d", pFecIngreso, pFecproc)
                        pFecproc = DateAdd("d", Val(Dia), pFecIngreso)
                    End If
                Else
                    Mes = Abs(DateDiff("m", "01/" & COL_SEGUNDOMES & "/" & Txtano.Text - 1, "01/" & Cmbmes.ListIndex + 1 & "/" & Txtano.Text)) + 1
               End If
                
                'mes = Abs(DateDiff("m", CDate(pFecIngreso), "01/" & CmbMes.ListIndex + 1 & "/" & Txtano.Text)) + 1
            Else
                Mes = Abs(DateDiff("m", "01/" & COL_SEGUNDOMES & "/" & Txtano.Text, "01/" & Cmbmes.ListIndex + 1 & "/" & Txtano.Text)) + 1
            End If
            
        Else
            año = 0
            'Dia = Abs(DateDiff("d", DateAdd("m", mes, pFecIngreso), DateAdd("d", -1, DateAdd("m", 1, "01/" & CmbMes.ListIndex + 1 & "/" & txtano.Text))))
            
'            pFecproc = DateAdd("d", -1, "01/" & Month(DateAdd("m", 1, pFecIngreso)) & "/" & Year(DateAdd("m", 1, pFecIngreso)))
'            Dia = DateDiff("d", pFecIngreso, pFecproc)
'            pFecproc = DateAdd("d", Val(Dia), pFecIngreso)
'            mes = DateDiff("m", pFecproc, DateAdd("d", -1, DateAdd("m", 1, "01/" & Cmbmes.ListIndex + 1 & "/" & Txtano.Text)))
            
            If Day(pFecIngreso) = 1 Then
                Dia = 0
                pFecproc = DateAdd("d", -1, DateAdd("m", 1, pFecIngreso))
            Else
                pFecproc = DateAdd("d", -1, "01/" & Month(DateAdd("m", 1, pFecIngreso)) & "/" & Year(DateAdd("m", 1, pFecIngreso)))
                If Trim(Str(Year(pFecIngreso))) = Trim(Txtano.Text) And Month(pFecIngreso) < Cmbmes.ListIndex + 1 Then
                    If Month(pFecIngreso) > Month(FecIngTmp) Then
                        Dia = DateDiff("d", pFecIngreso, pFecproc)
                    Else
                    'And CDate(pFecIngreso) > CDate(pFecValida)
                    If CDate(pFecIngreso) > CDate(FecIngTmp) Then
                        Dia = DateDiff("d", pFecIngreso, pFecproc)
                    Else
                        Dia = DateDiff("d", pFecValida, pFecValida)
                    End If
                    End If
                    
                    '
                Else
                    Dia = DateDiff("d", pFecIngreso, pFecproc)
                End If
                pFecproc = DateAdd("d", Val(Dia), pFecIngreso)
            End If
            
            If Trim(Str(Year(pFecIngreso))) = Trim(Txtano.Text) And Month(pFecIngreso) < Cmbmes.ListIndex + 1 Then
                If Month(pFecIngreso) > Month(FecIngTmp) Then
                    Mes = DateDiff("m", "01/" & Cmbmes.ListIndex + 1 & "/" & Txtano.Text, DateAdd("d", -1, DateAdd("m", 1, "01/" & Cmbmes.ListIndex + 1 & "/" & Txtano.Text))) + 1
                Else
                'And CDate(pFecIngreso) > CDate(pFecValida) Then
                If CDate(pFecIngreso) > CDate(FecIngTmp) Then
                    'mes = DateDiff("m", "01/" & CmbMes.ListIndex + 1 & "/" & Txtano.Text, DateAdd("d", -1, DateAdd("m", 1, "01/" & CmbMes.ListIndex + 1 & "/" & Txtano.Text))) + 1
                    Mes = DateDiff("m", pFecIngreso, DateAdd("d", -1, DateAdd("m", 1, "01/" & Cmbmes.ListIndex + 1 & "/" & Txtano.Text)))
                Else
                    Mes = Abs(DateDiff("m", "01/" & COL_SEGUNDOMES & "/" & Txtano.Text, "01/" & Cmbmes.ListIndex + 1 & "/" & Txtano.Text)) + 1
                End If
                End If
                

                
            Else
                Mes = DateDiff("m", pFecproc, DateAdd("d", -1, DateAdd("m", 1, "01/" & Cmbmes.ListIndex + 1 & "/" & Txtano.Text)))
            End If
            
          '  mes = DateDiff("m", pFecproc, DateAdd("d", -1, DateAdd("m", 1, "01/" & CmbMes.ListIndex + 1 & "/" & txtano.Text)))
            
            'If Day(pFecIngreso) = 1 Then mes = CStr(Val(mes) + 1)
            
        End If
    Else

        '************codigo modificado giovanni 23082007******************************
        'If pFecIngreso < CDate("01/" & COL_PRIMERMES & "/" & txtano.Text) Or pFecIngreso < CDate("01/" & COL_SEGUNDOMES & "/" & txtano.Text) Then
        If pFecIngreso < CDate("01/" & COL_PRIMERMES & "/" & Txtano.Text) Then
        '*****************************************************************************

            Mes = Abs(DateDiff("m", "01/" & COL_PRIMERMES & "/" & Txtano.Text, "01/" & Cmbmes.ListIndex + 1 & "/" & Txtano.Text)) + 1
                 
        Else
'            año = 0
'            mes = Abs(DateDiff("m", pFecIngreso, DateAdd("d", -1, DateAdd("m", 1, "01/" & CmbMes.ListIndex + 1 & "/" & txtano.Text))))
'            Dia = Abs(DateDiff("d", DateAdd("m", mes, pFecIngreso), DateAdd("d", -2, DateAdd("m", 1, "01/" & CmbMes.ListIndex + 1 & "/" & txtano.Text))))
            
            año = 0
            If Day(pFecIngreso) = 1 Then
                Dia = 0
                pFecproc = DateAdd("d", -1, DateAdd("m", 1, pFecIngreso))
            Else
                pFecproc = DateAdd("d", -1, "01/" & Month(DateAdd("m", 1, pFecIngreso)) & "/" & Year(DateAdd("m", 1, pFecIngreso)))
                Dia = DateDiff("d", pFecIngreso, pFecproc)
                pFecproc = DateAdd("d", Val(Dia), pFecIngreso)
            End If
            
            If Trim(Str(Year(pFecIngreso))) = Trim(Txtano.Text) And Month(pFecIngreso) < Cmbmes.ListIndex + 1 Then
                Mes = DateDiff("m", pFecproc, DateAdd("d", -1, DateAdd("m", 1, "01/" & Cmbmes.ListIndex + 1 & "/" & Txtano.Text)))
            Else
                Mes = DateDiff("m", pFecproc, DateAdd("d", -1, DateAdd("m", 1, "01/" & Cmbmes.ListIndex + 1 & "/" & Txtano.Text)))
            End If
            
            If Day(pFecIngreso) = 1 Then Mes = CStr(Val(Mes) + 1)
        End If
                
    End If
    
Salir:
'If mes > 6 Then mes = mes - 6
CalcularMeses = Space(2 - Len(Trim(año))) & año & " " & Space(2 - Len(Trim(Mes))) & Mes & " " & Space(2 - Len(Trim(Dia))) & Dia & " "
End Function

Private Function MontoIndecnizado(ByVal pValor As String, ByVal pImporte As String, ByVal pTipoTrab As String, montofalta As Double, montosubs As Double) As String
Dim montotmp As Currency

If pTipoTrab = "02" Then

  'repl totper2 with roun((manos*30*mjornal),2)+roun((meses*30*mjornal/12),2)+roun((mdias*30*mjornal/365),2)
    
    If wGrupoPla <> "01" Then
        montotmp = Val(Mid(pValor, 4, 2)) * ((pImporte * DIAS_TRABAJO) / 12)
        montotmp = Round(montotmp + Val(Mid(pValor, 7, 2)) * (pImporte / 12), 2)
    Else
        montotmp = Val(Mid(pValor, 4, 2)) * ((pImporte * DIAS_TRABAJO) / 12)
        montotmp = Round(montotmp + (Val(Mid(pValor, 7, 2)) * DIAS_TRABAJO * pImporte / 365), 2)
    End If
    
Else
'     Repl totper3 with roun((manos*mbasico),2)+roun(meses*mbasico/12,2)+roun(mdias*mbasico/365,2)
    
    If wGrupoPla <> "01" Then
        montotmp = Val(Mid(pValor, 4, 2)) * (pImporte / 12)
        montotmp = Round(montotmp + Val(Mid(pValor, 7, 2)) * ((pImporte / 12) / DIAS_TRABAJO), 2)
    Else
        montotmp = Val(Mid(pValor, 4, 2)) * (pImporte / 12)
        montotmp = Round(montotmp + (Val(Mid(pValor, 7, 2)) * pImporte / 365), 2)
    End If
End If

MontoIndecnizado = montotmp - montofalta - montosubs

End Function

Private Function BuscaColumna(ByVal pCampo As String, ByVal pMaxcol As Integer) As Integer
Dim iRow As Integer
BuscaColumna = 0
For iRow = 3 To pMaxcol
    If ArrReporte(iRow, 0) = pCampo Then
        BuscaColumna = iRow
        Exit Function
    End If
Next
End Function
Private Sub ReporteCts(Excluir As String)

Dim rs As Object
'Dim xlApp2 As Excel.Application
'Dim xlApp1  As Excel.Application
'Dim xlBook As Excel.Workbook
'Dim xlSheet As Excel.Worksheet


Dim xlApp2 As Object
Dim xlApp1  As Object
Dim xlBook As Object
Dim xlSheet As Object


Dim I As Integer
Dim Fila As Integer
Dim Columna As Integer
Dim x As Integer
Dim lNumTrab As Integer
lNumTrab = 0

Dim MArea  As String

Set xlApp1 = CreateObject("Excel.Application")
xlApp1.Workbooks.Add
Set xlApp2 = xlApp1.Application
Set xlBook = xlApp2.Workbooks(1)
Set xlSheet = xlApp2.Worksheets("HOJA1")

xlSheet.Range("A:A").ColumnWidth = 7
xlSheet.Range("B:B").ColumnWidth = 40
xlSheet.Range("C:C").ColumnWidth = 9.71
xlSheet.Range("D:Z").ColumnWidth = 14
xlSheet.Range("D:Z").NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "

xlSheet.Cells(1, 1).Value = Trim(Cmbcia.Text)
xlSheet.Cells(1, 1).Font.Bold = True
xlSheet.Cells(2, 2).Value = " REPORTE DE PROVISION POR TIEMPO DE SERVICIO " & Cmbtipo.Text & " - " & Cmbmes.Text & " " & Txtano.Text
xlSheet.Cells(2, 2).Font.Bold = True
xlSheet.Cells(2, 2).Font.Size = 12
xlSheet.Cells(2, 2).HorizontalAlignment = xlCenter

Fila = 4
xlSheet.Cells(Fila, 1).Value = "Codigo"
xlSheet.Cells(Fila, 2).Value = "Nombre Trabajador"
xlSheet.Cells(Fila, 3).Value = "DNI"
xlSheet.Cells(Fila, 4).Value = "F.Ing"

xlSheet.Cells(Fila + 2, 5).Value = "Dias"
xlSheet.Cells(Fila + 3, 5).Value = "Falta"
xlSheet.Cells(Fila + 2, 6).Value = "Dias"
xlSheet.Cells(Fila + 3, 6).Value = "Subs"

xlSheet.Cells(Fila, 7).Value = "F.Cese"
xlSheet.Cells(Fila, 8).Value = "Tiempo Servicio"

xlSheet.Cells(Fila + 2, 4).Value = "Tiempo"
xlSheet.Cells(Fila + 3, 4).Value = "T.Serv"
xlSheet.Cells(Fila + 2, 7).Value = "Monto Inden."
xlSheet.Cells(Fila + 3, 7).Value = "Periodo 1"
xlSheet.Cells(Fila + 2, 8).Value = "Monto Inden."
xlSheet.Cells(Fila + 3, 8).Value = "Periodo 2"
xlSheet.Cells(Fila + 2, 9).Value = "Monto Inden."
xlSheet.Cells(Fila + 3, 9).Value = "Sin Topes"
xlSheet.Cells(Fila + 2, 10).Value = "Total"
xlSheet.Cells(Fila + 3, 10).Value = "Faltas"
xlSheet.Cells(Fila + 2, 11).Value = "Total"
xlSheet.Cells(Fila + 3, 11).Value = "Subsid"

xlSheet.Cells(Fila + 2, 12).Value = "Monto Total"
xlSheet.Cells(Fila + 3, 12).Value = "Indenmizatorio"
xlSheet.Cells(Fila + 2, 13).Value = "Provisionado"
xlSheet.Cells(Fila + 3, 13).Value = "Año Anterior"
xlSheet.Cells(Fila + 2, 14).Value = "Ajuste de"
xlSheet.Cells(Fila + 3, 14).Value = "Provision"
xlSheet.Cells(Fila + 2, 15).Value = "Provisionado"
xlSheet.Cells(Fila + 3, 15).Value = "Este Año"
xlSheet.Cells(Fila + 2, 16).Value = "Provision"
xlSheet.Cells(Fila + 3, 16).Value = "Mes Actual"
xlSheet.Cells(Fila + 2, 17).Value = "Saldo Pend."
xlSheet.Cells(Fila + 3, 17).Value = "De Provision"

Captura_Tipo_Trabajador

Dim lTipoTrab As String
lTipoTrab = s_TipoTrabajador_Cts
If lTipoTrab = "03" Then lTipoTrab = "**"
Sql = "usp_Carga_Reporte_CTS '" & wcia & "'," & Txtano.Text & "," & Cmbmes.ListIndex + 1 & ",'" & lTipoTrab & "','" & Excluir & "'"
If (fAbrRst(rs, Sql)) Then rs.MoveFirst

MArea = ""
Dim lCol As Integer
Dim rsTit As ADODB.Recordset
lCol = 1
If rs.RecordCount > 0 Then lCol = rs!nCol
For I = 1 To lCol
   Sql = "select des_cts from Pla_Cts_Titulos where cia='" & wcia & "' and codinterno='" & Mid(rs.Fields(I + 5).Name, 2, 2) & "'"
    If (fAbrRst(rsTit, Sql)) Then
      xlSheet.Cells(Fila, I + 8).Value = Trim(rsTit(0) & "")
   Else
      xlSheet.Cells(Fila, I + 8).Value = UCase(rs.Fields(I + 5).Name)
   End If
   xlSheet.Range(xlSheet.Cells(Fila, I + 8), xlSheet.Cells(Fila + 1, I + 8)).Merge
   xlSheet.Range(xlSheet.Cells(Fila, I + 8), xlSheet.Cells(Fila + 1, I + 8)).WrapText = True
   xlSheet.Range(xlSheet.Cells(Fila, I + 8), xlSheet.Cells(Fila + 1, I + 8)).VerticalAlignment = xlTop
   xlSheet.Range(xlSheet.Cells(Fila, I + 8), xlSheet.Cells(Fila + 1, I + 8)).HorizontalAlignment = xlCenter
   rsTit.Close
Next
xlSheet.Cells(Fila, I + 8).Value = "Jornal"
xlSheet.Cells(Fila + 1, I + 8).Value = "Indennizat"
xlSheet.Range(xlSheet.Cells(Fila, I + 8), xlSheet.Cells(Fila + 1, I + 8)).HorizontalAlignment = xlCenter

Dim sLin As Integer
sLin = I + 8
If sLin < 17 Then sLin = 17

xlSheet.Range(xlSheet.Cells(4, 1), xlSheet.Cells(4, sLin)).Borders(xlEdgeTop).LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(7, 1), xlSheet.Cells(7, sLin)).Borders(xlEdgeBottom).LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(6, 1), xlSheet.Cells(7, I + 8)).HorizontalAlignment = xlCenter

xlSheet.Range(xlSheet.Cells(2, 1), xlSheet.Cells(2, I + 8)).Merge

Fila = 10
Do While Not rs.EOF
   If MArea <> Trim(rs!ccosto) Then
      If MArea <> "" Then
         rs.MovePrevious
         xlSheet.Cells(Fila, 12).Value = rs!monto_total
         xlSheet.Cells(Fila, 12).Borders(xlEdgeTop).LineStyle = xlContinuous
         xlSheet.Cells(Fila, 16).Value = rs!provision_actual
         xlSheet.Cells(Fila, 16).Borders(xlEdgeTop).LineStyle = xlContinuous
         rs.MoveNext
         Fila = Fila + 2
      End If
      If Trim(rs(2)) <> "ZZZZZ" Then
         xlSheet.Cells(Fila, 1).Value = Trim(rs(2))
         xlSheet.Cells(Fila, 1).Font.Bold = True
         MArea = Trim(rs!ccosto)
      End If
      Fila = Fila + 2
   End If
   If Trim(rs!PlaCod) <> "ZZZZZ" Then
      xlSheet.Cells(Fila, 1).Value = Trim(rs!PlaCod)
      xlSheet.Cells(Fila, 2).Value = Trim(rs!nombre)
      xlSheet.Cells(Fila, 3).Value = Trim(rs!nro_doc)
      xlSheet.Cells(Fila, 4).Value = rs!fIngreso
      xlSheet.Cells(Fila, 5).Value = rs!fcese
      For I = 1 To lCol
         xlSheet.Cells(Fila, I + 8).Value = rs(I + 5)
      Next
      xlSheet.Cells(Fila, I + 8).Value = rs!totalremun
      lNumTrab = lNumTrab + 1
      Fila = Fila + 1
      xlSheet.Cells(Fila, 4).Value = Trim(rs!recordact & "")
      xlSheet.Cells(Fila, 5).Value = rs!diasctsfalto
      xlSheet.Cells(Fila, 6).Value = rs!dias_subs
      xlSheet.Cells(Fila, 7).Value = rs!monto_inde1
      xlSheet.Cells(Fila, 8).Value = rs!monto_inde2
      xlSheet.Cells(Fila, 9).Value = rs!monto_indetope
      xlSheet.Cells(Fila, 10).Value = rs!Monto_Faltas
      xlSheet.Cells(Fila, 11).Value = rs!Monto_Subs
      xlSheet.Cells(Fila, 12).Value = rs!monto_total
      xlSheet.Cells(Fila, 13).Value = rs!prov_anoante
      xlSheet.Cells(Fila, 14).Value = rs!ajuste_prov
      xlSheet.Cells(Fila, 15).Value = rs!provision_ano
      xlSheet.Cells(Fila, 16).Value = rs!provision_actual
      xlSheet.Cells(Fila, 17).Value = rs!saldo_prov
      Fila = Fila + 2
   ElseIf Trim(rs!ccosto) = "ZZZZZ" Then
      For I = 1 To lCol
         xlSheet.Cells(Fila, I + 8).Value = rs(I + 5)
      Next
      xlSheet.Cells(Fila, I + 8).Value = rs!totalremun
   
      Fila = Fila + 1
      xlSheet.Cells(Fila, 5).Value = rs!diasctsfalto
      xlSheet.Cells(Fila, 6).Value = rs!dias_subs
      xlSheet.Cells(Fila, 7).Value = rs!monto_inde1
      xlSheet.Cells(Fila, 8).Value = rs!monto_inde2
      xlSheet.Cells(Fila, 9).Value = rs!monto_indetope
      xlSheet.Cells(Fila, 10).Value = rs!Monto_Faltas
      xlSheet.Cells(Fila, 11).Value = rs!Monto_Subs
      xlSheet.Cells(Fila, 12).Value = rs!monto_total
      xlSheet.Cells(Fila, 13).Value = rs!prov_anoante
      xlSheet.Cells(Fila, 14).Value = rs!ajuste_prov
      xlSheet.Cells(Fila, 15).Value = rs!provision_ano
      xlSheet.Cells(Fila, 16).Value = rs!provision_actual
      xlSheet.Cells(Fila, 17).Value = rs!saldo_prov
      xlSheet.Range(xlSheet.Cells(Fila - 1, 5), xlSheet.Cells(Fila - 1, sLin)).Borders(xlEdgeTop).LineStyle = xlContinuous
      xlSheet.Range(xlSheet.Cells(Fila, 5), xlSheet.Cells(Fila, sLin)).Borders(xlEdgeBottom).LineStyle = xlContinuous
   End If
   rs.MoveNext
Loop
Fila = Fila + 2
xlSheet.Cells(Fila, 2).Value = "NUMERO DE TRABAJADORES"
xlSheet.Cells(Fila, 3).Value = lNumTrab

xlApp2.Application.ActiveWindow.DisplayGridLines = False
xlSheet.Range("A1:A1").Select
xlApp2.Application.Caption = "PROVISION DE CTS"
xlApp2.ActiveWindow.Zoom = 80
xlApp2.Application.Visible = True

End Sub
Private Sub ReporteCtsRODA()
Dim rs As Object
'Dim xlApp2 As Excel.Application
'Dim xlApp1  As Excel.Application
'Dim xlBook As Excel.Workbook
'Dim xlSheet As Excel.Worksheet

Dim xlApp2 As Object
Dim xlApp1  As Object
Dim xlBook As Object
Dim xlSheet As Object



Dim I As Long
Dim Fila As Integer
Dim Columna As Integer
Dim ArrTotales() As Variant

ReDim Preserve ArrTotales(0 To 29)

Set xlApp1 = CreateObject("Excel.Application")
xlApp1.Workbooks.Add
Set xlApp2 = xlApp1.Application
Set xlBook = xlApp2.Workbooks(1)
Set xlSheet = xlApp2.Worksheets("HOJA1")

xlSheet.Range("A:A").ColumnWidth = 7
xlSheet.Range("B:B").ColumnWidth = 23.29
xlSheet.Range("C:C").ColumnWidth = 9.71
xlSheet.Range("D:D").ColumnWidth = 9.71
xlSheet.Range("E:E").ColumnWidth = 9.71
xlSheet.Range("F:U").ColumnWidth = 10.3
xlSheet.Range("C:U").HorizontalAlignment = xlCenter

xlSheet.Cells(1, 1).Value = Trim(Cmbcia.Text)
xlSheet.Cells(1, 1).Font.Bold = True
xlSheet.Cells(2, 2).Value = " REPORTE DE PROVISION POR TIEMPO DE SERVICIO "
xlSheet.Cells(2, 2).Font.Bold = True
xlSheet.Cells(2, 2).Font.Size = 12
xlSheet.Cells(2, 2).HorizontalAlignment = xlCenter
xlSheet.Range("B2:U2").Merge

xlSheet.Range("A6:U6").Merge
xlSheet.Range("A6:U6") = "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
Fila = 7
xlSheet.Cells(Fila, 1).Value = "Codigo"
xlSheet.Cells(Fila, 2).Value = "Nombre Trabajador"
xlSheet.Cells(Fila, 3).Value = "F.Ing"

xlSheet.Cells(Fila, 4).Value = "Dias"
xlSheet.Cells(Fila + 1, 4).Value = "CtsFalto"

xlSheet.Cells(Fila, 5).Value = "Total"
xlSheet.Cells(Fila + 1, 5).Value = "Dias"

xlSheet.Cells(Fila, 6).Value = "F.Cese"
xlSheet.Cells(Fila, 7).Value = "Tiempo Servicio"
xlSheet.Cells(Fila, 8).Value = "Jornal"
xlSheet.Cells(Fila + 1, 8).Value = "Basico"
xlSheet.Cells(Fila, 9).Value = "Bonif"
xlSheet.Cells(Fila + 1, 9).Value = "Afp 3%"
xlSheet.Cells(Fila, 10).Value = "Bonif"
xlSheet.Cells(Fila + 1, 10).Value = "Costo Vida"
xlSheet.Cells(Fila, 11).Value = "Bonif"
xlSheet.Cells(Fila + 1, 11).Value = "T.Servicio"
xlSheet.Cells(Fila, 12).Value = "Asignacion"
xlSheet.Cells(Fila + 1, 12).Value = "Familiar"
xlSheet.Cells(Fila, 13).Value = "Promedio"
xlSheet.Cells(Fila + 1, 13).Value = "Gratific"
xlSheet.Cells(Fila, 14).Value = "Promedio "
xlSheet.Cells(Fila + 1, 14).Value = "H.Extras"
xlSheet.Cells(Fila, 15).Value = "Promedio"
xlSheet.Cells(Fila + 1, 15).Value = "Otros Pagos"
xlSheet.Cells(Fila, 16).Value = "Promedio"
xlSheet.Cells(Fila + 1, 16).Value = "H. Verano"
xlSheet.Cells(Fila, 17).Value = "Promedio"
xlSheet.Cells(Fila + 1, 17).Value = "P.xTurno"
xlSheet.Cells(Fila, 18).Value = "Promedio"
xlSheet.Cells(Fila + 1, 18).Value = "P.xProducc"
xlSheet.Cells(Fila, 19).Value = "Promedio"
xlSheet.Cells(Fila + 1, 19).Value = "Bo.Prod."
xlSheet.Cells(Fila, 20).Value = "Promedio"
xlSheet.Cells(Fila + 1, 20).Value = "Rendim."
xlSheet.Cells(Fila, 21).Value = "Jornal"
xlSheet.Cells(Fila + 1, 21).Value = "Indennizat"
Fila = Fila + 2
xlSheet.Cells(Fila, 3).Value = "Tiempo"
xlSheet.Cells(Fila + 1, 3).Value = "T.Serv"
xlSheet.Cells(Fila, 6).Value = "Monto Inden."
xlSheet.Cells(Fila + 1, 6).Value = "Periodo 1"
xlSheet.Cells(Fila, 7).Value = "Monto Inden."
xlSheet.Cells(Fila + 1, 7).Value = "Periodo 2"
xlSheet.Cells(Fila, 8).Value = "Monto Inden"
xlSheet.Cells(Fila + 1, 8).Value = "Sin Topes"
xlSheet.Cells(Fila, 9).Value = "Monto Total"
xlSheet.Cells(Fila + 1, 9).Value = "Indennizatorio"
xlSheet.Cells(Fila, 10).Value = "Provisionado"
xlSheet.Cells(Fila + 1, 10).Value = "Año Anterior"
xlSheet.Cells(Fila, 11).Value = "Ajuste de"
xlSheet.Cells(Fila + 1, 11).Value = "Provision"
xlSheet.Cells(Fila, 12).Value = "Provisionado"
xlSheet.Cells(Fila + 1, 12).Value = "Este Año"
xlSheet.Cells(Fila, 13).Value = "Provision "
xlSheet.Cells(Fila + 1, 13).Value = "Mes Actual"
xlSheet.Cells(Fila, 14).Value = "Saldo Pend."
xlSheet.Cells(Fila + 1, 14).Value = "De Provision"
Fila = Fila + 2
xlSheet.Range("A" & Fila & ":U" & Fila).Merge
xlSheet.Range("A" & Fila & ":U" & Fila) = "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
Fila = Fila + 4

Captura_Tipo_Trabajador

' LFSA - 10/05/2012 - > El concepto i24 e i25 incluido en el grupo de extras
Select Case s_TipoTrabajador_Cts
    Case "01"
        Sql = "SELECT a.placod,ltrim(rtrim(b.ap_pat))+' '+ltrim(rtrim(b.ap_mat))," & _
        "CONVERT(VARCHAR(10),b.fingreso,103),'','',CONVERT(VARCHAR(10),b.fcese,103),'',i01,i06,i07,i04,i02,p15,p10+p11+p21+p24+p25,p16,p29,0, 0,p26,0,totalremun,a.recordact,a.diasctsfalto,a.recordiasfalto,monto_inde1,monto_inde2,monto_indetope," & _
        "monto_total,prov_anoante,ajuste_prov,provision_ano,provision_actual,saldo_prov FROM PLAPROVCTS a LEFT OUTER JOIN PLANILLAS b on (b.cia=a.cia and b.placod=a.placod and b.status!='*') WHERE " & _
        "a.cia='" & wcia & "' and year(fechaproceso)=" & Txtano.Text & " and a.status!='*' and month(fechaproceso)=" & Cmbmes.ListIndex + 1 & " and b.tipotrabajador='01' ORDER BY A.PLACOD"
    Case "02"
        Sql = "SELECT a.placod,ltrim(rtrim(b.ap_pat))+' '+ltrim(rtrim(b.ap_mat))," & _
        "CONVERT(VARCHAR(10),b.fingreso,103),'','',CONVERT(VARCHAR(10),b.fcese,103),'',i01,i06,i07,i04,i02,p15,p10+p11+p21+p24+p25,p16,p29,0, 0,p38,0,totalremun,a.recordact,a.diasctsfalto,a.recordiasfalto,monto_inde1,monto_inde2,monto_indetope," & _
        "monto_total,prov_anoante,ajuste_prov,provision_ano,provision_actual,saldo_prov FROM PLAPROVCTS a LEFT OUTER JOIN PLANILLAS b on (b.cia=a.cia and b.placod=a.placod and b.status!='*') WHERE " & _
        "a.cia='" & wcia & "' and year(fechaproceso)=" & Txtano.Text & " and a.status!='*' and month(fechaproceso)=" & Cmbmes.ListIndex + 1 & " and b.tipotrabajador='02' ORDER BY A.PLACOD"
    Case "03"
        Sql = "SELECT a.placod,ltrim(rtrim(b.ap_pat))+' '+ltrim(rtrim(b.ap_mat))," & _
        "CONVERT(VARCHAR(10),b.fingreso,103),'','',CONVERT(VARCHAR(10),b.fcese,103),'',i01,i06,i07,i04,i02,p15,p10+p11+p21+p24+p25,p16,p29,0, 0,p26,0,totalremun,a.recordact,a.diasctsfalto,a.recordiasfalto,monto_inde1,monto_inde2,monto_indetope," & _
        "monto_total,prov_anoante,ajuste_prov,provision_ano,provision_actual,saldo_prov FROM PLAPROVCTS a LEFT OUTER JOIN PLANILLAS b on (b.cia=a.cia and b.placod=a.placod and b.status!='*') WHERE " & _
        "a.cia='" & wcia & "' and year(fechaproceso)=" & Txtano.Text & " and a.status!='*' and month(fechaproceso)=" & Cmbmes.ListIndex + 1 & " ORDER BY A.PLACOD"
End Select
    

'sql = "SELECT a.placod,ltrim(rtrim(b.ap_pat))+' '+ltrim(rtrim(b.ap_mat))," & _
'"CONVERT(VARCHAR(10),b.fingreso,103),CONVERT(VARCHAR(10),b.fcese,103),'',i01,i06,i07,i04,i02,p15,p10+p11+p21,p16,0,0, p24+p25,p26,0,totalremun,a.recordact,monto_inde1,monto_inde2,monto_indetope," & _
'"monto_total,prov_anoante,ajuste_prov,provision_ano,provision_actual,saldo_prov FROM PLAPROVCTS a LEFT OUTER JOIN PLANILLAS b on (b.cia=a.cia and b.placod=a.placod and b.status!='*') WHERE " & _
'"a.cia='" & wcia & "' and year(fechaproceso)=" & Txtano.Text & " and a.status!='*' and month(fechaproceso)=" & Cmbmes.ListIndex + 1 & " and b.tipotrabajador='01' ORDER BY A.PLACOD"

If (fAbrRst(rs, Sql)) Then
    Do While Not rs.EOF
        
        Fila = Fila + 1
        For I = 0 To rs.Fields.count - 1
            If I = 21 Then
                Columna = 2
                Fila = Fila + 1
            End If
            Columna = Columna + 1
            If I > 6 Then
                If I = 21 Or I = 22 Or I = 23 Then
                    xlSheet.Cells(Fila, Columna).Value = Trim(rs(I))
                Else
                    xlSheet.Cells(Fila, Columna).Value = Format(Trim(rs(I)), "###,###,##0.00")
                    xlSheet.Cells(Fila, Columna).NumberFormat = "#,###,##0.00"
                End If
                'xlSheet.Cells(FILA, COLUMNA).Value = Format(Trim(rs(I)), "###,###,##0.00")
                'xlSheet.Cells(FILA, COLUMNA).Value = Trim(rs(I))
                'xlSheet.Cells(FILA, COLUMNA).NumberFormat = "#,###,##0.00"
                If (I <> 21 And I <> 22 And I <> 23) Then ArrTotales(I - 5) = ArrTotales(I - 5) + Val(rs(I)) Else ArrTotales(I - 5) = ""
            Else
                If I = 2 Then
                    xlSheet.Cells(Fila, Columna).Value = Format(Trim(rs(I)), "MM/DD/YYYY")
                Else
                    xlSheet.Cells(Fila, Columna).Value = Trim(rs(I))
                End If
                'xlSheet.Cells(FILA, COLUMNA).Value = Trim(rs(I))
                
            End If
        Next
        Columna = 0
        
        rs.MoveNext
    Loop
End If

Columna = 5
Fila = Fila + 3
For I = 0 To UBound(ArrTotales)
    If I = 16 Then Columna = 2: Fila = Fila + 1
    Columna = Columna + 1
    If I <> 16 Then
        xlSheet.Cells(Fila, Columna).Value = Format(Trim(ArrTotales(I)), "###,###,##0.00")
        xlSheet.Cells(Fila, Columna).NumberFormat = "#,###,##0.00"
    Else
        xlSheet.Cells(Fila, Columna).Value = ArrTotales(I)
    End If
Next I
'Call AddSheet(xlApp2)
xlApp2.Application.ActiveWindow.DisplayGridLines = False
xlSheet.Range("A1:A1").Select
xlApp2.Application.Caption = "PROVISION DE CTS"
xlApp2.ActiveWindow.Zoom = 80
xlApp2.Application.Visible = True


'Dim rs As Object
'Dim xlApp2 As Excel.Application
'Dim xlApp1  As Excel.Application
'Dim xlBook As Excel.Workbook
'Dim xlSheet As Excel.Worksheet
'Dim i As Long
'Dim Fila As Integer
'Dim Columna As Integer
'Dim ArrTotales() As Variant
'
'ReDim Preserve ArrTotales(0 To 23)
'
'Set xlApp1 = CreateObject("Excel.Application")
'xlApp1.Workbooks.Add
'Set xlApp2 = xlApp1.Application
'Set xlBook = xlApp2.Workbooks(1)
'Set xlSheet = xlApp2.Worksheets("HOJA1")
'
'xlSheet.Range("A:A").ColumnWidth = 7
'xlSheet.Range("B:B").ColumnWidth = 23.29
'xlSheet.Range("C:C").ColumnWidth = 9.71
'xlSheet.Range("D:S").ColumnWidth = 10.3
'xlSheet.Range("C:S").HorizontalAlignment = xlCenter
'
'xlSheet.Cells(1, 1).Value = Trim(Cmbcia.Text)
'xlSheet.Cells(1, 1).Font.Bold = True
'xlSheet.Cells(2, 2).Value = " REPORTE DE PROVISION POR TIEMPO DE SERVICIO "
'xlSheet.Cells(2, 2).Font.Bold = True
'xlSheet.Cells(2, 2).Font.Size = 12
'xlSheet.Cells(2, 2).HorizontalAlignment = xlCenter
'xlSheet.Range("B2:S2").Merge
'
'xlSheet.Range("A6:S6").Merge
'xlSheet.Range("A6:S6") = "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
'Fila = 7
'xlSheet.Cells(Fila, 1).Value = "Codigo"
'xlSheet.Cells(Fila, 2).Value = "Nombre Trabajador"
'xlSheet.Cells(Fila, 3).Value = "F.Ing"
'xlSheet.Cells(Fila, 4).Value = "F.Cese"
'xlSheet.Cells(Fila, 5).Value = "Tiempo Servicio"
'xlSheet.Cells(Fila, 6).Value = "Jornal"
'xlSheet.Cells(Fila + 1, 6).Value = "Basico"
'xlSheet.Cells(Fila, 7).Value = "Bonif"
'xlSheet.Cells(Fila + 1, 7).Value = "Afp 3%"
'xlSheet.Cells(Fila, 8).Value = "Bonif"
'xlSheet.Cells(Fila + 1, 8).Value = "Costo Vida"
'xlSheet.Cells(Fila, 9).Value = "Bonif"
'xlSheet.Cells(Fila + 1, 9).Value = "T.Servicio"
'xlSheet.Cells(Fila, 10).Value = "Asignacion"
'xlSheet.Cells(Fila + 1, 10).Value = "Familiar"
'xlSheet.Cells(Fila, 11).Value = "Promedio"
'xlSheet.Cells(Fila + 1, 11).Value = "Gratific"
'xlSheet.Cells(Fila, 12).Value = "Promedio "
'xlSheet.Cells(Fila + 1, 12).Value = "H.Extras"
'xlSheet.Cells(Fila, 13).Value = "Promedio"
'xlSheet.Cells(Fila + 1, 13).Value = "Otros Pagos"
'xlSheet.Cells(Fila, 14).Value = "Promedio"
'xlSheet.Cells(Fila + 1, 14).Value = "H. Verano"
'xlSheet.Cells(Fila, 15).Value = "Promedio"
'xlSheet.Cells(Fila + 1, 15).Value = "P.xTurno"
'xlSheet.Cells(Fila, 16).Value = "Promedio"
'xlSheet.Cells(Fila + 1, 16).Value = "P.xProducc"
'xlSheet.Cells(Fila, 17).Value = "Promedio"
'xlSheet.Cells(Fila + 1, 17).Value = "Bo.Prod."
'xlSheet.Cells(Fila, 18).Value = "Promedio"
'xlSheet.Cells(Fila + 1, 18).Value = "Rendim."
'xlSheet.Cells(Fila, 19).Value = "Jornal"
'xlSheet.Cells(Fila + 1, 19).Value = "Indennizat"
'Fila = Fila + 2
'xlSheet.Cells(Fila, 3).Value = "Tiempo"
'xlSheet.Cells(Fila + 1, 3).Value = "T.Serv"
'xlSheet.Cells(Fila, 4).Value = "Monto Inden."
'xlSheet.Cells(Fila + 1, 4).Value = "Periodo 1"
'xlSheet.Cells(Fila, 5).Value = "Monto Inden."
'xlSheet.Cells(Fila + 1, 5).Value = "Periodo 2"
'xlSheet.Cells(Fila, 6).Value = "Monto Inden"
'xlSheet.Cells(Fila + 1, 6).Value = "Sin Topes"
'xlSheet.Cells(Fila, 7).Value = "Monto Total"
'xlSheet.Cells(Fila + 1, 7).Value = "Indennizatorio"
'xlSheet.Cells(Fila, 8).Value = "Provisionado"
'xlSheet.Cells(Fila + 1, 8).Value = "Año Anterior"
'xlSheet.Cells(Fila, 9).Value = "Ajuste de"
'xlSheet.Cells(Fila + 1, 9).Value = "Provision"
'xlSheet.Cells(Fila, 10).Value = "Provisionado"
'xlSheet.Cells(Fila + 1, 10).Value = "Este Año"
'xlSheet.Cells(Fila, 11).Value = "Provision "
'xlSheet.Cells(Fila + 1, 11).Value = "Mes Actual"
'xlSheet.Cells(Fila, 12).Value = "Saldo Pend."
'xlSheet.Cells(Fila + 1, 12).Value = "De Provision"
'Fila = Fila + 2
'xlSheet.Range("A" & Fila & ":S" & Fila).Merge
'xlSheet.Range("A" & Fila & ":S" & Fila) = "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
'Fila = Fila + 4
'
'Captura_Tipo_Trabajador
'
'Select Case s_TipoTrabajador_Cts
'    Case "01"
'        Sql = "SELECT a.placod,ltrim(rtrim(b.ap_pat))+' '+ltrim(rtrim(b.ap_mat))," & _
'        "CONVERT(VARCHAR(10),b.fingreso,103),CONVERT(VARCHAR(10),b.fcese,103),'',i01,i06,i07,i04,i02,p15,p10+p11+p21,p16,0,0, p24+p25,p26,0,totalremun,a.recordact,monto_inde1,monto_inde2,monto_indetope," & _
'        "monto_total,prov_anoante,ajuste_prov,provision_ano,provision_actual,saldo_prov FROM PLAPROVCTS a LEFT OUTER JOIN PLANILLAS b on (b.cia=a.cia and b.placod=a.placod and b.status!='*') WHERE " & _
'        "a.cia='" & wcia & "' and year(fechaproceso)=" & Txtano.Text & " and a.status!='*' and month(fechaproceso)=" & Cmbmes.ListIndex + 1 & " and b.tipotrabajador='01' ORDER BY A.PLACOD"
'    Case "02"
'        Sql = "SELECT a.placod,ltrim(rtrim(b.ap_pat))+' '+ltrim(rtrim(b.ap_mat))," & _
'        "CONVERT(VARCHAR(10),b.fingreso,103),CONVERT(VARCHAR(10),b.fcese,103),'',i01,i06,i07,i04,i02,p15,p10+p11+p21,p16,0,0, p24+p25,p26,P38,totalremun,a.recordact,monto_inde1,monto_inde2,monto_indetope," & _
'        "monto_total,prov_anoante,ajuste_prov,provision_ano,provision_actual,saldo_prov FROM PLAPROVCTS a LEFT OUTER JOIN PLANILLAS b on (b.cia=a.cia and b.placod=a.placod and b.status!='*') WHERE " & _
'        "a.cia='" & wcia & "' and year(fechaproceso)=" & Txtano.Text & " and a.status!='*' and month(fechaproceso)=" & Cmbmes.ListIndex + 1 & " and b.tipotrabajador='02' ORDER BY A.PLACOD"
'    Case "03"
'        Sql = "SELECT a.placod,ltrim(rtrim(b.ap_pat))+' '+ltrim(rtrim(b.ap_mat))," & _
'        "CONVERT(VARCHAR(10),b.fingreso,103),CONVERT(VARCHAR(10),b.fcese,103),'',i01,i06,i07,i04,i02,p15,p10+p11+p21,p16,0,0, p24+p25,p26,0,totalremun,a.recordact,monto_inde1,monto_inde2,monto_indetope," & _
'        "monto_total,prov_anoante,ajuste_prov,provision_ano,provision_actual,saldo_prov FROM PLAPROVCTS a LEFT OUTER JOIN PLANILLAS b on (b.cia=a.cia and b.placod=a.placod and b.status!='*') WHERE " & _
'        "a.cia='" & wcia & "' and year(fechaproceso)=" & Txtano.Text & " and a.status!='*' and month(fechaproceso)=" & Cmbmes.ListIndex + 1 & " ORDER BY A.PLACOD"
'End Select
'
'
''sql = "SELECT a.placod,ltrim(rtrim(b.ap_pat))+' '+ltrim(rtrim(b.ap_mat))," & _
''"CONVERT(VARCHAR(10),b.fingreso,103),CONVERT(VARCHAR(10),b.fcese,103),'',i01,i06,i07,i04,i02,p15,p10+p11+p21,p16,0,0, p24+p25,p26,0,totalremun,a.recordact,monto_inde1,monto_inde2,monto_indetope," & _
''"monto_total,prov_anoante,ajuste_prov,provision_ano,provision_actual,saldo_prov FROM PLAPROVCTS a LEFT OUTER JOIN PLANILLAS b on (b.cia=a.cia and b.placod=a.placod and b.status!='*') WHERE " & _
''"a.cia='" & wcia & "' and year(fechaproceso)=" & Txtano.Text & " and a.status!='*' and month(fechaproceso)=" & Cmbmes.ListIndex + 1 & " and b.tipotrabajador='01' ORDER BY A.PLACOD"
'
'If (fAbrRst(rs, Sql)) Then
'    Do While Not rs.EOF
'
'        Fila = Fila + 1
'        For i = 0 To rs.Fields.count - 1
'            If i = 19 Then
'                Columna = 2
'                Fila = Fila + 1
'            End If
'            Columna = Columna + 1
'            'If COLUMNA = 12 Then Stop
'            If i > 4 Then
'                If i = 19 Then
'                    xlSheet.Cells(Fila, Columna).Value = Trim(rs(i))
'                Else
'                    xlSheet.Cells(Fila, Columna).Value = Format(Trim(rs(i)), "###,###,##0.00")
'                    xlSheet.Cells(Fila, Columna).NumberFormat = "#,###,##0.00"
'                End If
'                'xlSheet.Cells(FILA, COLUMNA).Value = Format(Trim(rs(I)), "###,###,##0.00")
'                'xlSheet.Cells(FILA, COLUMNA).Value = Trim(rs(I))
'                'xlSheet.Cells(FILA, COLUMNA).NumberFormat = "#,###,##0.00"
'                If i <> 19 Then ArrTotales(i - 5) = ArrTotales(i - 5) + Val(rs(i)) Else ArrTotales(i - 5) = ""
'            Else
'                If i = 2 Then
'                    xlSheet.Cells(Fila, Columna).Value = Format(Trim(rs(i)), "MM/DD/YYYY")
'                Else
'                    xlSheet.Cells(Fila, Columna).Value = Trim(rs(i))
'                End If
'                'xlSheet.Cells(FILA, COLUMNA).Value = Trim(rs(I))
'
'            End If
'        Next
'        Columna = 0
'
'        rs.MoveNext
'    Loop
'End If
'
'Columna = 5
'Fila = Fila + 3
'For i = 0 To UBound(ArrTotales)
'    If i = 14 Then Columna = 2: Fila = Fila + 1
'    Columna = Columna + 1
'    If i <> 14 Then
'        xlSheet.Cells(Fila, Columna).Value = Format(Trim(ArrTotales(i)), "###,###,##0.00")
'        xlSheet.Cells(Fila, Columna).NumberFormat = "#,###,##0.00"
'    Else
'        xlSheet.Cells(Fila, Columna).Value = ArrTotales(i)
'    End If
'Next i
'
'xlApp2.Application.ActiveWindow.DisplayGridlines = False
'xlSheet.Range("A1:A1").Select
'xlApp2.Application.Caption = "PROVISION DE CTS"
'xlApp2.ActiveWindow.Zoom = 80
'xlApp2.Application.Visible = True

End Sub

'**************codigo nuevo giovanni 17082007*********************************
Sub Generar_Asientos_Contables_Provision()
    Dim s_Centro_Costo As String
    Dim mCC As String
    Dim s_Cuenta9 As String
    Dim s_CuentaProvision As String
    Dim s_Cuenta6 As String
    Dim s_Cuenta7 As String
    Dim i_Contador_Vueltas As Integer
    Dim s_Tipo_Trabajador_Prov As String
    Dim Cadena As String

On Error GoTo MyErr
    If MsgBox("Desea generar asiento con destino?", vbYesNo + vbQuestion, "Sistema") = vbYes Then
        Bol_ConDestino = True
    Else
        Bol_ConDestino = False
    End If
    panel.Caption = "Generando Asientos de Provisión ..."
    panel.Visible = True
    panel.ZOrder 0
    Me.Refresh
    pBar.Min = 0
    pBar.Value = 0
    Call Captura_Mes_Seleccionado
    'Call Elimina_Registros_Existentes(wcia, "14")
    Cadena = "DELETE FROM ASIENTOS_PLA WHERE PLA_AÑO = " & CInt(Txtano.Text) & " AND PLA_MES = '" & s_MesSeleccion & "' AND PLA_CIA = '" & wcia & "' AND PLA_BOLETA = '14'"
    On Error Resume Next
    cn.Execute Cadena
    
    Call Recupera_Provision_Liquidacion(wcia, Txtano, s_MesSeleccion)
    Set rs_Liquidacion = Reportes_Centrales.rs_RptCentrales_pub
    Set Reportes_Centrales.rs_RptCentrales_pub = Nothing
    If rs_Liquidacion.EOF = False Then
        pBar.Max = rs_Liquidacion.RecordCount
        Do While Not rs_Liquidacion.EOF
            pBar.Value = rs_Liquidacion.AbsolutePosition
            If Recupera_Centro_Costo(wcia, rs_Liquidacion!ccosto) = True Then
                s_Centro_Costo = RTrim(Crear_Plan_Contable.s_CentroCostoPub)
                mCC = Trim(Crear_Plan_Contable.mCentroCosto)
                Dim TMP_CC As String
                TMP_CC = getCC_UltimaPlanilla(wcia, rs_Liquidacion!Codigo, CInt(Txtano), CInt(s_MesSeleccion))
                If TMP_CC <> "" Then
                    s_Centro_Costo = TMP_CC
                End If
                
                If s_Centro_Costo <> "" Then
                    Call Recupera_Informacion_Parametros_Asiento_Provision(wcia, rs_Liquidacion!tipo, _
                    "TIEMPO SERVICIO")
                    Set rs_Liquidacion2 = Reportes_Centrales.rs_RptCentrales_pub
                    Set Reportes_Centrales.rs_RptCentrales_pub = Nothing
                    s_Cuenta9 = s_Centro_Costo & rs_Liquidacion2!ctacentrocosto
                    s_CuentaProvision = rs_Liquidacion2!CtaProvision
                    Set rs_Liquidacion2 = Nothing
                    Call Codigo_Empresa_Starsoft
                    Call Conectar_Base_Datos_Access(s_CodEmpresa_Starsoft)
                    Call Recuperar_Cuentas_Naturaleza_Transferencia(s_Cuenta9)
                    Set rs_Liquidacion2 = CompRetenciones.rs_compRetenciones_Pub
                    Set CompRetenciones.rs_compRetenciones_Pub = Nothing
                    If rs_Liquidacion2.EOF Then
                        s_Cuenta6 = ""
                        s_Cuenta7 = ""
                    Else
                        s_Cuenta6 = rs_Liquidacion2!plancta_Cargo1
                        s_Cuenta7 = rs_Liquidacion2!plancta_abono1
                    End If
                    
                    Call Cerrar_Conexion_Base_Datos_Access
                    Call Genera_Numero_Voucher
                    Select Case rs_Liquidacion!tipo
                        Case "01": s_Tipo_Trabajador_Prov = "1"
                        Case "02": s_Tipo_Trabajador_Prov = "0"
                    End Select
                    For i_Contador_Vueltas = 1 To 4
                        Select Case i_Contador_Vueltas
                            Case 1
                                Call Ejecuta_Asientos_Contables(wcia, rs_Liquidacion!Codigo, Txtano, s_MesSeleccion, _
                                "00", i_Numero_VoucherG, s_Cuenta9, "", rs_Liquidacion!provmes, 1, 0, 14, s_Tipo_Trabajador_Prov, 14, mCC)
                            Case 2
                                If Bol_ConDestino Then
                                    Call Ejecuta_Asientos_Contables(wcia, rs_Liquidacion!Codigo, Txtano, s_MesSeleccion, _
                                    "00", i_Numero_VoucherG, s_Cuenta6, "", rs_Liquidacion!provmes, 1, 0, 14, s_Tipo_Trabajador_Prov, 14, mCC)
                                End If
                            Case 3
                                If Bol_ConDestino Then
                                    Call Ejecuta_Asientos_Contables(wcia, rs_Liquidacion!Codigo, Txtano, s_MesSeleccion, _
                                    "00", i_Numero_VoucherG, s_Cuenta7, "", rs_Liquidacion!provmes, 2, 1, 14, s_Tipo_Trabajador_Prov, 14, mCC)
                                End If
                            Case 4
                                Call Ejecuta_Asientos_Contables(wcia, rs_Liquidacion!Codigo, Txtano, s_MesSeleccion, _
                                "00", i_Numero_VoucherG, s_CuentaProvision, "", rs_Liquidacion!provmes, 2, 1, 14, s_Tipo_Trabajador_Prov, 14, mCC)
                        End Select
                    Next i_Contador_Vueltas
                    
                End If
            Else
                MsgBox "Personal Sin Centro de Costo: " & rs_Liquidacion!Codigo, vbCritical, App.Title
            End If
            rs_Liquidacion.MoveNext
        Loop
    End If
    'MsgBox "Asientos Generados Satisfactoriamente"
    panel.Caption = "Completo ..."
    MsgBox "Asientos Generados Satisfactoriamente", vbOKOnly + vbInformation, "Sistemas"
    panel.Visible = False
    Exit Sub
MyErr:
MsgBox Err.Description
panel.Visible = False
Me.MousePointer = 1
Exit Sub
Resume
End Sub
Sub Captura_Mes_Seleccionado()
    Select Case Cmbmes.Text
        Case "ENERO": s_MesSeleccion = "01": Case "FEBRERO": s_MesSeleccion = "02"
        Case "MARZO": s_MesSeleccion = "03": Case "ABRIL": s_MesSeleccion = "04"
        Case "MAYO": s_MesSeleccion = "05": Case "JUNIO": s_MesSeleccion = "06"
        Case "JULIO": s_MesSeleccion = "07": Case "AGOSTO": s_MesSeleccion = "08"
        Case "SETIEMBRE": s_MesSeleccion = "09": Case "OCTUBRE": s_MesSeleccion = "10"
        Case "NOVIEMBRE": s_MesSeleccion = "11": Case "DICIEMBRE": s_MesSeleccion = "12"
    End Select
End Sub
Sub Ejecuta_Asientos_Contables(CodCompañia As String, CodTrabajador As String, Año_Proceso As _
String, MesSeleccion As String, SemProceso As String, Voucher As String, CtaContable As String, _
DescCta As String, MontoInt As Double, Opcion As Integer, TipoAsiento As String, TipoBoleta As _
String, TipoTrabajador As String, Tipo_Boleta As String, CentroCosto As String)
    If Verifica_Existencia_Registro(CodCompañia, CodTrabajador, Año_Proceso, MesSeleccion, _
    SemProceso, Voucher, CtaContable, Tipo_Boleta, CentroCosto, TipoAsiento) = False Then
        Select Case Opcion
            Case 1
                Call Graba_Asiento_Contable(CodCompañia, CodTrabajador, Año_Proceso, MesSeleccion, _
                SemProceso, Voucher, CtaContable, TipoAsiento, DescCta, 0, MontoInt, TipoBoleta, _
                TipoTrabajador, CentroCosto, pla_area)
            Case 2
                Call Graba_Asiento_Contable(CodCompañia, CodTrabajador, Año_Proceso, MesSeleccion, _
                SemProceso, Voucher, CtaContable, TipoAsiento, DescCta, MontoInt, 0, TipoBoleta, _
                TipoTrabajador, CentroCosto, pla_area)
        End Select
    End If
End Sub
Sub Codigo_Empresa_Starsoft()
'    Call Recuperar_Codigo_Empresa_Starsoft(wcia)
'    Set rs_Liquidacion2 = Reportes_Centrales.rs_RptCentrales_pub
'    Set Reportes_Centrales.rs_RptCentrales_pub = Nothing
'    s_CodEmpresa_Starsoft = rs_Liquidacion2!ciastar
'    Set rs_Liquidacion2 = Nothing
    Call Trae_Cia_StarSoft(wcia, Txtano.Text)
    Set rs_Liquidacion2 = Reportes_Centrales.rs_RptCentrales_pub
    Set Reportes_Centrales.rs_RptCentrales_pub = Nothing
    s_CodEmpresa_Starsoft = rs_Liquidacion2!EMP_ID
    Set rs_Liquidacion2 = Nothing
End Sub
Sub Genera_Numero_Voucher()
    i_Numero_Voucher = i_Numero_Voucher + 1
    Select Case i_Numero_Voucher
        Case Is < 10: i_Numero_VoucherG = "000" & i_Numero_Voucher
        Case Is < 100: i_Numero_VoucherG = "00" & i_Numero_Voucher
        Case Is < 1000: i_Numero_VoucherG = "0" & i_Numero_Voucher
        Case Is < 10000: i_Numero_VoucherG = i_Numero_Voucher
    End Select
End Sub
Sub Captura_Tipo_Trabajador()
    'Captura_Tipo_Trabajador
    Select Case Cmbtipo.Text
        Case "EMPLEADO": s_TipoTrabajador_Cts = "01"
        Case "OBRERO": s_TipoTrabajador_Cts = "02"
        Case "TOTAL": s_TipoTrabajador_Cts = "03"
    End Select
End Sub

Private Sub Insertar_PlanillaCTS(ByVal Scia As String, ByVal Saño As String, ByVal Smes As String)

If Trim(Cmbmes.ListIndex + 1) = "4" Or Trim(Cmbmes.ListIndex + 1) = "10" Then
'    Sql = " Select Cia from plahistorico   where (cia = '" & Scia & "') " & _
'            "AND (YEAR(fechaproceso) =" & Saño & ") " & _
'            "AND (status <> '*')  " & _
'            "AND (MONTH(fechaproceso) =" & Trim(Cmbmes.ListIndex + 2) & ") " & _
'            "AND PROCESO ='06'"
'
'            If (fAbrRst(rsConsultaCTS, Sql)) Then
'                      If rsConsultaCTS.RecordCount > 0 Then
'                            If MsgBox("Ya se ha realizado la transferencia " & vbCrLf & "Desea actualizar la información?", vbInformation + vbYesNo, "Depositos de CTS") = vbOK Then Exit Sub
'                      End If
'            End If
'Sql = " Select Cia from plahistorico   where (cia = '" & Scia & "') " & _
'        "AND (YEAR(fechaproceso) =" & Saño & ") " & _
'        "AND (status <> '*')  " & _
'        "AND (MONTH(fechaproceso) =" & Smes & ") "
'
'        If (fAbrRst(rsConsultaCTS, Sql)) Then
'
'            MsgBox "Ya se ha realizado a la transferencia ", vbInformation, "Depositos de CTS": Exit Sub
'            Exit Sub
'        End If

Sql = "SELECT     cia, placod, tipotrab, monto_total,fechaproceso " & _
        "From plaprovcts " & _
        "WHERE     (cia = '" & Scia & "') " & _
        "AND (YEAR(fechaproceso) =" & Saño & ") " & _
        "AND (status <> '*')  " & _
        "AND (MONTH(fechaproceso) =" & Smes & ") "
        
        'rsConsultaCTS
        
        If (fAbrRst(rsConsultaCTS, Sql)) Then
             rs.MoveFirst
        Else
            MsgBox "No Hay Depositos", vbInformation, "Depositos de CTS": Exit Sub
        End If
        
Dim InTrans As Boolean
On Error GoTo MyErr
    cn.CommandTimeout = 0
    cn.BeginTrans: InTrans = True
    
    Cadena = "UPDATE PLAHISTORICO SET STATUS = '*', USER_MODI = '" & wuser & "', FEC_MODI = " & FechaSys & " WHERE CIA = '" & wcia & "' AND PROCESO = '06' AND YEAR(FECHAPROCESO) = " & CInt(Saño) & " AND MONTH(FECHAPROCESO) = " & Cmbmes.ListIndex + 2 & ""
    cn.Execute (Cadena)
Do While Not rsConsultaCTS.EOF
   
       Sql = "SET DATEFORMAT DMY " & _
             "INSERT INTO plahistorico(cia,  placod ,proceso,fechaproceso," & _
             "h01,h02,h03,h04,h05,h06,h07,h08,h09,h10,h11,h12,h13,h14,h15,h16,h17,h18,h19,h20," & _
             "h21,h22,h23,h24,h25,h26,h27,h28,h29,h30,i01,i02,i03,i04,i05,i06,i07,i08,i09,i10," & _
             "i11,i12,i13,i14,i15,i16,i17,i18,i19,i20,i21,i22,i23,i24,i25,i26,i27,i28,i29,i30," & _
             "i31,i32,i33,i34,i35,i36,i37,i38,i39,i40,i41,i42,i43,i44,i45,i46,i47,i48,i49,i50," & _
             "d01,d02,d03,d04,d05,d06,d07,d08,d09,d10,d11 , d12, d13, d14, d15, d16, d17, d18, d19, d20," & _
             "a01,a02,a03,a04,a05,a06,a07,a08,a09,a10,a11,a12,a13,a14,a15,a16,a17,a18,a19,a20," & _
             "saldoantcte,prestamo," & _
             "totaling,totneto," & _
             "fec_crea,codafp," & _
             "status," & _
             "d111,d112,d113,d114,d115," & _
             "ccosto1,ccosto2,ccosto3,ccosto4,ccosto5," & _
             "porc1,porc2,porc3,porc4,porc5," & _
             "basico,totalded,user_crea)" & _
        "VALUES( '" & Trim(rsConsultaCTS!cia) & "',  '" & Trim(rsConsultaCTS!PlaCod) & "' ,'06','15/" & Format(Cmbmes.ListIndex + 2, "00") & "/" & Saño & "'," & _
             "0,0,0,0,0,0,0,0,0,0," & _
             "0,0,0,0,0,0,0,0,0,0,"
         
        Sql = Sql & "0,0,0,0,0,0,0,0,0,0," & _
              "0,0,0,0,0,0,0,0,0,0," & _
              "0,0,0,0,0,0,0,0,0,0," & _
              "0,0,0,0,0,0,0,0,0,0," & _
              "0,0,0,0,0,0,0,0,0,0," & _
              "0,0,0,0," & rsConsultaCTS!monto_total & ",0,0,0,0,0," & _
              "0,0,0,0,0,0,0,0,0,0," & _
              "0,0,0,0,0,0,0,0,0,0," & _
              "0,0,0,0,0,0,0,0,0,0," & _
              "0,0,0,0,0,0,0,0,0,0," & _
              "0,0," & _
              " " & rsConsultaCTS!monto_total & "," & _
              " " & rsConsultaCTS!monto_total & "," & _
              " " & FechaSys & ",''," & _
              "'T'," & _
              "0,0,0,0,0," & _
              "0,0,0,0,0," & _
              "0,0,0,0,0," & _
              " 0,0,'" & wuser & "')"
              
        cn.Execute (Sql)
        
      rsConsultaCTS.MoveNext
      
Loop
    cn.CommitTrans: InTrans = False

MyErr:
    If InTrans Then cn.RollbackTrans
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbExclamation, Err.Source
        Err.Clear
    End If
    MsgBox "Se completo el proceso.", vbInformation + vbOKOnly, "Depositos de CTS"
Else
    MsgBox "Solamente puede Actualizar Planilla, para los periodos Correspondientes ( MAYO - NOVIEMBRE )", vbInformation + vbOKOnly, "Depositos de CTS"
    Exit Sub
End If
End Sub

Private Function Val_Fec_Ingreso_Inicio(ByVal pFecInicio As String, ByVal pFecIngreso As String) As String
Dim Temporal As String
Temporal = Empty
If CDate(pFecIngreso) < CDate(pFecInicio) Then
    Temporal = pFecInicio
Else
    Temporal = pFecIngreso
End If
Val_Fec_Ingreso_Inicio = Temporal
End Function
Private Function CalcularMeses1(ByVal pFecInicio As String, ByVal pFecfin As String, ByVal pFecIngreso As String, Codigo As String) As String
Dim mesestmp As String
Dim año As String, Mes As String, Dia As String
Dim FecIngTmp As String
Dim pFecproc As String
Dim pFecValida As String
Dim lFeb As Boolean
lFeb = False


Dim rsExluidos As ADODB.Recordset
Sql$ = "usp_Cts_Cargar_Excluidos '" & wcia & "'," & Txtano.Text & "," & Cmbmes.ListIndex + 1 & ",'" & Codigo & "'"
If (fAbrRst(rsExluidos, Sql$)) Then
   pFecInicio = Format(Day(rsExluidos!fIngreso), "00") + "/" + Format(Month(rsExluidos!fIngreso), "00") + "/" & Format(Year(rsExluidos!fIngreso), "0000")
End If
rsExluidos.Close: Set rsExluidos = Nothing

año = 0: Mes = 0: Dia = 0

If CDate(pFecIngreso) < CDate(pFecInicio) Then
    pFecIngreso = pFecInicio
End If

If CDate(pFecIngreso) > CDate(pFecfin) Then GoTo Salir
      
    'mes = DateDiff("m", pFecIngreso, DateAdd("d", -1, DateAdd("m", 1, pFecfin)))
    Mes = Abs(DateDiff("m", CDate(pFecIngreso), CDate(pFecfin)))
                    
    If Day(pFecIngreso) = 1 Then
        Dia = 0
        Mes = Mes + 1
        pFecproc = DateAdd("d", -1, DateAdd("m", 1, pFecIngreso))
    Else
        If Month(pFecIngreso) = Month(pFecfin) Then
            'Desde ene-2012
            
            'Dia = IIf(Day(pFecfin) = Day(pFecIngreso), 0, IIf(Month(pFecIngreso) = 2, (28 - Day(pFecIngreso)), (30 - Day(pFecIngreso))))
            Dia = IIf(Day(pFecfin) = Day(pFecIngreso), 0, IIf(Month(pFecIngreso) = 2, (30 - Day(pFecIngreso)), (30 - Day(pFecIngreso))))
                       
            'pFecproc = IIf(Month(pFecIngreso) = 2, "28/", "30/") & Month(pFecIngreso) & "/" & Year(pFecIngreso)
            'Dia = DateDiff("d", pFecIngreso, pFecproc)
        Else
            'pFecIngreso = DateAdd("d", 0, "" & Day(CDate(pFecIngreso)) & "/" & Month(DateAdd("m", 0, pFecfin)) & "/" & Year(DateAdd("m", 0, pFecfin)))
            If Month(DateAdd("m", 0, pFecfin)) = 2 And Day(CDate(pFecIngreso)) = 30 Then
               pFecIngreso = DateAdd("d", 0, "" & 28 & "/" & Month(DateAdd("m", 0, pFecfin)) & "/" & Year(DateAdd("m", 0, pFecfin)))
            Else
                If Month(pFecfin) = 2 Then
                    'pFecIngreso = DateAdd("d", 0, "" & 28 & "/" & Month(DateAdd("m", 0, pFecfin)) & "/" & Year(DateAdd("m", 0, pFecfin)))
                    'pFecIngreso = DateAdd("d", -(30 - Day(pFecIngreso)), pFecfin)
                    Dia = 30 - Day(pFecIngreso) + 1
                    lFeb = True
                Else
                    'If Day(CDate(pFecIngreso)) = 31 And (Month(pFecfin) = 4 Or Month(pFecfin) = 6 Or Month(pFecfin) = 9 Or Month(pFecfin) = 11 ) Then
                    If Day(CDate(pFecIngreso)) = 31 And (Month(pFecfin) = 4 Or Month(pFecfin) = 6 Or Month(pFecfin) = 9 Or Month(pFecfin) = 11 Or Month(pFecfin) = 10) Then
                        pFecIngreso = DateAdd("d", 0, "" & Day(CDate(pFecIngreso)) - 1 & "/" & Month(DateAdd("m", 0, pFecfin)) & "/" & Year(DateAdd("m", 0, pFecfin)))
                    Else
                        pFecIngreso = DateAdd("d", 0, "" & Day(CDate(pFecIngreso)) & "/" & Month(DateAdd("m", 0, pFecfin)) & "/" & Year(DateAdd("m", 0, pFecfin)))
                    End If
                End If
              
            End If
            If Month(pFecIngreso) = 2 Then
               pFecproc = "28/" & Month(pFecIngreso) & "/" & Year(pFecIngreso)
            Else
               pFecproc = "30/" & Month(pFecIngreso) & "/" & Year(pFecIngreso)
            End If
            
            If Not lFeb Then Dia = DateDiff("d", pFecIngreso, pFecproc)
            
            'pFecproc = DateAdd("d", Val(Dia), pFecIngreso)
        End If
        If Not lFeb Then Dia = Dia + 1
    End If
    
               
    'If Day(pFecIngreso) = 1 Then mes = CStr(Val(mes) + 1)
    
Salir:
'If mes > 6 Then mes = mes - 6

CalcularMeses1 = Space(2 - Len(Trim(año))) & año & " " & Space(2 - Len(Trim(Mes))) & Mes & " " & Space(2 - Len(Trim(Dia))) & Dia & " "
End Function

