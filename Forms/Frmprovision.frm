VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form Frmprovision 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Provisiones"
   ClientHeight    =   7275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7440
   Icon            =   "Frmprovision.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   7440
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FramePerdidas 
      BackColor       =   &H00C0C0C0&
      Height          =   5295
      Left            =   0
      TabIndex        =   21
      Top             =   840
      Visible         =   0   'False
      Width           =   7335
      Begin VB.Frame FrameEditPerdida 
         Height          =   2295
         Left            =   120
         TabIndex        =   23
         Top             =   2880
         Width           =   7095
         Begin VB.TextBox Txtcodpla 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1200
            MaxLength       =   8
            TabIndex        =   24
            Top             =   360
            Width           =   855
         End
         Begin VB.Label LblTipoTrab 
            Height          =   135
            Left            =   1200
            TabIndex        =   30
            Top             =   840
            Width           =   495
         End
         Begin MSForms.CommandButton CommandButton1 
            Height          =   495
            Left            =   2880
            TabIndex        =   29
            Top             =   1080
            Width           =   1425
            BackColor       =   -2147483637
            Caption         =   "Eliminar"
            PicturePosition =   327683
            Size            =   "2514;873"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton CommandButton2 
            Height          =   495
            Left            =   1200
            TabIndex        =   28
            Top             =   1080
            Width           =   1425
            BackColor       =   -2147483637
            Caption         =   "Aepttar"
            PicturePosition =   327683
            Size            =   "2514;873"
            Picture         =   "Frmprovision.frx":030A
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton CommandButton3 
            Height          =   495
            Left            =   5400
            TabIndex        =   27
            Top             =   1080
            Width           =   1440
            BackColor       =   12632256
            Caption         =   "Salir"
            PicturePosition =   327683
            Size            =   "2540;873"
            Picture         =   "Frmprovision.frx":08A4
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Trabajador"
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
            Index           =   0
            Left            =   120
            TabIndex        =   26
            Top             =   405
            Width           =   930
         End
         Begin VB.Label Lblnombre 
            BackColor       =   &H80000009&
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
            Height          =   255
            Left            =   2040
            TabIndex        =   25
            Top             =   390
            Width           =   4815
         End
      End
      Begin MSDataGridLib.DataGrid DtgPerdida 
         Bindings        =   "Frmprovision.frx":0BBE
         Height          =   2520
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   7065
         _ExtentX        =   12462
         _ExtentY        =   4445
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
         ColumnCount     =   2
         BeginProperty Column00 
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
         BeginProperty Column01 
            DataField       =   "Nombre"
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
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   870.236
            EndProperty
            BeginProperty Column01 
               Object.Visible         =   -1  'True
               ColumnWidth     =   5669.858
            EndProperty
         EndProperty
      End
   End
   Begin VB.CommandButton CmdVacPerd 
      Caption         =   "Vacaciones Perdidas"
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
      Left            =   5640
      TabIndex        =   20
      Top             =   6660
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton CmdProVacaciones 
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
      Left            =   2520
      TabIndex        =   18
      Top             =   6660
      Width           =   2415
   End
   Begin Threed.SSPanel Panelprogress 
      Height          =   735
      Left            =   840
      TabIndex        =   16
      Top             =   3240
      Visible         =   0   'False
      Width           =   5655
      _Version        =   65536
      _ExtentX        =   9975
      _ExtentY        =   1296
      _StockProps     =   15
      ForeColor       =   4210752
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FloodColor      =   4210752
      Alignment       =   6
      Begin MSComctlLib.ProgressBar Barra 
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Reporte"
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
      Top             =   6660
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Calcular"
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
      Top             =   6660
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   480
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   7455
      Begin VB.ComboBox Cmbcia 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   75
         Width           =   6255
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
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   825
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   5310
      Left            =   0
      TabIndex        =   5
      Top             =   930
      Width           =   7455
      Begin MSAdodcLib.Adodc AdoCabeza 
         Height          =   330
         Left            =   1200
         Top             =   3000
         Visible         =   0   'False
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   582
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
      Begin MSDataGridLib.DataGrid Dgrdcabeza 
         Bindings        =   "Frmprovision.frx":0BD7
         Height          =   5160
         Left            =   0
         TabIndex        =   19
         Top             =   0
         Width           =   7305
         _ExtentX        =   12885
         _ExtentY        =   9102
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
         ColumnCount     =   3
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
            DataField       =   "provmes"
            Caption         =   "Provision"
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
            DataField       =   "placod"
            Caption         =   "codigo"
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
               ColumnWidth     =   5715.213
            EndProperty
            BeginProperty Column01 
               Alignment       =   1
               Object.Visible         =   -1  'True
               ColumnWidth     =   824.882
            EndProperty
            BeginProperty Column02 
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   0
      TabIndex        =   0
      Top             =   465
      Width           =   7455
      Begin VB.ComboBox Cmbtipo 
         Height          =   315
         ItemData        =   "Frmprovision.frx":0BEF
         Left            =   5040
         List            =   "Frmprovision.frx":0BF1
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   75
         Width           =   2295
      End
      Begin VB.TextBox Txtano 
         Height          =   315
         Left            =   2775
         TabIndex        =   2
         Top             =   75
         Width           =   615
      End
      Begin VB.ComboBox Cmbmes 
         Height          =   315
         ItemData        =   "Frmprovision.frx":0BF3
         Left            =   720
         List            =   "Frmprovision.frx":0C1B
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   75
         Width           =   2055
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Trabajador"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3840
         TabIndex        =   12
         Top             =   75
         Width           =   1125
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Periodo"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   75
         Width           =   540
      End
      Begin MSForms.SpinButton SpinButton1 
         Height          =   315
         Left            =   3375
         TabIndex        =   3
         Top             =   75
         Width           =   255
         Size            =   "450;556"
      End
   End
   Begin MSAdodcLib.Adodc AdoVacPerd 
      Height          =   330
      Left            =   0
      Top             =   6240
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   582
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
   Begin VB.Label Lbltipo 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   3240
      TabIndex        =   14
      Top             =   6300
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Label Lbltotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6000
      TabIndex        =   10
      Top             =   6300
      Width           =   1335
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Provisión"
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
      Left            =   4440
      TabIndex        =   9
      Top             =   6300
      Width           =   1230
   End
End
Attribute VB_Name = "Frmprovision"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nFil As Integer
Dim nCol As Integer
'Dim xlApp2 As Excel.Application
'Dim xlApp1  As Excel.Application
'Dim xlBook As Excel.Workbook
'Dim xlSheet As Excel.Worksheet


Dim xlApp2 As Object
Dim xlApp1  As Object
Dim xlBook As Object
Dim xlSheet As Object


'***********codigo nuevo giovanni 17082007*********************
Dim s_MesSeleccion As String
Dim rs_Liquidacion As ADODB.Recordset
Dim rs_Liquidacion2 As ADODB.Recordset
Dim s_CodEmpresa_Starsoft As String
Dim i_Numero_Voucher As Integer
Dim i_Numero_VoucherG As String
Dim s_Fecha_Ingreso As String
Dim s_Dia_Proceso_Prov As String
'***************************************************************

Dim VTipo As String
Dim mHourMonth As Integer
Dim rs2 As ADODB.Recordset
Dim Sql As String

Dim ArrReporte() As Variant

Const COL_CODIGO = 0
Const COL_NOMBRE = 1
Const COL_RCDAÑOANTERIOR = 2
Const COL_VACACTOMADAS = 3
Const COL_RCDAÑOACTUAL = 4
Const COL_RCDACUMULADO = 5

Private Sub Cmbcia_Click()
Call fc_Descrip_Maestros2("01055", "", Cmbtipo)
Cmbtipo.AddItem "TOTAL"
Cmbtipo.ItemData(Cmbtipo.NewIndex) = "99"
End Sub

Private Sub CmbMes_Click()
Procesa_Consultas
End Sub

Private Sub CmbTipo_Click()
Dim wciamae As String
VTipo = fc_CodigoComboBox(Cmbtipo, 2)
wciamae = Determina_Maestro("01076")

If VTipo = "01" Then
   Sql = "Select flag2 from maestros_2 where cod_maestro2='04' and status<>'*'"
Else
   Sql = "Select flag2 from maestros_2 where cod_maestro2='01' and status<>'*'"
End If
Sql = Sql$ & wciamae
mHourMonth = 0
If (fAbrRst(rs, Sql)) Then mHourMonth = Val(rs!flag2)
rs.Close
Procesa_Consultas
End Sub
Private Sub CmdProVacaciones_Click()

If Txtano.Text < 2013 Then
   MsgBox "En este sistema se pueden trabajar provisiones desde Setiembre del 2013", vbInformation
   Exit Sub
End If
If Txtano.Text = 2013 And Cmbmes.ListIndex + 1 < 9 Then
   MsgBox "En este sistema se pueden trabajar provisiones desde Setiembre del 2013", vbInformation
   Exit Sub
End If


Dim strPeriodo As String, strTipotrabajador As String, strTipoProvision As String
Dim strTitulo As String, strLote As String, strVoucher As String, strSubDiario As String
strLote = "": strVoucher = "": strTitulo = "": strSubDiario = ""

strTipoProvision = IIf(Lbltipo.Caption = "V", "02", "03")

strPeriodo = Txtano.Text + Format(Cmbmes.ListIndex + 1, "00")
strTipotrabajador = Format(Cmbtipo.ListIndex + 1, "00")

Sql = "spPlaSeteoAsientoLotes '" & strTipoProvision & "','" & strTipotrabajador & "'"
If (fAbrRst(rs, Sql)) Then
    strLote = rs(0).Value
    strVoucher = rs(1).Value
    strTitulo = rs(2).Value
    strSubDiario = rs(3).Value
End If
rs.Close: Set rs = Nothing
Call GenerarAsientoProvision(strPeriodo, strTipoProvision, strTipotrabajador, strLote, strVoucher, strTitulo, Cmbmes.Text, strSubDiario)

'Asiento de Vacaciones Tomadas
'Sql = "uSp_Pla_Asiento '" & wcia & "'," & TxtAño.Text & "," & s_MesSeleccion & ",'" & s_Tipo_Boleta & "','" & s_TipoTrabajador & "','" & mLote & "','" & mVoucher & "'," & mIdDiario & ",'" & wuser & "','" & wNamePC & "','N'"

End Sub

Private Sub CmdVacPerd_Click()
If VTipo = "" Then Exit Sub
Txtcodpla.Text = ""
Lblnombre.Caption = ""
LblTipoTrab.Caption = ""

Carga_Vaca_Perdida
   
FramePerdidas.Visible = True
End Sub
Private Sub Carga_Vaca_Perdida()
Sql$ = "Usp_Pla_Vaca_Perdidas '" & wcia & "','" & VTipo & "','" & Val(Txtano.Text) & "','" & Cmbmes.ListIndex + 1 & "'"
cn.CursorLocation = adUseClient
Set AdoVacPerd.Recordset = cn.Execute(Sql$, 64)
If AdoVacPerd.Recordset.RecordCount > 0 Then AdoVacPerd.Recordset.MoveFirst
End Sub

Private Sub Command1_Click()
If VTipo = "" Or VTipo = "99" Then MsgBox "Debe Indicar Tipo de Trabajador", vbInformation, "Provisiones": Exit Sub
'If Lbltipo.Caption = "V" Then Reporte_Provision_Vaca
If Lbltipo.Caption = "V" Then ReporteVaca
If Lbltipo.Caption = "G" Then Reporte_Provision_Grati
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

If VTipo = "" Or VTipo = "99" Then MsgBox "Debe Indicar Tipo de Trabajador", vbInformation, "Provisiones": Exit Sub
Panelprogress.Caption = "Preparando para generar Provisión..."
Panelprogress.Visible = True
Panelprogress.ZOrder 0
Me.Refresh
Barra.Min = 0
Barra.Value = 0
If Lbltipo.Caption = "V" Then PROVISION_VACACIONES ' Calcula_Provision_Vaca
If Lbltipo.Caption = "G" Then PROVISIONES_GRATI 'Calcula_Provision_Grati
Procesa_Consultas
Panelprogress.Visible = False
End Sub

Private Sub CommandButton1_Click()
Sql$ = "Usp_Pla_Vaca_Perdidas_Registro '" & wcia & "','" & LblTipoTrab.Caption & "','" & Txtcodpla.Text & "','" & Val(Txtano.Text) & "','" & Cmbmes.ListIndex + 1 & "','" & wuser & "',3"
cn.Execute Sql$

Txtcodpla.Text = ""
Lblnombre.Caption = ""
LblTipoTrab.Caption = ""

Carga_Vaca_Perdida

End Sub

Private Sub CommandButton2_Click()
Sql$ = "Usp_Pla_Vaca_Perdidas_Registro '" & wcia & "','" & LblTipoTrab.Caption & "','" & Txtcodpla.Text & "','" & Val(Txtano.Text) & "','" & Cmbmes.ListIndex + 1 & "','" & wuser & "',1"
cn.Execute Sql$

Txtcodpla.Text = ""
Lblnombre.Caption = ""
LblTipoTrab.Caption = ""

Carga_Vaca_Perdida
End Sub

Private Sub CommandButton3_Click()
Txtcodpla.Text = ""
Lblnombre.Caption = ""
LblTipoTrab.Caption = ""
FramePerdidas.Visible = False
End Sub

Private Sub DtgPerdida_DblClick()
If AdoVacPerd.Recordset.RecordCount <= 0 Then Exit Sub

Txtcodpla.Text = Trim(DtgPerdida.Columns(0))
Lblnombre.Caption = Trim(DtgPerdida.Columns(1))
LblTipoTrab.Caption = VTipo
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
Me.Height = 7650
Me.Width = 7530
Call rCarCbo(Cmbcia, Carga_Cia, "C", "00")
Call rUbiIndCmbBox(Cmbcia, wcia, "00")
Txtano.Text = Format(Year(Date), "0000")
Cmbmes.ListIndex = Month(Date) - 1

End Sub
Private Sub Calcula_Provision_Vaca()
Dim Dia As Integer
Dim fecproc As String
Dim fing As String
Dim fpromed As String
Dim mfecanoant As String
Dim raAnt As String
Dim raAct As String
Dim rVaca As String
Dim rAcum As String
Dim VacaPag As Integer
Dim mFactProm As Integer
Dim cont1 As Integer
Dim mCadIng As String
Dim mBase As Currency
Dim mtoting As Currency
Dim mCadProm As String
Dim mfectope As String
Dim I As Integer
Dim J As Integer
Dim x As Integer
Dim nFields As Integer
Dim mProvAnt As Currency
Dim mVacPagada As Currency
Dim mVacPorPagar As Currency
Dim mProvMes As Currency
Dim PLAS(0 To 50) As Double

Dia = Ultimo_Dia(Cmbmes.ListIndex + 1, Val(Txtano.Text))
fecproc = Format(Dia, "00") & "/" & Format(Cmbmes.ListIndex + 1, "00") & "/" & Trim(Txtano.Text)
mfecanoant = "31/12/" & Format(Val(Txtano.Text) - 1, "0000")

Sql = "select * from planillas where cia='" & wcia & "' " & _
"and tipotrabajador='" & VTipo & "' and fcese is null and status<>'*' order by placod"

If (fAbrRst(rs, Sql)) Then rs.MoveFirst

Screen.MousePointer = vbArrowHourglass
Sql$ = wInicioTrans
cn.Execute Sql$
Panelprogress.Caption = "Calculando Provision de Vacaciones"

Panelprogress.Visible = True
Panelprogress.ZOrder 0
Me.Refresh
Barra.Max = rs.RecordCount
Barra.Value = 0

Do While Not rs.EOF
   Barra.Value = rs.AbsolutePosition
   
   raAnt = "": rVaca = "": VacaPag = 0: fing = "": raAct = "": rAcum = ""
   fing = Format(rs!fIngreso, "dd/mm/yyyy")
   
   If Mid(fing, 7, 4) = Mid(fecproc, 7, 4) Then
      raAct = Format(Val(Mid(fecproc, 7, 4)) - Mid(fing, 7, 4), "0000") & "."
      raAct = raAct & Format(Val(Mid(fecproc, 4, 2)) - Mid(fing, 4, 2), "00") & "."
      raAct = raAct & Format(Val(Mid(fecproc, 1, 2)) - Mid(fing, 1, 2), "00")
   Else
      raAct = Format(Val(Mid(fecproc, 7, 4)) - Mid(mfecanoant, 7, 4), "0000") & "."
      raAct = raAct & Format(Val(Mid(fecproc, 4, 2)) - Mid(mfecanoant, 4, 2), "00") & "."
      If Val(Mid(fecproc, 1, 2)) - Mid(mfecanoant, 1, 2) < 0 Then
         raAct = raAct & "00"
      Else
         raAct = raAct & Format(Val(Mid(fecproc, 1, 2)) - Mid(mfecanoant, 1, 2), "00")
      End If
   End If
   
   If Mid(raAct, 6, 1) = "-" Then
      raAct = Format(Val(Mid(raAct, 1, 4) - 1), "0000") & "." & Format(Val(Mid(raAct, 6, 3)) + 12, "00") & "." & Right(raAct, 2)
   End If
   
   'BUSCA EL RECORDACU DEL AÑO PASADO
   Sql = "select recordacu from plaprovvaca where cia='" & rs!cia & "' and placod='" & Trim(rs!PlaCod) & "' and year(fechaproceso)=" & Val(Mid(mfecanoant, 7, 4)) & " " _
       & "and month(fechaproceso)=12 and status<>'*'"
       
   'SQL = "SELECT RECOANT AS recordacu FROM VACACIO$ WHERE PLACOD='" & Trim(rs!PLACOD) & "'"
   
   raAnt = "0000.00.00"
   If (fAbrRst(rs2, Sql)) Then
       If Not IsNull(rs2(0)) Then raAnt = rs2(0)
   End If
   
   rs2.Close
     
   Sql = "select h.fechaproceso,h.placod,p.fingreso from plahistorico h,planillas p where h.cia='" & rs!cia & "' and h.proceso='02' and h.placod='" & Trim(rs!PlaCod) & "' " _
       & "and year(fechaproceso)=" & Val(Mid(fecproc, 7, 4)) & " and month(fechaproceso)<=" & Cmbmes.ListIndex + 1 & " and h.status<>'*' " _
       & "and p.cia=h.cia and p.placod=h.placod and p.status<>'*' and h.fechaproceso>p.fingreso"
      
   If (fAbrRst(rs2, Sql)) Then rs2.MoveFirst
   cont1 = 0: VacaPag = 0
   
   Do While Not rs2.EOF
      If Month(rs2(0)) = Val(Mid(fecproc, 4, 2)) Then VacaPag = VacaPag + 1
      cont1 = cont1 + 1
      rs2.MoveNext
   Loop
    
   rVaca = Format(cont1, "0000") & ".00.00"
   rs2.Close
   
   rAcum = FECHA_ACUMULADO(rVaca, raAnt, raAct)
      
   Sql$ = "select concepto,tipo,factor_horas,sum(importe) as base from plaremunbase a where cia='" & wcia & "' " _
        & "and placod='" & Trim(rs!PlaCod) & "' and concepto<>'03' and status<>'*' group by concepto,factor_horas," & _
        "A.TIPO order by concepto"
        
   If (fAbrRst(rs2, Sql$)) Then rs2.MoveFirst
   mtoting = 0: mCadIng = "": cont1 = 1
   
   Do While Not rs2.EOF
      mBase = 0
      For I = 1 To 50
          If I = Val(rs2(0)) Then
             If Trim(VTipo) = "01" Then
                If Trim(rs2!tipo) = "04" Then
                   mBase = rs2(3)
                Else
                   mBase = Round(rs2(3) / rs2(2) * mHourMonth, 2)
                End If
             Else
                If Trim(rs2!tipo) = "01" Then
                   mBase = rs2(3)
                Else
                   mBase = Round(rs2(3) / rs2(2) * mHourMonth, 2)
                End If
             End If
             'mCadIng = mCadIng & "" & mBase & "" & ","
             mtoting = mtoting + mBase
             PLAS(cont1) = mBase
             cont1 = cont1 + 1
             Exit For
          ElseIf I > cont1 Then
             mBase = 0
             'mCadIng = mCadIng & "" & mBase & "" & ","
             PLAS(cont1) = mBase
             cont1 = cont1 + 1
          End If
      Next
      rs2.MoveNext
   Loop
    
   'TOTHXS = (RS("i10") + RS("i21") + RS("i24") + RS("I25") + RS("i11")) / 6 / 30
   
   mFactProm = 0
   
   'Promedios
   Sql = "select codinterno,factor from platasaanexo where cia='" & _
   Trim(rs!cia) & "' and modulo='01' and tipomovimiento='" & VTipo & "' and status<>'*'"
   
   If (fAbrRst(rs2, Sql$)) Then rs2.MoveFirst: mFactProm = rs2(1)
   mCadProm = ""
   nFields = 0
   Do While Not rs2.EOF
      mCadProm = mCadProm & "sum(i" & rs2(0) & " ) as i" & rs2(0) & ","
      nFields = nFields + 1
      rs2.MoveNext
   Loop
   rs2.Close
   fpromed = ""
   
   If Trim(mCadProm) <> "" Then
      mCadProm = Mid(mCadProm, 1, Len(Trim(mCadProm)) - 1)
'      mFactProm = Mid(fecproc, 4, 2) - 5
'      mFactProm = DateAdd("m", -5, fecproc)
      fpromed = Fecha_Promedios(mFactProm, fecproc)
      
      If Val(Mid(fecproc, 4, 2)) = 1 Then
         mfectope = "31/12/" & Format(Val(Mid(fecproc, 7, 4)) - 1, "0000")
      Else
         Dia = Ultimo_Dia(Cmbmes.ListIndex + 1, Val(Txtano.Text))
         mfectope = Format(Dia, "00") & "/" & Format(Cmbmes.ListIndex + 1, "00") & "/" & Mid(fecproc, 7, 4)
      End If
      
      Sql = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
      Sql = Sql & " select " & mCadProm & ",sum(I10) as i10,sum(I11) as i11,sum(I21) as i21,sum(I24) as i24,sum(I25) as i25 " & _
      "from plahistorico where cia='" & rs!cia & "' and  placod='" & _
      Trim(rs!PlaCod) & "' and fechaproceso BETWEEN '" & _
      Format(fpromed, FormatFecha) + FormatTimei & "' AND '" & Format(mfectope, FormatFecha) + FormatTimef & "' " _
      & "and proceso='01' and status<>'*'"
      
'      If Trim(rs("PLACOD")) = "O1009" Then
'         Debug.Print "SEGA"
'      End If
          
     If (fAbrRst(rs2, Sql$)) Then rs2.MoveFirst
     mCadProm = ""
        For I = 1 To 50
            mBase = 0
            For J = 0 To nFields - 1
                mBase = 0
                If Format(I, "00") = Mid(rs2(J).Name, 2, 2) Then
                   If IsNull(rs2(J)) Then mBase = 0 Else mBase = rs2(J) / mFactProm
                   Exit For
                End If
            Next J
            mtoting = mtoting + mBase
            mCadProm = mCadProm & "" & mBase & "" & ","
        Next I
        
        
        PLAS(10) = IIf(IsNull(rs2("I10")), 0, rs2("I10"))
        PLAS(21) = IIf(IsNull(rs2("i21")), 0, rs2("i21"))
        PLAS(24) = IIf(IsNull(rs2("I24")), 0, rs2("I24"))
        PLAS(25) = IIf(IsNull(rs2("I25")), 0, rs2("I25"))
        PLAS(11) = IIf(IsNull(rs2("I11")), 0, rs2("I11"))
   
        mCadIng = ""
        For I = 1 To 50
            mCadIng = mCadIng & PLAS(I) & ","
        Next
   
        'mCadIng = Left(mCadIng, Len(mCadIng) - 1)
   End If
   If rs2.State = 1 Then rs2.Close
   
   
   If Val(Mid(fecproc, 4, 2)) - 1 <= 0 Then
      mfectope = "31/01/" & Val(Mid(fecproc, 7, 4) - 1)
   Else
      mfectope = Mid(fecproc, 1, 2) & "/" & Format(Val(Mid(fecproc, 4, 2)) - 1, "00") & "/" & Mid(mfectope, 7, 4)
   End If
  
   Sql = "select provtotal from plaprovvaca where cia='" & rs!cia & _
   "' and placod='" & Trim(rs!PlaCod) & "' and year(fechaproceso)=" & Val(Mid(mfectope, 7, 4)) & " " & _
   "and month(fechaproceso)=" & Val(Mid(mfectope, 4, 2)) & " and status<>'*'"
   
   Sql = "SELECT TOTPER2 FROM VACACIO WHERE PLACOD='" & Trim(rs!PlaCod) & "'"
'   Debug.Print mfectope
   If (fAbrRst(rs2, Sql)) Then mProvAnt = rs2(0) Else mProvAnt = 0
   rs2.Close

   mVacPorPagar = 0
   If VTipo = "01" Then
      mVacPagada = mtoting * VacaPag
      If Val(Mid(raAct, 1, 4)) <> 0 Then mVacPorPagar = (Val(Mid(raAct, 1, 4)) * mtoting)
      If Val(Mid(raAct, 6, 2)) <> 0 Then mVacPorPagar = mVacPorPagar + Round(Val(Mid(raAct, 6, 2)) * mtoting / 12, 2)
      'If Val(Mid(raAct, 9, 2)) <> 0 Then mVacPorPagar = mVacPorPagar + Round(Val(Mid(raAct, 9, 2)) * mtoting / 365, 2)
      If Val(Mid(raAct, 9, 2)) <> 0 Then mVacPorPagar = mVacPorPagar + Round(Val(Mid(raAct, 9, 2)) * mtoting / 360, 2)
   Else
      mVacPagada = (mtoting * 30) * VacaPag
      If Val(Mid(raAct, 1, 4)) <> 0 Then mVacPorPagar = (Val(Mid(raAct, 1, 4)) * 30 * mtoting)
      If Val(Mid(raAct, 6, 2)) <> 0 Then mVacPorPagar = mVacPorPagar + Round(Val(Mid(raAct, 6, 2)) * 30 * mtoting / 12, 2)
      'If Val(Mid(raAct, 9, 2)) <> 0 Then mVacPorPagar = mVacPorPagar + Round(Val(Mid(raAct, 9, 2)) * 30 * mtoting / 365, 2)
      If Val(Mid(raAct, 9, 2)) <> 0 Then mVacPorPagar = mVacPorPagar + Round(Val(Mid(raAct, 9, 2)) * 30 * mtoting / 360, 2)
   End If
   
   mProvMes = mVacPorPagar - (mProvAnt - mVacPagada)
   
   If mProvMes <> 0 Then
      Sql = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
      Sql = Sql & "insert into plaprovvaca values('" & Trim(rs!cia) & "','" & Trim(rs!PlaCod) & "','" & Trim(rs!TipoTrabajador) & "', " _
      & "'" & raAnt & "','" & rVaca & "','" & raAct & "','" & rAcum & "'," & mCadIng & mCadProm
      Sql = Sql & "" & mtoting & "," & mProvAnt & "," & mVacPagada & "," & _
      mVacPorPagar & "," & mProvMes & ",'" & Format(fecproc, FormatFecha) & "'," _
      & FechaSys & ",'" & wuser & "','" & Trim(rs!Area) & "','')"
'      Debug.Print SQL
      
      cn.Execute Sql$
   End If
   
   rs.MoveNext
Loop

Sql$ = wFinTrans
cn.Execute Sql$
Panelprogress.Visible = False
Carga_Prov_Vaca
Screen.MousePointer = vbDefault
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
Procesa_Consultas
End Sub

Private Sub Txtano_KeyPress(KeyAscii As Integer)
Txtano.Text = Txtano.Text + fc_ValNumeros(KeyAscii)
End Sub
Private Sub Carga_Prov_Vaca()
Dim mcad As String
If Trim(Txtano.Text) = "" Then Exit Sub
If Cmbmes.ListIndex < 0 Then Exit Sub
If VTipo = "" Or Cmbtipo.Text = "" Then Exit Sub

If VTipo = "99" Or VTipo = "" Then mcad = "" Else mcad = " and a.tipotrab='" & VTipo & "' "

Sql = nombre()
Sql = Sql & "a.placod,a.provmes " _
& "from plaprovvaca a,planillas b " _
& "where a.cia='" & wcia & "' and year(a.fechaproceso)='" & Val(Txtano.Text) & "' and month(a.fechaproceso)='" & Cmbmes.ListIndex + 1 & "' and a.status<>'*' " _
& "and a.placod=b.placod and a.cia=b.cia and b.status<>'*'" & mcad & "order by nombre"

cn.CursorLocation = adUseClient
Set AdoCabeza.Recordset = cn.Execute(Sql$, 64)
If AdoCabeza.Recordset.RecordCount > 0 Then
   AdoCabeza.Recordset.MoveFirst
   Command4.Enabled = False
   CmdVacPerd.Enabled = False
   Sql = "select SUM(provmes) from plaprovvaca a " _
       & "where a.cia='" & wcia & "' and year(a.fechaproceso)='" & Val(Txtano.Text) & "' and month(a.fechaproceso)='" & Cmbmes.ListIndex + 1 & "' and a.status<>'*'" & mcad
   
   If (fAbrRst(rs, Sql)) Then Lbltotal.Caption = Format(rs(0), "###,###.00")
   rs.Close
Else
   Command4.Enabled = True
   Lbltotal.Caption = "0.00"
   CmdVacPerd.Enabled = True
End If
Dgrdcabeza.Refresh
End Sub
Public Sub Elimina_Prov_Vaca()
If Txtano.Text < 2013 Then
   MsgBox "En este sistema se pueden trabajar provisiones desde Setiembre del 2013", vbInformation
   Exit Sub
End If
If Txtano.Text = 2013 And Cmbmes.ListIndex + 1 < 9 Then
   MsgBox "En este sistema se pueden trabajar provisiones desde Setiembre del 2013", vbInformation
   Exit Sub
End If

If VTipo = "" Or VTipo = "99" Then MsgBox "Debe Indicar Tipo de Trabajador", vbInformation, "Provisiones": Exit Sub
    If MsgBox("Desea Eliminar Provision ? ", vbQuestion + vbYesNo + vbDefaultButton1, "Seteo de Usuario") = vbNo Then
        Exit Sub
    Else
        Sql = wInicioTrans
        cn.Execute Sql
    
        Sql = "update plaprovvaca set status='*' where cia='" & wcia & "' and year(fechaproceso)=" & Val(Txtano.Text) & " and month(fechaproceso)=" & Cmbmes.ListIndex + 1 & " and tipotrab='" & VTipo & "' and status<>'*'"
        cn.Execute Sql
    
        Sql = wFinTrans
        cn.Execute Sql
    
        Carga_Prov_Vaca
    End If
    Screen.MousePointer = vbDefault
'Else
 '   GoTo MENSAJE
'End If
Exit Sub
mensaje:
    MsgBox "No se puede Eliminar la Informacion de la Provision", vbInformation
End Sub
Private Sub Reporte_Provision_Vaca()
Dim MArea As String
Dim wciamae As String
Dim mcad As String
Dim mCadBas As String
Dim mCadProm As String
Dim mFieldBas As Integer
Dim mFieldProm As Integer
Dim MField As Integer
Dim contbase As Integer
Dim mText As String
Dim msum As Integer
Dim sumparc As Integer
Dim tot1 As Currency
Dim tot2 As Currency
Dim tot3 As Currency
Dim tot4 As Currency
Dim tot5 As Currency
Dim con As String, CON1 As String
Dim I As Integer
Dim x As Integer
Dim TOTHXS As Variant, TOTASGFAM As Variant
Dim RX As New ADODB.Recordset
'**************codigo agregado giovanni 24082007*****************************
Dim i_Contador_Filas As Integer
'****************************************************************************

'**************codigo agregado giovanni 24082007*****************************
i_Contador_Filas = 0
'****************************************************************************

If VTipo = "" Or VTipo = "99" Then MsgBox "Debe Indicar Tipo de Trabajador", vbInformation, "Provisiones": Exit Sub
If AdoCabeza.Recordset.RecordCount <= 0 Then Exit Sub

Screen.MousePointer = vbArrowHourglass
mcad = ""
For I = 1 To 50
    mcad = mcad & "sum(i" & Format(I, "00") & "),"
Next

For I = 1 To 50
    mcad = mcad & "sum(p" & Format(I, "00") & "),"
Next

If Trim(mcad) <> "" Then
   mcad = Mid(mcad, 1, Len(Trim(mcad)) - 1)
End If

'************codigo modificado giovanni 07092007**********************
'sql = "select " & mcad & " from plaprovvaca " _
'& "where cia='" & wcia & "' and year(fechaproceso)='" & Val(Txtano.Text) & "' and month(fechaproceso)='" & Cmbmes.ListIndex + 1 & "' and tipotrab='" & VTipo & "' and status<>'*' GROUP BY PLACOD ORDER BY PLACOD"
Sql = "select " & mcad & " from plaprovvaca " _
    & "where cia='" & wcia & "' and year(fechaproceso)='" & Val(Txtano.Text) & "' and month(fechaproceso)='" & Cmbmes.ListIndex + 1 & "' and tipotrab='" & VTipo & "' and status<>'*' "
'*********************************************************************
    
Panelprogress.Caption = "Generando Reporte de Provision de Vacaciones"
Panelprogress.Visible = True
Panelprogress.ZOrder 0
Me.Refresh
Barra.Value = 0
If (fAbrRst(rs, Sql)) Then
   mCadBas = ""
   mCadProm = ""
   Barra.Max = 99
   For I = 0 To 99
       Barra.Value = I
       If rs(I) <> 0 Then
          If I <= 49 Then
             mCadBas = mCadBas & Format(I + 1, "00")
          Else
             mCadProm = mCadProm & Format(I - 50 + 1, "00")
          End If
       End If
   Next

End If

If Trim(mCadBas) = "" Then mFieldBas = 0 Else mFieldBas = Len(Trim(mCadBas))
If Trim(mCadProm) = "" Then mFieldProm = 0 Else mFieldProm = Len(Trim(mCadProm))
MArea = ""
wciamae = Determina_Maestro("01044")

'****SE EMPIEZA A FORMATEAR LA HOJA DE EXCEL

Set xlApp1 = CreateObject("Excel.Application")
xlApp1.Workbooks.Add
Set xlApp2 = xlApp1.Application
Set xlBook = xlApp2.Workbooks(1)
Set xlSheet = xlApp2.Worksheets("HOJA1")

xlSheet.Range("A:A").ColumnWidth = 3
xlSheet.Range("B:B").ColumnWidth = 7
xlSheet.Range("C:C").ColumnWidth = 45
xlSheet.Range("D:G").ColumnWidth = 12
xlSheet.Range("D:X").HorizontalAlignment = xlCenter
xlSheet.Range("J:X").NumberFormat = "#,###,##0.00;[Red](#,###,##0.00)"

xlSheet.Cells(1, 2).Value = Trim(Cmbcia.Text)
xlSheet.Cells(1, 2).Font.Size = 12
xlSheet.Cells(1, 2).Font.Bold = True

xlSheet.Cells(3, 2).Value = "REPORTE DE PROVISION DE VACACIONES " & Cmbtipo.Text & " - " & Cmbmes.Text & " " & Txtano.Text
xlSheet.Cells(3, 2).Font.Size = 11
xlSheet.Cells(3, 2).Font.Bold = True

xlSheet.Range(xlSheet.Cells(6, 2), xlSheet.Cells(7, 24)).Font.Bold = True
xlSheet.Cells(6, 2).Value = "Codigo"
xlSheet.Range(xlSheet.Cells(6, 2), xlSheet.Cells(7, 2)).Merge
xlSheet.Cells(6, 2).VerticalAlignment = xlCenter

xlSheet.Cells(6, 3).Value = "Nombre Trabajador"
xlSheet.Range(xlSheet.Cells(6, 3), xlSheet.Cells(7, 3)).Merge
xlSheet.Cells(6, 3).VerticalAlignment = xlCenter

'--
xlSheet.Cells(6, 4).Value = "Fecha"
xlSheet.Cells(7, 4).Value = "Ingreso"

'xlSheet.Range(xlSheet.Cells(6, nCol), xlSheet.Cells(7, nCol)).Borders(xlEdgeLeft).LineStyle = xlContinuous
'xlSheet.Range(xlSheet.Cells(6, nCol), xlSheet.Cells(7, nCol)).Borders(xlEdgeRight).LineStyle = xlContinuous

xlSheet.Cells(6, 5).Value = "DNI"

'xlSheet.Range(xlSheet.Cells(6, nCol), xlSheet.Cells(7, nCol)).Borders(xlEdgeLeft).LineStyle = xlContinuous
'xlSheet.Range(xlSheet.Cells(6, nCol), xlSheet.Cells(7, nCol)).Borders(xlEdgeRight).LineStyle = xlContinuous

'--


xlSheet.Cells(6, 6).Value = "Record Año"
xlSheet.Cells(7, 6).Value = "Año Mes Dia"

xlSheet.Cells(6, 7).Value = "Vacacio"
xlSheet.Cells(7, 7).Value = "Tomadas"

xlSheet.Cells(6, 8).Value = "Record Año"
xlSheet.Cells(7, 8).Value = "Actual"

xlSheet.Cells(6, 9).Value = "Record Acum."
xlSheet.Cells(7, 9).Value = "Año Mes Dia"

xlSheet.Cells(6, 10) = "Jornal"
xlSheet.Cells(7, 10) = "Basico"
xlSheet.Range(xlSheet.Cells(6, 10), xlSheet.Cells(7, 10)).HorizontalAlignment = xlCenter

xlSheet.Cells(6, 11) = "Promedio"
xlSheet.Cells(7, 11) = "H. Extras"

xlSheet.Cells(6, 12) = "Asignac."
xlSheet.Cells(7, 12) = "familiar"

xlSheet.Cells(6, 13) = "Promedio"
xlSheet.Cells(7, 13) = "Produccion"

xlSheet.Cells(6, 14) = "Bonif."
xlSheet.Cells(7, 14) = "T. Serv."

xlSheet.Cells(6, 15) = "Bonif."
xlSheet.Cells(7, 15) = "Volunt."
xlSheet.Range(xlSheet.Cells(6, 15), xlSheet.Cells(7, 15)).VerticalAlignment = xlCenter

xlSheet.Cells(6, 16) = "AFP"
xlSheet.Range(xlSheet.Cells(6, 16), xlSheet.Cells(7, 16)).Merge
xlSheet.Range(xlSheet.Cells(6, 16), xlSheet.Cells(7, 16)).VerticalAlignment = xlCenter

xlSheet.Cells(6, 17) = "Costo"
xlSheet.Cells(7, 17) = "Vida"
xlSheet.Range(xlSheet.Cells(6, 17), xlSheet.Cells(7, 17)).VerticalAlignment = xlCenter

xlSheet.Cells(6, 18) = "Promedio"
xlSheet.Cells(7, 18) = "Otros Pagos"
xlSheet.Range(xlSheet.Cells(6, 18), xlSheet.Cells(7, 18)).VerticalAlignment = xlCenter

xlSheet.Cells(6, 19) = "Aportes."
xlSheet.Cells(7, 19) = "Patronal"
xlSheet.Range(xlSheet.Cells(6, 19), xlSheet.Cells(7, 19)).VerticalAlignment = xlCenter

xlSheet.Cells(6, 20) = "Remun."
xlSheet.Cells(7, 20) = "Vacac."

nCol = 21 'antes 18
contbase = 15

xlSheet.Cells(6, nCol).Value = "Mes"
xlSheet.Cells(7, nCol).Value = "Anterior"
xlSheet.Range(xlSheet.Cells(6, nCol), xlSheet.Cells(7, nCol)).Borders(xlEdgeLeft).LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(6, nCol), xlSheet.Cells(7, nCol)).Borders(xlEdgeTop).LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(6, nCol), xlSheet.Cells(7, nCol)).Borders(xlEdgeBottom).LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(6, nCol), xlSheet.Cells(7, nCol)).Borders(xlEdgeRight).LineStyle = xlContinuous

nCol = nCol + 1
xlSheet.Cells(6, nCol).Value = "Vaca."
xlSheet.Cells(7, nCol).Value = "Pagada"
xlSheet.Range(xlSheet.Cells(6, nCol), xlSheet.Cells(7, nCol)).Borders(xlEdgeLeft).LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(6, nCol), xlSheet.Cells(7, nCol)).Borders(xlEdgeRight).LineStyle = xlContinuous

nCol = nCol + 1
xlSheet.Cells(6, nCol).Value = "Vaca."
xlSheet.Cells(7, nCol).Value = "X Pagar"
xlSheet.Range(xlSheet.Cells(6, nCol), xlSheet.Cells(7, nCol)).Borders(xlEdgeLeft).LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(6, nCol), xlSheet.Cells(7, nCol)).Borders(xlEdgeRight).LineStyle = xlContinuous

nCol = nCol + 1
xlSheet.Cells(6, nCol).Value = "Prov."
xlSheet.Cells(7, nCol).Value = "Del Mes"
xlSheet.Range(xlSheet.Cells(6, nCol), xlSheet.Cells(7, nCol)).Borders(xlEdgeLeft).LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(6, nCol), xlSheet.Cells(7, nCol)).Borders(xlEdgeRight).LineStyle = xlContinuous

'nCol = nCol + 1
'xlSheet.Cells(6, nCol).Value = "Fecha"
'xlSheet.Cells(7, nCol).Value = "Ingreso"
'xlSheet.Range(xlSheet.Cells(6, nCol), xlSheet.Cells(7, nCol)).Borders(xlEdgeLeft).LineStyle = xlContinuous
'xlSheet.Range(xlSheet.Cells(6, nCol), xlSheet.Cells(7, nCol)).Borders(xlEdgeRight).LineStyle = xlContinuous
'
'
'nCol = nCol + 1
'xlSheet.Cells(6, nCol).Value = "DNI"
'xlSheet.Range(xlSheet.Cells(6, nCol), xlSheet.Cells(7, nCol)).Borders(xlEdgeLeft).LineStyle = xlContinuous
'xlSheet.Range(xlSheet.Cells(6, nCol), xlSheet.Cells(7, nCol)).Borders(xlEdgeRight).LineStyle = xlContinuous


xlSheet.Range(xlSheet.Cells(3, 2), xlSheet.Cells(3, nCol)).Merge
xlSheet.Range(xlSheet.Cells(3, 2), xlSheet.Cells(3, nCol)).HorizontalAlignment = xlCenter

xlSheet.Range(xlSheet.Cells(6, 4), xlSheet.Cells(7, 24)).Borders(xlEdgeBottom).LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(6, 4), xlSheet.Cells(7, 24)).Borders(xlEdgeTop).LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(6, 4), xlSheet.Cells(7, 24)).Borders(xlEdgeRight).LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(6, 4), xlSheet.Cells(7, 24)).Borders(xlEdgeLeft).LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(6, 4), xlSheet.Cells(7, 24)).Borders(xlInsideVertical).LineStyle = xlContinuous

xlSheet.Range(xlSheet.Cells(6, 2), xlSheet.Cells(7, nCol)).HorizontalAlignment = xlCenter
xlSheet.Range(xlSheet.Cells(6, 2), xlSheet.Cells(7, 3)).Borders.LineStyle = xlContinuous

'*************SE TERMINA FORMATEO*********************


nFil = 8
sumparc = nFil
Sql = nombre()
Sql = Sql & "b.nro_doc,a.* from plaprovvaca a,planillas b " & _
 "where a.cia='" & wcia & "' and year(a.fechaproceso)='" & Val(Txtano.Text) & "' and month(a.fechaproceso)='" & Cmbmes.ListIndex + 1 & "' and a.tipotrab='" & VTipo & "' and a.status<>'*' " _
& "and a.placod=b.placod and a.cia=b.cia and b.status<>'*' order by a.area,a.placod"

If (fAbrRst(rs, Sql)) Then rs.MoveFirst: MArea = Trim(rs!Area)
If Trim(MArea) <> "" Then
   Sql$ = "Select descripcion as descrip from pla_ccostos where cia='" & wcia & "' and codigo='" & MArea & "' and status<>'*'"
   If (fAbrRst(rs2, Sql)) Then xlSheet.Cells(nFil, 2).Value = rs2!DESCRIP
   xlSheet.Cells(nFil, 2).Font.Bold = True
   nFil = nFil + 2
   sumparc = sumparc + 2
   rs2.Close
End If

tot1 = 0: tot2 = 0: tot3 = 0: tot4 = 0: tot5 = 0

Barra.Max = rs.RecordCount
Barra.Value = 0

Do While Not rs.EOF

   Barra.Value = rs.AbsolutePosition
   If Trim(rs!Area) <> Trim(MArea) Then

      '***********codigo agregado giovanni 25082007************************
      sumparc = i_Contador_Filas + 3
      '********************************************************************
      
      msum = (sumparc - 2) * -1
      nFil = nFil + 1
      x = 1
      For I = nCol - 5 To nCol - 1
         xlSheet.Cells(nFil, I).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
         Select Case x
                Case Is = 1: tot1 = tot1 + xlSheet.Cells(nFil, I).Value
                Case Is = 2: tot2 = tot2 + xlSheet.Cells(nFil, I).Value
                Case Is = 3: tot3 = tot3 + xlSheet.Cells(nFil, I).Value
                Case Is = 4: tot4 = tot4 + xlSheet.Cells(nFil, I).Value
                Case Is = 5: tot5 = tot5 + xlSheet.Cells(nFil, I).Value
         End Select
         x = x + 1
      Next I
      sumparc = 0
      
      nFil = nFil + 2
      sumparc = sumparc + 3
      MArea = rs!Area
      Sql$ = "Select descripcion as descrip from pla_ccostos where cia='" & wcia & "' and codigo='" & MArea & "' and status<>'*'"
      If (fAbrRst(rs2, Sql)) Then xlSheet.Cells(nFil, 2).Value = rs2!DESCRIP
      xlSheet.Cells(nFil, 2).Font.Bold = True
      nFil = nFil + 2
      sumparc = sumparc + 2
      rs2.Close
      '******************codigo agregado giovanni 24082007**********************
      i_Contador_Filas = 0
      '*************************************************************************
   End If
   
   xlSheet.Cells(nFil, 2).Value = Trim(rs!PlaCod)
   xlSheet.Cells(nFil, 3).Value = Trim(rs!nombre)
   
   'FEC. INGRESO
   xlSheet.Cells(nFil, 4) = Format(Trim(rs("fingreso")), "MM/DD/YYYY")
   
   'DNI
   xlSheet.Cells(nFil, 5) = "'" + Trim(rs("nro_doc"))
   
   xlSheet.Cells(nFil, 6).Value = RTrim(rs("RECORDANOANT"))
   xlSheet.Cells(nFil, 7).Value = RTrim(rs("RECORDVACA"))
   xlSheet.Cells(nFil, 8).Value = RTrim(rs("RECORDANOACT"))
   xlSheet.Cells(nFil, 9).Value = RTrim(rs("RECORDACU"))
   
   'jornalbasico
   xlSheet.Cells(nFil, 10) = CStr(rs("i01"))
      
    'Promedio Horas Extras
   TOTHXS = CStr(rs("P10") + rs("P11") + rs("P24") + rs("P25") + rs("P21"))
   
   nCol = 11
   'total horas extras
   xlSheet.Cells(nFil, nCol) = CStr(TOTHXS)
   'asignacion familiar
   xlSheet.Cells(nFil, nCol + 1) = CStr(rs("i02"))
    
    'Promedio otros pagos, Reintegros, Bonif.xRend., Bonif.xProd
    xlSheet.Cells(nFil, nCol + 2) = CStr(rs("P38"))
   
   'BONIFICACION X TIEMPO DE SERVICIO
   xlSheet.Cells(nFil, nCol + 3) = CStr(rs("I04"))
   
   'bonif. volunt.
   xlSheet.Cells(nFil, nCol + 4) = CStr(rs("I26"))
   
   'afp
   xlSheet.Cells(nFil, nCol + 5) = CStr(rs("I05") + rs("I06"))
   
   'COSTO DE VIDA
   xlSheet.Cells(nFil, nCol + 6) = CStr(rs("I07"))
   
   'OTROS PAGOS
   xlSheet.Cells(nFil, nCol + 7) = CStr(rs("P13") + rs("P16") + rs("P37"))
   
   
   'APORTE PATRONAL
   xlSheet.Cells(nFil, nCol + 8) = CStr(rs("aportepatronal"))
   
   'REMUNERACION TOTAL
   xlSheet.Cells(nFil, nCol + 9) = CStr(rs("remtotal"))
   
   'MES ANTERIOR
   xlSheet.Cells(nFil, nCol + 10) = CStr(rs("PROVMESANT"))
   
   'VACACIONES PAGADAS
   xlSheet.Cells(nFil, nCol + 11) = CStr(rs("PROVPAGADAS"))
   
   'VACACIONES POR PAGAR
   xlSheet.Cells(nFil, nCol + 12) = CStr(rs("provtotal"))
   
   'PROV. DEL MES
   xlSheet.Cells(nFil, nCol + 13) = CStr(rs("provmes"))
   
   
   nCol = nCol + 14

   nFil = nFil + 1
   
   '***********codigo agregado giovanni 25082007***********************
   i_Contador_Filas = i_Contador_Filas + 1
   '*******************************************************************
   
   rs.MoveNext
Loop

    '****************codigo agregado giovanni 25082007*******************************
        sumparc = i_Contador_Filas + 3
        msum = (sumparc - 2) * -1
        nFil = nFil + 1
        x = 1
        For I = nCol - 5 To nCol - 1
            xlSheet.Cells(nFil, I).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
        Next I
    '*********************************************************************************

Panelprogress.Visible = False

msum = (sumparc - 2) * -1
'nFil = nFil + 1
x = 1
For I = nCol - 5 To nCol - 1
 '   xlSheet.Cells(nFil, I).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
    Select Case x
           Case Is = 1: tot1 = tot1 + xlSheet.Cells(nFil, I).Value
           Case Is = 2: tot2 = tot2 + xlSheet.Cells(nFil, I).Value
           Case Is = 3: tot3 = tot3 + xlSheet.Cells(nFil, I).Value
           Case Is = 4: tot4 = tot4 + xlSheet.Cells(nFil, I).Value
           Case Is = 5: tot5 = tot5 + xlSheet.Cells(nFil, I).Value
    End Select
    x = x + 1
Next I
sumparc = 0

nFil = nFil + 1

msum = (nFil) * -1
nFil = nFil + 1
For I = 10 To nCol - 4
   xlSheet.Cells(nFil, I).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
Next I

x = 1
For I = nCol - 5 To nCol
    Select Case x
           Case Is = 1: xlSheet.Cells(nFil, I).Value = tot1
           Case Is = 2: xlSheet.Cells(nFil, I).Value = tot2
           Case Is = 3: xlSheet.Cells(nFil, I).Value = tot3
           Case Is = 4: xlSheet.Cells(nFil, I).Value = tot4
           Case Is = 5: xlSheet.Cells(nFil, I).Value = tot5
    End Select
    
    x = x + 1
Next

xlApp2.Application.ActiveWindow.DisplayGridLines = False
xlSheet.Range("A1:A1").Select
xlApp2.Application.Caption = "PROVISION DE VACACIONES"
xlApp2.ActiveWindow.Zoom = 80
xlApp2.Application.Visible = True

If Not xlApp2 Is Nothing Then Set xlApp2 = Nothing
If Not xlApp1 Is Nothing Then Set xlApp1 = Nothing
If Not xlBook Is Nothing Then Set xlBook = Nothing
If Not xlSheet Is Nothing Then Set xlSheet = Nothing
Screen.MousePointer = 0
End Sub

Public Sub Provisiones(tipo As String)
Lbltipo.Caption = tipo

If Lbltipo.Caption = "V" Then CmdVacPerd.Visible = True Else CmdVacPerd.Visible = False

If tipo = "V" Then Me.Caption = "Provision de Vacacion"
If tipo = "G" Then Me.Caption = "Provision de Gratificacion"
If tipo = "D" Then
   Me.Caption = "Vacaciones Devengadas"
   Cmbmes.ListIndex = 0
   Cmbmes.Enabled = False
   Label3.Caption = "Total Neto"
   Frame3.BackColor = &HFF&
   Dgrdcabeza.Columns(1).Caption = "Tot. Ing."
End If
End Sub
Public Sub Procesa_Devengadas()
Dim mano As Integer
Dim mmes As Integer
Dim mcad As String
mano = Val(Txtano.Text)
mmes = Cmbmes.ListIndex + 1

If VTipo = "99" Or VTipo = "" Then mcad = "" Else mcad = " and b.tipotrabajador='" & VTipo & "' "

Sql$ = nombre()
Sql$ = Sql$ & "a.placod,a.totaling AS provmes " _
     & "from plahistorico a,planillas b " _
     & "where a.cia='" & wcia & "' and a.proceso='02' and year(a.fechaproceso)=" & mano & " and month(a.fechaproceso)=" & mmes & " " & mcad
Sql = Sql & "and a.status='D' and a.placod=b.placod and a.cia=b.cia and b.status<>'*'"
cn.CursorLocation = adUseClient

Set AdoCabeza.Recordset = cn.Execute(Sql$, 64)
If AdoCabeza.Recordset.RecordCount > 0 Then
   AdoCabeza.Recordset.MoveFirst
   Dgrdcabeza.Enabled = True
Else
   Dgrdcabeza.Enabled = False
End If
Dgrdcabeza.Refresh
Screen.MousePointer = vbDefault
End Sub
Private Sub Procesa_Consultas()
If Lbltipo.Caption = "V" Then Carga_Prov_Vaca
If Lbltipo.Caption = "D" Then Procesa_Devengadas
If Lbltipo.Caption = "G" Then Carga_Prov_Grati
End Sub
Private Sub Carga_Prov_Grati()
Dim mcad As String
If VTipo = "" Or Cmbtipo.Text = "" Then Exit Sub

If Trim(Txtano.Text) = "" Then Exit Sub
If Cmbmes.ListIndex < 0 Then Exit Sub
If VTipo = "99" Or VTipo = "" Then mcad = "" Else mcad = " and a.tipotrab='" & VTipo & "' "
Sql = nombre()
Sql = Sql & "a.placod,a.gratmes AS provmes " _
& "from plaprovgrati a,planillas b " _
& "where a.cia='" & wcia & "' and year(a.fechaproceso)='" & Val(Txtano.Text) & "' and month(a.fechaproceso)='" & Cmbmes.ListIndex + 1 & "' and a.status<>'*' " _
& "and a.placod=b.placod and a.cia=b.cia and b.status<>'*'" & mcad & "order by nombre"

cn.CursorLocation = adUseClient
Set AdoCabeza.Recordset = cn.Execute(Sql$, 64)
If AdoCabeza.Recordset.RecordCount > 0 Then
   AdoCabeza.Recordset.MoveFirst
   Command4.Enabled = False
   CmdVacPerd.Enabled = False
   Sql = "select SUM(gratmes) from plaprovgrati a " _
       & "where a.cia='" & wcia & "' and year(a.fechaproceso)='" & Val(Txtano.Text) & "' and month(a.fechaproceso)='" & Cmbmes.ListIndex + 1 & "' and a.status<>'*'" & mcad
   
   If (fAbrRst(rs, Sql)) Then Lbltotal.Caption = Format(rs(0), "###,###.00")
   rs.Close
Else
   Command4.Enabled = True
   CmdVacPerd.Enabled = True
   Lbltotal.Caption = "0.00"
End If
Dgrdcabeza.Refresh
End Sub
Private Sub Calcula_Provision_Grati()
Dim Dia As Integer
Dim fecproc As String
Dim mtoting As Currency
Dim mBase As Currency
Dim mCadIng As String
Dim mCadProm As String
Dim cont1 As Integer
Dim fpromed As String
Dim mFactProm As Integer
Dim nFields As Integer
Dim mfectope As String
Dim I As Integer
Dim J As Integer
Dim mGratAnt As Currency
Dim mGratMes As Currency
Dim mProvMes As Currency
Dim mMeses As Integer


Dia = Ultimo_Dia(Cmbmes.ListIndex + 1, Val(Txtano.Text))
fecproc = Format(Dia, "00") & "/" & Format(Cmbmes.ListIndex + 1, "00") & "/" & Txtano.Text

Sql = "select * from planillas where cia='" & wcia & "' and tipotrabajador='" & VTipo & "' and fcese is null and status<>'*' order by placod"
If (fAbrRst(rs, Sql)) Then rs.MoveFirst

Screen.MousePointer = vbArrowHourglass

Panelprogress.Caption = "Calculando Provision de Gratificaciones"
Panelprogress.Visible = True
Panelprogress.ZOrder 0
Me.Refresh
Barra.Max = rs.RecordCount
Barra.Value = 0
Sql$ = wInicioTrans
cn.Execute Sql$

Do While Not rs.EOF
   Barra.Value = rs.AbsolutePosition
   mGratAnt = 0
   'Numero de Meses
   If Cmbmes.ListIndex + 1 = 1 Or Cmbmes.ListIndex + 1 = 7 Then
      mGratAnt = 0: mMeses = 0
      If Year(rs!fIngreso) < Val(Txtano.Text) Then
         mMeses = 1
      Else
         If Month(rs!fIngreso) < Cmbmes.ListIndex + 1 Then
            mMeses = 1
         Else
            If Day(rs!fIngreso) = 1 Then mMeses = 1
         End If
      End If
   Else
      Sql = "select gratmes from plaprovgrati where cia='" & rs!cia & "' and month(fechaproceso)=" & Cmbmes.ListIndex & " AND YEAR(FECHAPROCESO)=" & Txtano.Text & " and placod='" & rs!PlaCod & "' and status<>'*'"
      If (fAbrRst(rs2, Sql)) Then
         mGratAnt = rs2(0): mMeses = rs2(1) + 1:
      Else
        mGratAnt = 0
        If Year(rs!fIngreso) < Val(Txtano.Text) Then
           mMeses = 1
           If Month(rs!fIngreso) < Cmbmes.ListIndex + 1 Then
              mMeses = 1
           Else
             If Day(rs!fIngreso) = 1 Then mMeses = 1
           End If
        End If
      End If
   End If
   
   'Remuneraciones
   Sql$ = "select concepto,tipo,factor_horas,sum(importe) as base from plaremunbase a where cia='" & wcia & "' " _
        & "and placod='" & Trim(rs!PlaCod) & "' and concepto<>'03' " & _
        "and status<>'*' group by concepto,factor_horas,tipo order by concepto"
        
   If (fAbrRst(rs2, Sql$)) Then rs2.MoveFirst
   mtoting = 0: mCadIng = "": cont1 = 0
   Do While Not rs2.EOF
      mBase = 0
      For I = 1 To 50
          If I = Val(rs2(0)) Then
             If VTipo = "01" Then
                If rs2!tipo = "04" Then
                   mBase = rs2(3)
                Else
                   mBase = Round(rs2(3) / rs2(2) * mHourMonth, 2)
                End If
             Else
                If rs2!tipo = "01" Then
                   mBase = rs2(3)
                Else
                   mBase = Round(rs2(3) / rs2(2) * mHourMonth, 2)
                End If
             End If
             mCadIng = mCadIng & "" & mBase & "" & ","
             mtoting = mtoting + mBase
             cont1 = cont1 + 1
             Exit For
          ElseIf I > cont1 Then
             mBase = 0
             mCadIng = mCadIng & "" & mBase & "" & ","
             cont1 = cont1 + 1
          End If
      Next
      rs2.MoveNext
   Loop
   
   For I = (cont1 + 1) To 50
       mCadIng = mCadIng & "0,"
   Next
   mFactProm = 0
   '49
   'Promedios
   Sql = "select codinterno,factor from platasaanexo where cia='" & _
   rs!cia & "' and modulo='01' and tipomovimiento='" & _
   VTipo & "' and status<>'*'"
   
   If (fAbrRst(rs2, Sql$)) Then rs2.MoveFirst: mFactProm = rs2(1)
   mCadProm = ""
   nFields = 0
   Do While Not rs2.EOF
      mCadProm = mCadProm & "sum(i" & rs2(0) & " ) as i" & rs2(0) & ","
      nFields = nFields + 1
      rs2.MoveNext
   Loop
   '4=53
   rs2.Close
   fpromed = ""
   If Trim(mCadProm) <> "" Then
      mCadProm = Mid(mCadProm, 1, Len(Trim(mCadProm)) - 1)
      fpromed = Fecha_Promedios(mFactProm, fecproc)
      
      If Val(Mid(fecproc, 4, 2)) = 1 Then
         mfectope = "31/12/" & Format(Val(Mid(fecproc, 7, 4)) - 1, "0000")
      Else
         Dia = Ultimo_Dia(Cmbmes.ListIndex, Val(Txtano.Text))
         mfectope = Format(Dia, "00") & "/" & Format(Cmbmes.ListIndex, "00") & "/" & Mid(fecproc, 7, 4)
      End If
      
      Sql = "SET DATEFORMAT " & Coneccion.FormatFechaSql
      Sql = Sql & " select " & mCadProm & " from plahistorico where cia='" & rs!cia & "' and  placod='" & Trim(rs!PlaCod) & "' and fechaproceso " _
          & "BETWEEN '" & Format(fpromed, FormatFecha) + FormatTimei & "' AND '" & Format(mfectope, FormatFecha) + FormatTimef & "' " _
          & "and proceso='01' and status<>'*'"
          
     If (fAbrRst(rs2, Sql$)) Then rs2.MoveFirst
     mCadProm = ""
        For I = 1 To 50
            mBase = 0
            For J = 0 To nFields - 1
                If Format(I, "00") = Mid(rs2(J).Name, 2, 2) Then
                   If IsNull(rs2(J)) Then mBase = 0 Else mBase = rs2(J) / mFactProm
                   Exit For
                Else
                   mBase = 0
                End If
            Next J
            mtoting = mtoting + mBase
            mCadProm = mCadProm & "" & mBase & "" & ","
        Next I
   End If
   rs2.Close
   
   If mMeses <> 0 Then
      mGratMes = mMeses * mtoting / 6
      mProvMes = mGratMes - mGratAnt
      If mProvMes <> 0 Then
         Sql = "SET DATEFORMAT " & Coneccion.FormatFechaSql & "  "
         Sql = Sql & "insert into plaprovgrati values('" & Trim(wcia) & _
         "','" & Trim(rs!PlaCod) & "','" & Trim(rs!TipoTrabajador) & "'," _
         & "" & mMeses & ",'' ,'' ,'' ," & mCadIng & mCadProm
         Sql = Sql & "" & mtoting & "," & mtoting & "," & mGratAnt & _
         "," & mGratMes & "," & mProvMes & ",'" & Format(fecproc, FormatFecha) & _
         "'," & FechaSys & ",'" & wuser & "','" & Trim(rs!Area) & "','')"
             
         cn.Execute Sql$
         
      End If
   End If
   rs.MoveNext
Loop

Sql$ = wFinTrans
cn.Execute Sql$
Panelprogress.Visible = False
Carga_Prov_Grati
Screen.MousePointer = vbDefault

End Sub
Public Sub Elimina_Prov_Grati()
'If Val(Txtano) < 2007 Then GoTo MENSAJE2
'If Val(Txtano) = 2007 And Cmbmes.ListIndex + 1 > 6 Then

If Txtano.Text < 2013 Then
   MsgBox "En este sistema se pueden trabajar provisiones desde Setiembre del 2013", vbInformation
   Exit Sub
End If
If Txtano.Text = 2013 And Cmbmes.ListIndex + 1 < 9 Then
   MsgBox "En este sistema se pueden trabajar provisiones desde Setiembre del 2013", vbInformation
   Exit Sub
End If

    If VTipo = "" Or VTipo = "99" Then MsgBox "Debe Indicar Tipo de Trabajador", vbInformation, "Provisiones": Exit Sub
    If MsgBox("Desea Eliminar Provision ? ", vbQuestion + vbYesNo + vbDefaultButton1, "Seteo de Usuario") = vbNo Then
        Exit Sub
    Else
        Sql = wInicioTrans
        cn.Execute Sql
    
        Sql = "update plaprovgrati set status='*' where cia='" & wcia & "' and year(fechaproceso)=" & Val(Txtano.Text) & " and month(fechaproceso)=" & Cmbmes.ListIndex + 1 & " and tipotrab='" & VTipo & "' and status<>'*'"
        cn.Execute Sql
    
        Sql = wFinTrans
        cn.Execute Sql
    
        Carga_Prov_Grati
    End If
    Screen.MousePointer = vbDefault
'Else
 '   GoTo MENSAJE2
'End If
'Exit Sub
'MENSAJE2:
 '   MsgBox "No se puede Eliminar la Informacion de la Provision", vbInformation
End Sub
Private Sub Reporte_Provision_Grati_Antes()
Dim MArea As String
Dim wciamae As String
Dim mcad As String
Dim mCadBas As String
Dim mCadProm As String
Dim mFieldBas As Integer
Dim mFieldProm As Integer
Dim MField As Integer
Dim contbase As Integer
Dim mText As String
Dim msum As Integer
Dim sumparc As Integer
Dim tot1 As Currency
Dim tot2 As Currency
Dim tot3 As Currency
Dim tot4 As Currency
Dim tot5 As Currency

Dim I As Integer
Dim x As Integer
If VTipo = "" Or VTipo = "99" Then MsgBox "Debe Indicar Tipo de Trabajador", vbInformation, "Provisiones": Exit Sub
If AdoCabeza.Recordset.RecordCount <= 0 Then Exit Sub

Screen.MousePointer = vbArrowHourglass
mcad = ""
For I = 1 To 50
    mcad = mcad & "sum(i" & Format(I, "00") & "),"
Next
For I = 1 To 50
    mcad = mcad & "sum(p" & Format(I, "00") & "),"
Next
If Trim(mcad) <> "" Then
   mcad = Mid(mcad, 1, Len(Trim(mcad)) - 1)
End If

Panelprogress.Caption = "Generando Reporte de Provision de Gratificaciones"
Panelprogress.Visible = True
Panelprogress.ZOrder 0
Me.Refresh
Barra.Value = 0
Sql = "select " & mcad & " from plaprovgrati " _
    & "where cia='" & wcia & "' and year(fechaproceso)='" & Val(Txtano.Text) & "' and month(fechaproceso)='" & Cmbmes.ListIndex + 1 & "' and tipotrab='" & VTipo & "' and status<>'*' GROUP BY PLACOD ORDER BY PLACOD"
    
If (fAbrRst(rs, Sql)) Then
   mCadBas = ""
   mCadProm = ""
   Barra.Max = 99
   For I = 0 To 99
       Barra.Value = I
       If rs(I) <> 0 Then
          If I <= 49 Then
             mCadBas = mCadBas & Format(I + 1, "00")
          Else
             mCadProm = mCadProm & Format(I - 50 + 1, "00")
          End If
       End If
   Next
End If

If Trim(mCadBas) = "" Then mFieldBas = 0 Else mFieldBas = Len(Trim(mCadBas))
If Trim(mCadProm) = "" Then mFieldProm = 0 Else mFieldProm = Len(Trim(mCadProm))
MArea = ""
wciamae = Determina_Maestro("01044")

Set xlApp1 = CreateObject("Excel.Application")
xlApp1.Workbooks.Add
Set xlApp2 = xlApp1.Application
Set xlBook = xlApp2.Workbooks(1)
Set xlSheet = xlApp2.Worksheets("HOJA1")

xlSheet.Range("A:A").ColumnWidth = 3
xlSheet.Range("B:B").ColumnWidth = 7
xlSheet.Range("C:C").ColumnWidth = 45
xlSheet.Range("E:G").ColumnWidth = 12
xlSheet.Range("E:AZ").NumberFormat = "#,###,##0.00"
'xlSheet.Columns("D:D").NumberFormat = "@"

xlSheet.Cells(1, 2).Value = Cmbcia.Text
xlSheet.Cells(1, 2).Font.Size = 12
xlSheet.Cells(1, 2).Font.Bold = True

xlSheet.Cells(3, 2).Value = "REPORTE DE PROVISION DE GRATIFICACIONES " & Cmbtipo.Text & " - " & Cmbmes.Text & " " & Txtano.Text
xlSheet.Cells(3, 2).Font.Size = 11
xlSheet.Cells(3, 2).Font.Bold = True

xlSheet.Cells(6, 2).Value = "Codigo"
xlSheet.Range(xlSheet.Cells(6, 2), xlSheet.Cells(7, 2)).Merge
xlSheet.Cells(6, 2).VerticalAlignment = xlCenter

xlSheet.Cells(6, 3).Value = "Nombre Trabajador"
xlSheet.Cells(6, 3).VerticalAlignment = xlCenter

xlSheet.Cells(6, 4).Value = "F.Ingreso"
xlSheet.Range("D:D").HorizontalAlignment = xlCenter

xlSheet.Cells(6, 5).Value = "DNI"
xlSheet.Range("E:E").HorizontalAlignment = xlCenter

'--Antes 5
xlSheet.Cells(6, 6).Value = "Jornal"
xlSheet.Cells(7, 6).Value = "Basico"
xlSheet.Cells(6, 6).HorizontalAlignment = xlRight
xlSheet.Cells(7, 6).HorizontalAlignment = xlRight

xlSheet.Cells(6, 7).Value = "Bonific"
xlSheet.Cells(7, 7).Value = "AFP"
xlSheet.Cells(6, 7).HorizontalAlignment = xlRight
xlSheet.Cells(7, 7).HorizontalAlignment = xlRight

xlSheet.Cells(6, 8).Value = "Bonific."
xlSheet.Cells(7, 8).Value = "Costo Vida"
xlSheet.Cells(6, 8).HorizontalAlignment = xlRight
xlSheet.Cells(7, 8).HorizontalAlignment = xlRight


xlSheet.Cells(6, 9).Value = "Bonifica"
xlSheet.Cells(7, 9).Value = "T.Servicio"
xlSheet.Cells(6, 9).HorizontalAlignment = xlRight
xlSheet.Cells(7, 9).HorizontalAlignment = xlRight


xlSheet.Cells(6, 10).Value = "Asignacion"
xlSheet.Cells(7, 10).Value = "Familiar"
xlSheet.Cells(6, 10).HorizontalAlignment = xlRight
xlSheet.Cells(7, 10).HorizontalAlignment = xlRight


xlSheet.Cells(6, 11).Value = "Promedio"
xlSheet.Cells(7, 11).Value = "H.Extras"
xlSheet.Cells(6, 11).HorizontalAlignment = xlRight
xlSheet.Cells(7, 11).HorizontalAlignment = xlRight


xlSheet.Cells(6, 12).Value = "Otros"
xlSheet.Cells(7, 12).Value = "Pagos"
xlSheet.Cells(6, 12).HorizontalAlignment = xlRight
xlSheet.Cells(7, 12).HorizontalAlignment = xlRight

xlSheet.Cells(6, 13).Value = "Aportes"
xlSheet.Cells(7, 13).Value = "Patronales"
xlSheet.Cells(6, 13).HorizontalAlignment = xlRight
xlSheet.Cells(7, 13).HorizontalAlignment = xlRight

xlSheet.Cells(6, 14).Value = "Jornal"
xlSheet.Cells(7, 14).Value = "Remuner."
xlSheet.Cells(6, 14).HorizontalAlignment = xlRight
xlSheet.Cells(7, 14).HorizontalAlignment = xlRight

xlSheet.Cells(6, 15).Value = "Prov."
xlSheet.Cells(7, 15).Value = "Mes Ant."
xlSheet.Cells(6, 15).HorizontalAlignment = xlRight
xlSheet.Cells(7, 15).HorizontalAlignment = xlRight

xlSheet.Cells(6, 16).Value = "Prov."
xlSheet.Cells(7, 16).Value = "Por Pagar"
xlSheet.Cells(6, 16).HorizontalAlignment = xlRight
xlSheet.Cells(7, 16).HorizontalAlignment = xlRight

xlSheet.Cells(6, 17).Value = "Prov."
xlSheet.Cells(7, 17).Value = "Del Mes"
xlSheet.Cells(6, 17).HorizontalAlignment = xlRight
xlSheet.Cells(7, 17).HorizontalAlignment = xlRight

xlSheet.Cells(6, 18).Value = "Afiliados"
xlSheet.Cells(7, 18).Value = "EPS"

xlSheet.Range("R:R").HorizontalAlignment = xlCenter


nFil = 8
sumparc = nFil
Sql = nombre()
Sql = Sql & "a.*,b.fingreso,b.afiliado_eps_serv, b.nro_doc " _
& "from plaprovgrati a,planillas b " _
& "where a.cia='" & wcia & "' and year(a.fechaproceso)='" & Val(Txtano.Text) & "' and month(a.fechaproceso)='" & Cmbmes.ListIndex + 1 & "' and a.tipotrab='" & VTipo & "' and a.status<>'*' " _
& "and a.placod=b.placod and a.cia=b.cia and b.status<>'*' order by b.area,nombre"

If (fAbrRst(rs, Sql)) Then rs.MoveFirst: MArea = rs!Area
If MArea <> "" Then
   Sql$ = "Select descrip from maestros_2 where status<>'*' and cod_maestro2='" & Trim(MArea) & "'"
   Sql$ = Sql$ & wciamae
  
   If (fAbrRst(rs2, Sql)) Then xlSheet.Cells(nFil, 2).Value = rs2!DESCRIP
   xlSheet.Cells(nFil, 2).Font.Bold = True
   nFil = nFil + 2
   sumparc = sumparc + 2
   rs2.Close
End If

nCol = 18

tot1 = 0: tot2 = 0: tot3 = 0: tot4 = 0: tot5 = 0
Barra.Value = 0
Barra.Max = rs.RecordCount
Do While Not rs.EOF
   Barra.Value = rs.AbsolutePosition
   If rs!Area <> MArea Then
      msum = (sumparc - 2) * -1
      nFil = nFil + 1
      x = 1

      For I = nCol - 4 To nCol - 1
         xlSheet.Cells(nFil, I).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
         Select Case x
                Case Is = 1: tot1 = tot1 + xlSheet.Cells(nFil, I).Value
                Case Is = 2: tot2 = tot2 + xlSheet.Cells(nFil, I).Value
                Case Is = 3: tot3 = tot3 + xlSheet.Cells(nFil, I).Value
                Case Is = 4: tot4 = tot4 + xlSheet.Cells(nFil, I).Value
                Case Is = 5: tot5 = tot5 + xlSheet.Cells(nFil, I).Value
         End Select
         x = x + 1
      Next I
      sumparc = 0
      
      nFil = nFil + 2
      sumparc = sumparc + 3
      MArea = rs!Area
      Sql$ = "Select descrip from maestros_2 where status<>'*' and cod_maestro2='" & MArea & "'"
      Sql$ = Sql$ & wciamae
      If (fAbrRst(rs2, Sql)) Then xlSheet.Cells(nFil, 2).Value = rs2!DESCRIP
      
        xlSheet.Cells(nFil, 2).Font.Bold = True
        nFil = nFil + 2
        sumparc = sumparc + 2
        rs2.Close
    End If
    nCol = 2
    xlSheet.Cells(nFil, nCol).Value = CStr(rs!PlaCod)
    nCol = nCol + 1
    xlSheet.Cells(nFil, nCol).Value = CStr(rs!nombre)
    nCol = nCol + 1
    
    'FEC. INGRESO
     xlSheet.Cells(nFil, nCol).Value = Format(Trim(rs("fingreso")), "MM/DD/YYYY")
    'DNI
    nCol = nCol + 1
    xlSheet.Cells(nFil, nCol).Value = "'" + Trim(rs("nro_doc"))
    
    nCol = nCol + 1
    xlSheet.Cells(nFil, nCol).Value = CStr(rs("I01"))
    nCol = nCol + 1
    xlSheet.Cells(nFil, nCol).Value = CStr(rs("I05") + rs("I06"))
    nCol = nCol + 1
    xlSheet.Cells(nFil, nCol).Value = CStr(rs("I07"))
    nCol = nCol + 1
    xlSheet.Cells(nFil, nCol).Value = CStr(rs("I04"))
    nCol = nCol + 1
    xlSheet.Cells(nFil, nCol).Value = CStr(rs("I02"))
    nCol = nCol + 1
    'Horas Extras 'GALLOS lo cumula en P10
    xlSheet.Cells(nFil, nCol).Value = CStr(rs("P10") + rs("P11") + rs("P21") + rs("P24") + rs("P25"))
    nCol = nCol + 1
    'Otros pagos, Reintegros, Bonif.xRend., Bonif.xProd
    xlSheet.Cells(nFil, nCol).Value = CStr(rs("P13") + rs("P16") + rs("P37") + rs("P38"))
    nCol = nCol + 1
    xlSheet.Cells(nFil, nCol).Value = CStr(rs("Aporta"))
    nCol = nCol + 1
    xlSheet.Cells(nFil, nCol).Value = CStr(rs!remtotal)
    nCol = nCol + 1
    xlSheet.Cells(nFil, nCol).Value = CStr(rs!gratmesant)
    nCol = nCol + 1
    xlSheet.Cells(nFil, nCol).Value = CStr(rs!provtotal)
    nCol = nCol + 1
    xlSheet.Cells(nFil, nCol).Value = CStr(rs!gratmes)
    nCol = nCol + 1
    xlSheet.Cells(nFil, nCol).Value = CStr(IIf(rs!afiliado_eps_serv, "SI", ""))
'    nCol = nCol + 1
    
    nFil = nFil + 1
    sumparc = sumparc + 1

   rs.MoveNext
Loop

Panelprogress.Visible = False
msum = (sumparc - 2) * -1
nFil = nFil + 1
x = 1
For I = nCol - 4 To nCol - 1
    xlSheet.Cells(nFil, I).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
    Select Case x
           Case Is = 1: tot1 = tot1 + xlSheet.Cells(nFil, I).Value
           Case Is = 2: tot2 = tot2 + xlSheet.Cells(nFil, I).Value
           Case Is = 3: tot3 = tot3 + xlSheet.Cells(nFil, I).Value
           Case Is = 4: tot4 = tot4 + xlSheet.Cells(nFil, I).Value
           Case Is = 5: tot5 = tot5 + xlSheet.Cells(nFil, I).Value
    End Select
    x = x + 1
Next I
sumparc = 0

nFil = nFil + 1

msum = (nFil) * -1
nFil = nFil + 1
For I = 6 To nCol - 6
   xlSheet.Cells(nFil, I).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
Next I

x = 1
For I = nCol - 4 To nCol - 1
    Select Case x
           Case Is = 1: xlSheet.Cells(nFil, I).Value = tot1
           Case Is = 2: xlSheet.Cells(nFil, I).Value = tot2
           Case Is = 3: xlSheet.Cells(nFil, I).Value = tot3
           Case Is = 4: xlSheet.Cells(nFil, I).Value = tot4
           Case Is = 5: xlSheet.Cells(nFil, I).Value = tot5
    End Select
    
    x = x + 1
Next

xlApp2.Application.ActiveWindow.DisplayGridLines = False
xlSheet.Range("A1:A1").Select
xlApp2.Application.Caption = "PROVISION DE GRATIFICACION"
xlApp2.ActiveWindow.Zoom = 80
xlApp2.Application.Visible = True

If Not xlApp2 Is Nothing Then Set xlApp2 = Nothing
If Not xlApp1 Is Nothing Then Set xlApp1 = Nothing
If Not xlBook Is Nothing Then Set xlBook = Nothing
If Not xlSheet Is Nothing Then Set xlSheet = Nothing

Screen.MousePointer = 0

End Sub

Function FECHA_ACUMULADO(fecha As Variant, fec1 As Variant, fec2 As Variant) As Variant
            
    If Left(fec1, 1) = "-" Then fec1 = Mid(fec1, 2)
    If Left(fec2, 1) = "-" Then fec2 = Mid(fec2, 2)
    Dim ANNO As Variant, ANNO1 As Variant, ANNO2 As Variant
    Dim dia1 As Variant, DIA2 As Variant, Mes1 As Variant
    Dim MES2 As Variant, Mes As Variant, Dia As Variant
    
    If fecha = "0000.00.00" Then
        FECHA_ACUMULADO = fecha
        Exit Function
    End If
    ANNO = Val(Left(fecha, 4))
    ANNO1 = Left(fec1, 1)
    ANNO2 = Val(Left(fec2, 4))
    dia1 = Mid(fec1, 6)
    DIA2 = Val(Mid(fec2, 9))
    Mes1 = Mid(fec1, 3, 2)
    MES2 = Val(Mid(fec2, 6, 2))
    
    ANNO = ANNO1 - ANNO
    Mes = (Val(Mes1) + MES2)
    If (Mes) > 12 Then
        Mes = (Mes) - 12
        ANNO = ANNO + 1
    End If
    
    Mes = Format(Mes, "00")
    Dia = Format(Val(dia1) + Val(DIA2), "00")
    
    FECHA_ACUMULADO = ANNO & "." & Mes & "." & Dia
           
End Function
Private Sub PROVISION_VACACIONES()
Dim sSQL As String
Dim FecFin As String
Dim FecIni As String

FecFin = Format(Format(Ultimo_Dia(Cmbmes.ListIndex + 1, Val(Txtano.Text)), "00") & "/" & Format(Cmbmes.ListIndex + 1, "00") & "/" & Txtano.Text, "DD/MM/YYYY")
FecIni = DateAdd("m", -5, CDate("01/" & Month(FecFin) & "/" & Year(FecFin)))

sSQL = "usp_Pla_Calcula_PRov_Vaca '" & wcia & "','" & VTipo & "','" & FecIni & "','" & FecFin & "','" & wuser & "'"
cn.Execute (sSQL)

End Sub

Private Sub PROVISION_VACACIONES_antes()
Dim sSQL As String
Dim MAXROW As Long
Dim rs As ADODB.Recordset
Dim RX As ADODB.Recordset
Dim EXTRAS As String, PRODUCCION As String, OTROSPAGOS As String
Dim FechaProc As String
Dim FecFin As String, FecIni As String, fecha1 As String
Dim RcdAñoPasado As String, VacTomadas As String, RcdAñoActual As String, RcdAñoTotal As String, ImpProvVacaAnt As String
Dim RcdPerdida As String
Dim RSBUSCAR As ADODB.Recordset
Dim strCodigo As String
Dim MAXCOL As Integer
Dim dblFactor As Currency
Dim MaxColInicial As Integer
Dim MaxColFin As Integer
Dim MaxColTemp As Integer
Dim I As Integer
Dim Campo As String
Dim SQLVAR As String
Dim sSQLI As String
Dim sSQLP As String
Dim sCol As Integer
Dim sAportePatronal As String

Dim cargaimporte As Integer
Dim VacTomadasMes As String
Dim PAÑO_PROCESO As Integer

Dim Factor_EsSalud As Currency
Dim TotalAporta As Currency

Dim nVecesPer As Integer
Dim NroTrans As Integer
'On Error GoTo ErrorTrans
NroTrans = 0


MAXROW = 0

'SETEAMOS LAS FECHAS A TRABAJAR

FecFin = Format(Format(Ultimo_Dia(Cmbmes.ListIndex + 1, Val(Txtano.Text)), "00") & "/" & Format(Cmbmes.ListIndex + 1, "00") & "/" & Txtano.Text, "DD/MM/YYYY")
FecIni = DateAdd("m", -5, CDate("01/" & Month(FecFin) & "/" & Year(FecFin)))

'IMPLEMENTACION GALLOS

nVecesPer = 3
'************************************************************
'LFSA - 08/08/2012 - obtener APORTE PATRONAL POR CIA
sAportePatronal = ""
sSQL = "select AportePatronal from cia WHERE cod_cia='" & wcia & "'"
If (fAbrRst(rs, sSQL)) Then
    Do While Not rs.EOF
        sAportePatronal = rs!AportePatronal
        rs.MoveNext
    Loop
End If
'************************************************************
'--Unif.

sSQL = "select 'I' + campo from estructura_provisiones where tipo='V' AND campo not in"
sSQL = sSQL & " (select codinterno from platasaanexo where cia='" & wcia & "' and tipomovimiento='02' and basecalculo='16')"
sSQL = sSQL & " Union All select 'P'+codinterno from platasaanexo where cia='" & wcia & "' and tipomovimiento='02' and basecalculo='16' AND tipo IN('1','3') AND STATUS!='*'"

Erase ArrReporte

MAXCOL = 6
MaxColInicial = MAXCOL
I = MAXCOL

If (fAbrRst(rs, sSQL)) Then
    Do While Not rs.EOF
        MAXCOL = MAXCOL + 1
        rs.MoveNext
    Loop
    
    'AGREGAMOS LOS CAMPOS QUE FALTAN DE LOS IMPORTES
    MAXCOL = MAXCOL + 6
    
    rs.MoveFirst
    ReDim Preserve ArrReporte(0 To MAXCOL, 0 To MAXROW)
        
    Do While Not rs.EOF
        ArrReporte(I, MAXROW) = rs(0)
        I = I + 1
        rs.MoveNext
    Loop

    MAXROW = MAXROW + 1
End If

    
    'OBTENEMOS LOS CONCEPTOS REMUNERATIVOS A TRABAJAR
    '-- HORAS EXTRAS --
    
    If Month(FecFin) >= 1 And Month(FecFin) < 6 Then
       PAÑO_PROCESO = Val(Year(FecFin)) - 1
    Else
       PAÑO_PROCESO = Val(Year(FecFin))
    End If
    
    sSQL = " SELECT codinterno FROM platasaanexo WHERE tipomovimiento='02' AND status!='*' AND basecalculo='16' AND tipo='1' AND cia='" & wcia & "'"
    EXTRAS = ""
    
    Dim FiltroFecIng As String, TipoExtra As String
    FiltroFecIng = " AND LEFT(CONVERT(VARCHAR,ph.fechaproceso,112),8) >= LEFT(CONVERT(VARCHAR,a.fingreso,112),8)"
    'CALCULO DE PROMEDIOS
    If (fAbrRst(rs, sSQL)) Then
        Do While Not rs.EOF
            Select Case Trim(rs!codinterno)
            Case "10", "11", "21", "24", "25"
                 TipoExtra = "E"
            Case Else
                 TipoExtra = "B"
             End Select
            
            EXTRAS = EXTRAS & "P" & Trim(rs!codinterno) & "=CASE WHEN dbo.fc_cargapagos2('" & wcia & "',a.placod,'" & Val(Month(FecIni)) & "','" & Val(Month(FecFin)) & "','" & PAÑO_PROCESO & "','" & TipoExtra & "')>=" & nVecesPer & " THEN (SELECT SUM(COALESCE(I" & Trim(rs!codinterno) & ",0)) FROM plahistorico ph WHERE ph.cia='" & wcia & "' and ph.status!='*'  and  ph.FECHAPROCESO>='" & FecIni & "' " & _
            "AND ph.FECHAPROCESO<='" & FecFin & "'" & FiltroFecIng & " and ph.placod=a.placod AND PH.PROCESO='01') ELSE 0 END, "
            
            rs.MoveNext
        Loop
    
        rs.MoveFirst
            
        EXTRAS = Mid(EXTRAS, 1, Len(Trim(EXTRAS)) - 1)
        rs.Close
    Else
        EXTRAS = "EXTRAS=0"
    End If
    
    PRODUCCION = ""
    PRODUCCION = "PRODUCCION=0"
    
    
    '-- HORAS OTROS PAGOS --
    '--'Son considerados en CALCULO DE PROMEDIOS
    
'    OTROSPAGOS = ""
    sSQL = " SELECT codinterno FROM platasaanexo WHERE tipomovimiento='02' AND status!='*' AND basecalculo='16' AND tipo='3' AND CIA='" & wcia & "'"

    If (fAbrRst(rs, sSQL)) Then
        Do While Not rs.EOF
            Select Case Trim(rs!codinterno)
            Case "10", "11", "21", "24", "25"
                 TipoExtra = "E"
            Case Else
                 TipoExtra = "B"
             End Select
            
            OTROSPAGOS = OTROSPAGOS & "P" & Trim(rs!codinterno) & "=CASE WHEN dbo.fc_cargapagos2('" & wcia & "',a.placod,'" & Val(Month(FecIni)) & "','" & Val(Month(FecFin)) & "','" & PAÑO_PROCESO & "','" & TipoExtra & "')>=" & nVecesPer & " THEN (SELECT SUM(COALESCE(I" & Trim(rs!codinterno) & ",0)) FROM plahistorico ph WHERE ph.cia='" & wcia & "' and ph.status!='*'  and  ph.FECHAPROCESO>='" & FecIni & "' " & _
            "AND ph.FECHAPROCESO<='" & FecFin & "'" & FiltroFecIng & " and ph.placod=a.placod AND PH.PROCESO='01') ELSE 0 END, "
            
            rs.MoveNext
        Loop
    
        rs.MoveFirst
            
        OTROSPAGOS = Mid(OTROSPAGOS, 1, Len(Trim(OTROSPAGOS)) - 1)
        rs.Close
    Else
        OTROSPAGOS = "OTROSPAGOS=0"
    End If

sSQL = " SELECT FACTOR FROM platasaanexo WHERE cia='" & wcia & "' and tipomovimiento='02' AND status!='*' AND basecalculo='16' AND tipo='3'"
If (fAbrRst(rs, sSQL)) Then
    dblFactor = rs(0)
End If

'*******************************************************************
'FECHA DE MODIFICACION: 03/01/2008
'MODIFICADO POR : RICARDO HINOSTROZA
'MOTIVO: LA FECHA DE CONSULTA SE GENERABA INCORRECTAMENTE
'SE CAMBIO :Format(Cmbmes.ListIndex + 2, "00")  POR VARIABLE E INSTRUCCION IF PARA EL CONTROL DEL FLUJO
'*******************************************************************

Dim mmes As String
Dim MAÑO As String
 If Format(Cmbmes.ListIndex + 2, "00") = "13" Then
    mmes = "01"
    MAÑO = Str(Val(Txtano.Text) + 1)
 Else
    mmes = Format(Cmbmes.ListIndex + 2, "00")
    MAÑO = Txtano.Text
 End If
 
Dim Fecha_CI As Date
Dim FechaCF As Date
Fecha_CI = DateSerial(Txtano.Text, (Cmbmes.ListIndex + 1) + 0, 1)
FechaCF = DateSerial(Txtano.Text, (Cmbmes.ListIndex + 1) + 1, 0)
 
sSQL = "SET DATEFORMAT DMY SELECT"

'PARA GALLOS SE INCLUYE AL TRABAJADOR AUN SI SU FECHA DE INGRESO ES = AL ULTIMO DIA DEL MES EN PROCESO.

sSQL = sSQL & " a.placod,LTRIM(RTRIM(a.ap_pat))+' '+LTRIM(RTRIM(a.ap_mat))+' '+LTRIM(RTRIM(a.ap_cas))+' '+LTRIM(RTRIM(a.nom_1))+' '+LTRIM(RTRIM(a.nom_2)),"
sSQL = sSQL & " fingreso , fcese, b.placod, b.codinterno, b.descripcion, b.importe,a.area, a.tipotrabajador,b.TipoRem,b.factor_horas,"
sSQL = sSQL & " " & EXTRAS & ", " & PRODUCCION & ", " & OTROSPAGOS
sSQL = sSQL & " FROM planillas a INNER JOIN ("
sSQL = sSQL & " SELECT prb.PLACOD , pc.codinterno, pc.descripcion, prb.importe,prb.tipo as TipoRem,factor_horas FROM plaremunbase prb INNER JOIN placonstante pc ON"
sSQL = sSQL & " (pc.cia='" & wcia & "' AND prb.cia=pc.cia and pc.tipomovimiento='02' and pc.status!='*' and pc.codinterno=prb.concepto) WHERE"
sSQL = sSQL & " prb.STATUS!='*' ) B ON (b.placod=a.placod) LEFT OUTER JOIN plahistorico ph ON (ph.cia='" & wcia & "' and ph.status!='*' "
sSQL = sSQL & " and  ph.FECHAPROCESO>='" & FecIni & "' AND ph.FECHAPROCESO<='" & FecFin & "' and ph.placod=a.placod)"
sSQL = sSQL & " WHERE a.cat_trab <> '04' and a.status!='*'  and a.tipotrabajador='" & VTipo & "' and a.cia='" & wcia & "' AND a.fingreso <= '" & FechaCF & "' and (fcese >= '" & FechaCF & "' or fcese is null ) "
sSQL = sSQL & " GROUP BY a.PLACOD , a.ap_pat, a.ap_mat, a.ap_cas, a.nom_1, a.nom_2, fingreso, fcese, b.PLACOD, b.codinterno, b.descripcion, b.importe,a.area,a.tipotrabajador,b.TipoRem,b.factor_horas"


fecha1 = "01/01/" & Txtano.Text
If (fAbrRst(rs, sSQL)) Then
    DoEvents
    Panelprogress.Caption = "Generando Provisión ..."
    cn.BeginTrans
    NroTrans = 1
    Barra.Min = 0
    Barra.Max = rs.RecordCount
    Set RSBUSCAR = rs.Clone
       
    
Do While Not rs.EOF
    Barra.Value = rs.AbsolutePosition
    If strCodigo <> Trim(rs!PlaCod) Then
        ReDim Preserve ArrReporte(0 To MAXCOL, 0 To MAXROW)
        strCodigo = Trim(rs!PlaCod)
        
        'If Trim(rs!PlaCod) = "RE007" Then Stop
        
        ObtenerFechas rs!PlaCod, rs!fIngreso, fecha1, FecFin, RcdAñoPasado, VacTomadas, RcdAñoActual, RcdPerdida, RcdAñoTotal, ImpProvVacaAnt
        
        ArrReporte(COL_CODIGO, MAXROW) = Trim(rs!PlaCod)
        ArrReporte(COL_NOMBRE, MAXROW) = Trim(rs(1))
        ArrReporte(COL_RCDAÑOANTERIOR, MAXROW) = RcdAñoPasado
        ArrReporte(COL_VACACTOMADAS, MAXROW) = VacTomadas
        ArrReporte(COL_RCDAÑOACTUAL, MAXROW) = RcdAñoActual
        ArrReporte(COL_RCDACUMULADO, MAXROW) = RcdAñoTotal

        'LLENADO DE LOS CAMPOS SEGUN TABLA ESTRUCTURA_PROVISIONES
        For I = MaxColInicial To MAXCOL - 7
            Debug.Print Trim(ArrReporte(I, 0))
            If Left(Trim(ArrReporte(I, 0)), 1) = "I" Then
                RSBUSCAR.Filter = "placod='" & Trim(rs!PlaCod) & "' and codinterno='" & Right(Trim(ArrReporte(I, 0)), 2) & "'"
                If Not RSBUSCAR.EOF Then
                    If rs!TipoTrabajador = "01" Then
                        If RSBUSCAR!TipoRem = "01" Then 'DIARIO
                            ArrReporte(I, MAXROW) = RSBUSCAR!importe * 30
                        Else    'MENSUAL
                            ArrReporte(I, MAXROW) = RSBUSCAR!importe
                        End If
                    Else
                        ArrReporte(I, MAXROW) = Round(RSBUSCAR!importe / (RSBUSCAR!FACTOR_HORAS / hORAS_X_DIA), 2)
                    End If
                Else
                    ArrReporte(I, MAXROW) = 0
                End If
                RSBUSCAR.Filter = ""
            ElseIf Left(Trim(ArrReporte(I, 0)), 1) = "P" Then
                
               If rs.Fields(Trim(ArrReporte(I, 0))) > 0 Then
                    If rs!TipoTrabajador = "01" Then
                        ArrReporte(I, MAXROW) = rs.Fields(Trim(ArrReporte(I, 0))) / 6
                    Else
                        ArrReporte(I, MAXROW) = Round(rs.Fields(Trim(ArrReporte(I, 0))) / 180, 2)
                    End If
                Else
                    ArrReporte(I, MAXROW) = 0
                End If
            End If
        Next

        'SUMAMOS TODOS LOS CAMPOS DE LOS IMPORTES
        
        
        TotalAporta = 0
        Factor_EsSalud = Trae_Porc_EsSalud_EPS(Trim(rs!PlaCod), Trim(rs!TipoTrabajador))
                
        For MaxColTemp = MaxColInicial To MAXCOL - 7  ' - APORTE PATRONAL
            TotalAporta = TotalAporta + ArrReporte(MaxColTemp, MAXROW)
        Next MaxColTemp
        
        'LFSA - 08/08/2012 VALIDACION APORTE PATRONAL
        If sAportePatronal = "S" Then
            ArrReporte(I, MAXROW) = TotalAporta * (1 + Factor_EsSalud * 0.01)
        Else
            ArrReporte(I, MAXROW) = TotalAporta
        End If
        I = I + 1

        'IMPORTE DE LAS VACACIONES ANTERIOR
        ArrReporte(I, MAXROW) = ImpProvVacaAnt
        I = I + 1
        
        VacTomadasMes = " 0. 0. 0."
        
        '*************************************************
        '
        ' FEC MODIFICACION : 03/01/2008
        ' SE REEMPLAZA :
        ' SQLVAR = " SELECT count(*) FROM plahistorico WHERE proceso='02' and cia='" & wcia & "' and status!='*' and placod='" & ArrReporte(COL_CODIGO, MAXROW) & "' AND year(fecperiodovaca)=" & Me.Txtano.Text & " and month(fecperiodovaca)= " & Cmbmes.ListIndex + 1
        ' POR :
        ' SQLVAR = " SELECT count(*) FROM plahistorico WHERE proceso='02' and cia='" & wcia & "' and status!='*' and placod='" & ArrReporte(COL_CODIGO, MAXROW) & "' AND year(fecperiodovaca)=" & Me.Txtano.Text & " and month(fecperiodovaca)= " & Cmbmes.ListIndex
        ' MODIFICADO POR RICARDO HINOSTROZA
        '
        '*************************************************
        '
        ' FECHA DE MODIFICACION : 23/01/2008
        ' SE MODIFICA MONTO DE VACACIONES PAGADAS
        '(LO PROVISIONADO ES REEMPLAZADO POR LO REAL EN PLANILLA
        ' MODIFICADO POR RICARDO HINOSOTROZA
        '
        '*************************************************
        
        Dim MONTO_VACACIONESTOMADAS As Double
        MONTO_VACACIONESTOMADAS = 0
        
        SQLVAR = " SELECT isnull(count(placod),0),isnull(sum(totaling),0)FROM plahistorico WHERE proceso='02' and cia='" & wcia & "' and status!='*' and placod='" & ArrReporte(COL_CODIGO, MAXROW) & "' AND year(FECHAPROCESO)=" & Me.Txtano.Text & " and month(FECHAPROCESO)= " & Cmbmes.ListIndex + 1
        
        If (fAbrRst(RX, SQLVAR)) Then
            cargaimporte = RX(0)
            MONTO_VACACIONESTOMADAS = RX(1)
            VacTomadasMes = " " & Trim(RX(0)) & ". 0. 0"
            RX.Close
        End If
        
        'IMPORTE DE LAS VACACIONES TOMADAS
        
        'ArrReporte(I, MAXROW) = CalculaImpVaca(VacTomadasMes, ArrReporte(MaxColTemp, MAXROW), rs!TipoTrabajador, cargaimporte)
        
        ArrReporte(I, MAXROW) = MONTO_VACACIONESTOMADAS
        I = I + 1

        'IMPORTE DE LAS VACACIONES X PAGAR

        ArrReporte(I, MAXROW) = CalculaImpVaca(RcdAñoTotal, ArrReporte(MaxColTemp, MAXROW), rs!TipoTrabajador)
        I = I + 1
        'IMPORTE DE LAS PROVISIONES DEL MES
        
        
        If MONTO_VACACIONESTOMADAS > 0 Then
            If ArrReporte(MaxColTemp + 2, MAXROW) <> 0 Then
               ArrReporte(I, MAXROW) = ArrReporte(MaxColTemp + 3, MAXROW) - (ArrReporte(MaxColTemp + 1, MAXROW) - ArrReporte(MaxColTemp + 2, MAXROW))
            Else
               ArrReporte(I, MAXROW) = ArrReporte(MaxColTemp + 3, MAXROW) - (ArrReporte(MaxColTemp + 1, MAXROW) - ArrReporte(MaxColTemp + 2, MAXROW))
            End If
        Else
            ArrReporte(I, MAXROW) = ArrReporte(MaxColTemp + 3, MAXROW) - (ArrReporte(MaxColTemp + 1, MAXROW) - ArrReporte(MaxColTemp + 2, MAXROW))
        End If
        
        I = I + 1
        'AREA DEL PERSONAL
        ArrReporte(I, MAXROW) = Trim(rs!Area)
        
        '*******************************************************************
        'LFSA - 08/08/2012
        'APORTE PATRONAL SETEADO POR EMPRESA
        I = I + 1
        If sAportePatronal = "S" Then
            ArrReporte(I, MAXROW) = TotalAporta * Factor_EsSalud * 0.01
        Else
            ArrReporte(I, MAXROW) = 0
        End If
        '*******************************************************************
        
        MAXROW = MAXROW + 1
    
        'RECORREMOS LOS CAMPOS DE INGRESO DE LA TABLA X(
        sSQLI = ""
        For I = 1 To 50
            Campo = "I" & Format(I, "00")
            sCol = BuscaColumna(Campo, MAXCOL)
            If sCol > 0 Then
                sSQLI = sSQLI & ArrReporte(sCol, MAXROW - 1) & ","
            Else
                sSQLI = sSQLI & "0,"
            End If
        Next
        
        'RECORREMOS LOS CAMPOS DE PROMEDIO DE LA TABLA X(
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

        sSQL = "INSERT plaprovvaca VALUES ('" & wcia & "','" & ArrReporte(COL_CODIGO, MAXROW - 1) & "','" & Format(Cmbtipo.ItemData(Cmbtipo.ListIndex), "00") & "','" & ArrReporte(COL_RCDAÑOANTERIOR, MAXROW - 1) & "',"
        sSQL = sSQL & "'" & ArrReporte(COL_VACACTOMADAS, MAXROW - 1) & "','" & ArrReporte(COL_RCDAÑOACTUAL, MAXROW - 1) & "','" & ArrReporte(COL_RCDACUMULADO, MAXROW - 1) & "'," & sSQLI & sSQLP & ArrReporte(MaxColTemp, MAXROW - 1) & ","
        sSQL = sSQL & ArrReporte(MaxColTemp + 1, MAXROW - 1) & "," & ArrReporte(MaxColTemp + 2, MAXROW - 1) & "," & ArrReporte(MaxColTemp + 3, MAXROW - 1) & "," & ArrReporte(MaxColTemp + 4, MAXROW - 1) & ",'" & FecFin & "',GETDATE(),'" & wuser & "','" & ArrReporte(MaxColTemp + 5, MAXROW - 1) & "',' ','" & Format(rs!fIngreso, "DD/MM/YYYY") & "'," & ArrReporte(MaxColTemp + 6, MAXROW - 1) & ","
        sSQL = sSQL & "'" & RcdPerdida & "',0)"

        cn.Execute (sSQL)
    End If
    rs.MoveNext
Loop
    DoEvents
    Panelprogress.Caption = "Completo ..."
    
    cn.CommitTrans
End If

Exit Sub

ErrorTrans:
    If NroTrans = 1 Then
        cn.RollbackTrans
    End If
    
    MsgBox Err.Description, vbCritical, Me.Caption

End Sub
Private Sub PROVISION_VACACIONES_2011()
'DESCONTINUADA

Dim sSQL As String
Dim MAXROW As Long
Dim rs As ADODB.Recordset
Dim RX As ADODB.Recordset
Dim EXTRAS As String, PRODUCCION As String, OTROSPAGOS As String
Dim FechaProc As String
Dim FecFin As String, FecIni As String, fecha1 As String
Dim RcdAñoPasado As String, VacTomadas As String, RcdAñoActual As String, RcdAñoTotal As String, ImpProvVacaAnt As String
Dim RSBUSCAR As ADODB.Recordset
Dim strCodigo As String
Dim MAXCOL As Integer
Dim dblFactor As Currency
Dim MaxColInicial As Integer
Dim MaxColFin As Integer
Dim MaxColTemp As Integer
Dim I As Integer
Dim Campo As String
Dim SQLVAR As String
Dim sSQLI As String
Dim sSQLP As String
Dim sCol As Integer
Dim cargaimporte As Integer
Dim VacTomadasMes As String
Dim PAÑO_PROCESO As Integer

MAXROW = 0

' SETEAMOS LAS FECHAS A TRABAJAR
If Cmbmes.ListIndex + 1 = 12 Then
    FechaProc = "01/01/" & Txtano.Text + 1
Else
    FechaProc = "01/" & Format(Cmbmes.ListIndex + 2, "00") & "/" & Txtano.Text
End If

FecFin = DateAdd("d", -1, CDate(FechaProc))
'**************codigo modificado giovanni 07092007*****************************
'FecIni = DateAdd("m", -6, CDate("01/" & Month(FecFin) & "/" & Year(FecFin)))

If wGrupoPla <> "01" Then
    FecIni = DateAdd("m", -5, CDate("01/" & Month(FecFin) & "/" & Year(FecFin)))
Else
    FecIni = "01/01/" & Txtano.Text
End If
'*****************************************************************************

If wGrupoPla <> "01" Then
    sSQL = "select 'I' + campo from estructura_provisiones where tipo='V' AND campo not in"
    sSQL = sSQL & " (select codinterno from platasaanexo where cia='" & wcia & "' and tipomovimiento='02' and basecalculo='16')"
    sSQL = sSQL & " Union All select 'P'+codinterno from platasaanexo where cia='" & wcia & "' and tipomovimiento='02' and basecalculo='16' AND tipo='1' AND STATUS!='*'"
Else
    sSQL = "select 'I' + campo from estructura_provisiones where tipo='V' AND campo not in"
    sSQL = sSQL & " (select codinterno from platasaanexo where cia='" & wcia & "' and tipomovimiento='02' and basecalculo='16')"
End If


Erase ArrReporte

MAXCOL = 6
MaxColInicial = MAXCOL
I = MAXCOL

If (fAbrRst(rs, sSQL)) Then
    Do While Not rs.EOF
        MAXCOL = MAXCOL + 1
        rs.MoveNext
    Loop
    
    'AGREGAMOS LOS CAMPOS QUE FALTAN DE LOS IMPORTES
    MAXCOL = MAXCOL + 5
    
    rs.MoveFirst
    ReDim Preserve ArrReporte(0 To MAXCOL, 0 To MAXROW)
        
    Do While Not rs.EOF
        ArrReporte(I, MAXROW) = rs(0)
        I = I + 1
        rs.MoveNext
    Loop

    MAXROW = MAXROW + 1
End If


If wGrupoPla <> "01" Then
    
    
    'OBTENEMOS LOS CONCEPTOS REMUNERATIVOS A TRABAJAR
    '-- HORAS EXTRAS --
    
    If Month(FecFin) >= 1 And Month(FecFin) < 6 Then
       PAÑO_PROCESO = Val(Year(FecFin)) - 1
    Else
       PAÑO_PROCESO = Val(Year(FecFin))
    End If
    
    
    
    sSQL = " SELECT codinterno FROM platasaanexo WHERE tipomovimiento='02' AND status!='*' AND basecalculo='16' AND tipo='1' AND cia='" & wcia & "'"
    EXTRAS = ""
    If (fAbrRst(rs, sSQL)) Then
        Do While Not rs.EOF
            EXTRAS = EXTRAS & "P" & Trim(rs!codinterno) & "=CASE WHEN dbo.fc_cargapagos2('" & wcia & "',a.placod,'" & Val(Month(FecIni)) & "','" & Val(Month(FecFin)) & "','" & PAÑO_PROCESO & "')=3 THEN (SELECT SUM(COALESCE(I" & Trim(rs!codinterno) & ",0)) FROM plahistorico ph WHERE ph.cia='" & wcia & "' and ph.status!='*'  and  ph.FECHAPROCESO>='" & FecIni & "' " & _
            "AND ph.FECHAPROCESO<='" & FecFin & "' and ph.placod=a.placod AND PH.PROCESO='01') ELSE 0 END, "
            
            rs.MoveNext
        Loop
    
        rs.MoveFirst
            
        EXTRAS = Mid(EXTRAS, 1, Len(Trim(EXTRAS)) - 1)
        rs.Close
    Else
        EXTRAS = "EXTRAS=0"
    End If
    
    '-- HORAS PRODUCCION --
    PRODUCCION = ""
    sSQL = " SELECT codinterno FROM platasaanexo WHERE tipomovimiento='02' AND status!='*' AND basecalculo='16' AND tipo='2' AND CIA='" & wcia & "'"
    If (fAbrRst(rs, sSQL)) Then
        Do While Not rs.EOF
            PRODUCCION = PRODUCCION & "P" & Trim(rs!codinterno) & "=(SELECT SUM(COALESCE(I" & Trim(rs!codinterno) & ",0)) FROM plahistorico ph WHERE ph.cia='" & wcia & "' and ph.status!='*'  and  ph.FECHAPROCESO>='" & FecIni & "' " & _
                    "AND ph.FECHAPROCESO<='" & FecFin & "' and ph.placod=a.placod AND PH.PROCESO='01'),"
            
            rs.MoveNext
        Loop
        rs.MoveFirst
        PRODUCCION = Mid(PRODUCCION, 1, Len(Trim(PRODUCCION)) - 1)
        rs.Close
    Else
        PRODUCCION = "PRODUCCION=0"
    End If
    
    '-- HORAS OTROS PAGOS --
    OTROSPAGOS = ""
    sSQL = " SELECT codinterno FROM platasaanexo WHERE tipomovimiento='02' AND status!='*' AND basecalculo='16' AND tipo='3' AND CIA='" & wcia & "'"
    
    If (fAbrRst(rs, sSQL)) Then
        Do While Not rs.EOF
            OTROSPAGOS = OTROSPAGOS & "P" & Trim(rs!codinterno) & "=CASE WHEN dbo.fc_cargapagos('" & wcia & "',a.placod,'" & FecIni & "','" & FecFin & "','" & "I" & Trim(rs!codinterno) & "')=1 THEN (SELECT SUM(COALESCE(I" & Trim(rs!codinterno) & ",0)) FROM plahistorico ph WHERE ph.cia='" & wcia & "' and ph.status!='*'  and  ph.FECHAPROCESO>='" & FecIni & "' " & _
                    "AND ph.FECHAPROCESO<='" & FecFin & "' and ph.placod=a.placod AND PH.PROCESO='01') ELSE 0 END ,"
                    
            rs.MoveNext
        Loop
        rs.MoveFirst
        OTROSPAGOS = Mid(OTROSPAGOS, 1, Len(Trim(OTROSPAGOS)) - 1)
        rs.Close
    Else
        OTROSPAGOS = "OTROSPAGOS=0"
    End If

Else

    'IMPLEMENTACION GALLOS
    'CONCEPTO NO SE CONSIDERADOS
    
    EXTRAS = "EXTRAS=0"
    PRODUCCION = "PRODUCCION=0"
    OTROSPAGOS = "OTROSPAGOS=0"
End If


sSQL = " SELECT FACTOR FROM platasaanexo WHERE cia='" & wcia & "' and tipomovimiento='02' AND status!='*' AND basecalculo='16' AND tipo='3'"
If (fAbrRst(rs, sSQL)) Then
    dblFactor = rs(0)
End If
'******************************************************************************************************************************

'************codigo agregado giovanni 05092007**********************
Select Case Format(Cmbmes.ListIndex + 1, "00")
    Case "02": s_Dia_Proceso_Prov = "28"
    Case Else: s_Dia_Proceso_Prov = "30"
End Select
'*******************************************************************

'*******************************************************************
'FECHA DE MODIFICACION: 03/01/2008
'MODIFICADO POR : RICARDO HINOSTROZA
'MOTIVO: LA FECHA DE CONSULTA SE GENERABA INCORRECTAMENTE
'SE CAMBIO :Format(Cmbmes.ListIndex + 2, "00")  POR VARIABLE E INSTRUCCION IF PARA EL CONTROL DEL FLUJO
'*******************************************************************

Dim mmes As String
Dim MAÑO As String
 If Format(Cmbmes.ListIndex + 2, "00") = "13" Then
    mmes = "01"
    MAÑO = Str(Val(Txtano.Text) + 1)
 Else
    mmes = Format(Cmbmes.ListIndex + 2, "00")
    MAÑO = Txtano.Text
 End If
 
Dim Fecha_CI As Date
Dim FechaCF As Date
Fecha_CI = DateSerial(Txtano.Text, (Cmbmes.ListIndex + 1) + 0, 1)
FechaCF = DateSerial(Txtano.Text, (Cmbmes.ListIndex + 1) + 1, 0)
 
sSQL = "SET DATEFORMAT DMY SELECT"
If wGrupoPla <> "01" Then
    sSQL = sSQL & " a.placod,LTRIM(RTRIM(a.ap_pat))+' '+LTRIM(RTRIM(a.ap_mat))+' '+LTRIM(RTRIM(a.ap_cas))+' '+LTRIM(RTRIM(a.nom_1))+' '+LTRIM(RTRIM(a.nom_2)),"
    sSQL = sSQL & " fingreso , fcese, b.placod, b.codinterno, b.descripcion, b.importe,a.area, a.tipotrabajador,b.factor_horas,"
    sSQL = sSQL & " " & EXTRAS & ", " & PRODUCCION & ", " & OTROSPAGOS
    sSQL = sSQL & " FROM planillas a INNER JOIN ("
    sSQL = sSQL & " SELECT prb.PLACOD , pc.codinterno, pc.descripcion, prb.importe,factor_horas FROM plaremunbase prb INNER JOIN placonstante pc ON"
    sSQL = sSQL & " (pc.cia='" & wcia & "' AND prb.cia=pc.cia and pc.tipomovimiento='02' and pc.status!='*' and pc.codinterno=prb.concepto) WHERE"
    sSQL = sSQL & " prb.STATUS!='*' ) B ON (b.placod=a.placod) LEFT OUTER JOIN plahistorico ph ON (ph.cia='" & wcia & "' and ph.status!='*' "
    sSQL = sSQL & " and  ph.FECHAPROCESO>='" & FecIni & "' AND ph.FECHAPROCESO<='" & FecFin & "' and ph.placod=a.placod)"
    'sSQL = sSQL & " WHERE a.status!='*'  and a.tipotrabajador='" & VTipo & "' and a.cia='" & wcia & "' AND a.fingreso < '" & s_Dia_Proceso_Prov & "/" & Format(Cmbmes.ListIndex + 1, "00") & "/" & Txtano.Text & "' and (fcese >= '" & "01/" & mmes & "/" & Trim(MAÑO) & "' or fcese is null ) "
    sSQL = sSQL & " WHERE a.status!='*'  and a.tipotrabajador='" & VTipo & "' and a.cia='" & wcia & "' AND a.fingreso < '" & FechaCF & "' and (fcese >= '" & FechaCF & "' AND DATEDIFF(DAY, A.FINGRESO, A.FCESE) >= 30 or fcese is null ) "
    sSQL = sSQL & " GROUP BY a.PLACOD , a.ap_pat, a.ap_mat, a.ap_cas, a.nom_1, a.nom_2, fingreso, fcese, b.PLACOD, b.codinterno, b.descripcion, b.importe,a.area,a.tipotrabajador,b.factor_horas"
Else
    'PARA GALLOS SE INCLUYE AL TRABAJADOR AUN SI SU FECHA DE INGRESO ES = AL ULTIMO DIA DEL MES EN PROCESO.
    
    sSQL = sSQL & " a.placod,LTRIM(RTRIM(a.ap_pat))+' '+LTRIM(RTRIM(a.ap_mat))+' '+LTRIM(RTRIM(a.ap_cas))+' '+LTRIM(RTRIM(a.nom_1))+' '+LTRIM(RTRIM(a.nom_2)),"
    sSQL = sSQL & " fingreso , fcese, b.placod, b.codinterno, b.descripcion, b.importe,a.area, a.tipotrabajador,b.factor_horas,"
    sSQL = sSQL & " " & EXTRAS & ", " & PRODUCCION & ", " & OTROSPAGOS
    sSQL = sSQL & " FROM planillas a INNER JOIN ("
    sSQL = sSQL & " SELECT prb.PLACOD , pc.codinterno, pc.descripcion, prb.importe,factor_horas FROM plaremunbase prb INNER JOIN placonstante pc ON"
    sSQL = sSQL & " (pc.cia='" & wcia & "' AND prb.cia=pc.cia and pc.tipomovimiento='02' and pc.status!='*' and pc.codinterno=prb.concepto) WHERE"
    sSQL = sSQL & " prb.STATUS!='*' ) B ON (b.placod=a.placod) LEFT OUTER JOIN plahistorico ph ON (ph.cia='" & wcia & "' and ph.status!='*' "
    sSQL = sSQL & " and  ph.FECHAPROCESO>='" & FecIni & "' AND ph.FECHAPROCESO<='" & FecFin & "' and ph.placod=a.placod)"
    sSQL = sSQL & " WHERE a.status!='*'  and a.tipotrabajador='" & VTipo & "' and a.cia='" & wcia & "' AND a.fingreso <= '" & FechaCF & "' and (fcese >= '" & FechaCF & "' or fcese is null ) "
    sSQL = sSQL & " GROUP BY a.PLACOD , a.ap_pat, a.ap_mat, a.ap_cas, a.nom_1, a.nom_2, fingreso, fcese, b.PLACOD, b.codinterno, b.descripcion, b.importe,a.area,a.tipotrabajador,b.factor_horas"
End If

fecha1 = "01/01/" & Txtano.Text
If (fAbrRst(rs, sSQL)) Then
    Panelprogress.Caption = "Generando Provisión ..."
    Barra.Min = 0
    Barra.Max = rs.RecordCount
    Set RSBUSCAR = rs.Clone

Do While Not rs.EOF
    Barra.Value = rs.AbsolutePosition
    If strCodigo <> Trim(rs!PlaCod) Then
        ReDim Preserve ArrReporte(0 To MAXCOL, 0 To MAXROW)
        strCodigo = Trim(rs!PlaCod)
                
        'ObtenerFechas rs!PlaCod, rs!fIngreso, fecha1, FecFin, RcdAñoPasado, VacTomadas, RcdAñoActual, RcdAñoTotal, ImpProvVacaAnt
        
        ArrReporte(COL_CODIGO, MAXROW) = Trim(rs!PlaCod)
        ArrReporte(COL_NOMBRE, MAXROW) = Trim(rs(1))
        ArrReporte(COL_RCDAÑOANTERIOR, MAXROW) = RcdAñoPasado
        ArrReporte(COL_VACACTOMADAS, MAXROW) = VacTomadas
        ArrReporte(COL_RCDAÑOACTUAL, MAXROW) = RcdAñoActual
        ArrReporte(COL_RCDACUMULADO, MAXROW) = RcdAñoTotal

        'LLENADO DE LOS CAMPOS SEGUN TABLA ESTRUCTURA_PROVISIONES
        For I = MaxColInicial To MAXCOL - 6
            If Left(Trim(ArrReporte(I, 0)), 1) = "I" Then
                RSBUSCAR.Filter = "placod='" & Trim(rs!PlaCod) & "' and codinterno='" & Right(Trim(ArrReporte(I, 0)), 2) & "'"
                If Not RSBUSCAR.EOF Then
                    If rs!TipoTrabajador = "01" Then
                        ArrReporte(I, MAXROW) = RSBUSCAR!importe
                    Else
                        ArrReporte(I, MAXROW) = Round(RSBUSCAR!importe / (RSBUSCAR!FACTOR_HORAS / hORAS_X_DIA), 2)
                    End If
                Else
                    ArrReporte(I, MAXROW) = 0
                End If
                RSBUSCAR.Filter = ""
            ElseIf Left(Trim(ArrReporte(I, 0)), 1) = "P" Then
                If rs.Fields(Trim(ArrReporte(I, 0))) > 0 Then
                    If rs!TipoTrabajador = "01" Then
                        ArrReporte(I, MAXROW) = rs.Fields(Trim(ArrReporte(I, 0))) / dblFactor
                    Else
                        'ArrReporte(i, MAXROW) = Round(rs.Fields(Trim(ArrReporte(i, 0))) / 240, 2)
                        ArrReporte(I, MAXROW) = Round(rs.Fields(Trim(ArrReporte(I, 0))) / 180, 2)
                    End If
                Else
                    ArrReporte(I, MAXROW) = 0
                End If
            End If
        Next

        'SUMAMOS TODOS LOS CAMPOS DE LOS IMPORTES
        For MaxColTemp = MaxColInicial To MAXCOL - 6
            ArrReporte(I, MAXROW) = ArrReporte(I, MAXROW) + ArrReporte(MaxColTemp, MAXROW)
        Next MaxColTemp
        I = I + 1

        'IMPORTE DE LAS VACACIONES ANTERIOR
        ArrReporte(I, MAXROW) = ImpProvVacaAnt
        I = I + 1
        
        VacTomadasMes = " 0. 0. 0."
        
        '*************************************************
        '
        ' FEC MODIFICACION : 03/01/2008
        ' SE REEMPLAZA :
        ' SQLVAR = " SELECT count(*) FROM plahistorico WHERE proceso='02' and cia='" & wcia & "' and status!='*' and placod='" & ArrReporte(COL_CODIGO, MAXROW) & "' AND year(fecperiodovaca)=" & Me.Txtano.Text & " and month(fecperiodovaca)= " & Cmbmes.ListIndex + 1
        ' POR :
        ' SQLVAR = " SELECT count(*) FROM plahistorico WHERE proceso='02' and cia='" & wcia & "' and status!='*' and placod='" & ArrReporte(COL_CODIGO, MAXROW) & "' AND year(fecperiodovaca)=" & Me.Txtano.Text & " and month(fecperiodovaca)= " & Cmbmes.ListIndex
        ' MODIFICADO POR RICARDO HINOSTROZA
        '
        '*************************************************
        '
        ' FECHA DE MODIFICACION : 23/01/2008
        ' SE MODIFICA MONTO DE VACACIONES PAGADAS
        '(LO PROVISIONADO ES REEMPLAZADO POR LO REAL EN PLANILLA
        ' MODIFICADO POR RICARDO HINOSOTROZA
        '
        '*************************************************
        
        Dim MONTO_VACACIONESTOMADAS As Double
        MONTO_VACACIONESTOMADAS = 0
        
        'SQLVAR = " SELECT count(*) FROM plahistorico WHERE proceso='02' and cia='" & wcia & "' and status!='*' and placod='" & ArrReporte(COL_CODIGO, MAXROW) & "' AND year(fecperiodovaca)=" & Me.Txtano.Text & " and month(fecperiodovaca)= " & Cmbmes.ListIndex + 1
                
        SQLVAR = " SELECT isnull(count(placod),0),isnull(sum(totaling),0)FROM plahistorico WHERE proceso='02' and cia='" & wcia & "' and status!='*' and placod='" & ArrReporte(COL_CODIGO, MAXROW) & "' AND year(FECHAPROCESO)=" & Me.Txtano.Text & " and month(FECHAPROCESO)= " & Cmbmes.ListIndex + 1
        
        If (fAbrRst(RX, SQLVAR)) Then
            cargaimporte = RX(0)
            'VALOR VAC. TOMADAS NO SE CONSIDERA PARA GRUPO GALLOS
            MONTO_VACACIONESTOMADAS = IIf(wGrupoPla <> "01", RX(1), 0)
            VacTomadasMes = " " & Trim(RX(0)) & ". 0. 0"
            RX.Close
        End If
        
        'IMPORTE DE LAS VACACIONES TOMADAS
        
        'ArrReporte(I, MAXROW) = CalculaImpVaca(VacTomadasMes, ArrReporte(MaxColTemp, MAXROW), rs!TipoTrabajador, cargaimporte)
        
        ArrReporte(I, MAXROW) = MONTO_VACACIONESTOMADAS
        I = I + 1

        'IMPORTE DE LAS VACACIONES X PAGAR
        ArrReporte(I, MAXROW) = CalculaImpVaca(RcdAñoTotal, ArrReporte(MaxColTemp, MAXROW), rs!TipoTrabajador)
        I = I + 1
        'IMPORTE DE LAS PROVISIONES DEL MES
        
        
        If MONTO_VACACIONESTOMADAS > 0 Then
            If Trim(rs!PlaCod) = "TO008" And Trim(wcia) = "03" Then
                ArrReporte(I, MAXROW) = Abs(ArrReporte(MaxColTemp + 3, MAXROW) - (ArrReporte(MaxColTemp + 1, MAXROW) - ArrReporte(MaxColTemp + 2, MAXROW))) - MONTO_VACACIONESTOMADAS
            Else
                If ArrReporte(MaxColTemp + 2, MAXROW) <> 0 Then
                   'SE MODIFICO EL CALCULO POR ORDEN DE JCBR 22/06/2009
                   ArrReporte(I, MAXROW) = ArrReporte(MaxColTemp + 3, MAXROW) - (ArrReporte(MaxColTemp + 1, MAXROW) - ArrReporte(MaxColTemp + 2, MAXROW))
                   'ArrReporte(I, MAXROW) = ArrReporte(MaxColTemp + 1, MAXROW) - ArrReporte(MaxColTemp + 2, MAXROW)
                Else
                   ArrReporte(I, MAXROW) = ArrReporte(MaxColTemp + 3, MAXROW) - (ArrReporte(MaxColTemp + 1, MAXROW) - ArrReporte(MaxColTemp + 2, MAXROW))
                End If
            End If
        Else
            ArrReporte(I, MAXROW) = ArrReporte(MaxColTemp + 3, MAXROW) - (ArrReporte(MaxColTemp + 1, MAXROW) - ArrReporte(MaxColTemp + 2, MAXROW))
            
            If wGrupoPla = "01" Then
                ArrReporte(I, MAXROW) = IIf(Val(ArrReporte(I, MAXROW)) < 0, 0, ArrReporte(I, MAXROW))
            End If
        End If
        
        I = I + 1
        'AREA DEL PERSONAL
        ArrReporte(I, MAXROW) = Trim(rs!Area)
        
        MAXROW = MAXROW + 1
    
        'RECORREMOS LOS CAMPOS DE INGRESO DE LA TABLA X(
        sSQLI = ""
        For I = 1 To 50
            Campo = "I" & Format(I, "00")
            sCol = BuscaColumna(Campo, MAXCOL)
            If sCol > 0 Then
                sSQLI = sSQLI & ArrReporte(sCol, MAXROW - 1) & ","
            Else
                sSQLI = sSQLI & "0,"
            End If
        Next
        
        'RECORREMOS LOS CAMPOS DE PROMEDIO DE LA TABLA X(
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

        sSQL = "INSERT plaprovvaca VALUES ('" & wcia & "','" & ArrReporte(COL_CODIGO, MAXROW - 1) & "','" & Format(Cmbtipo.ItemData(Cmbtipo.ListIndex), "00") & "','" & ArrReporte(COL_RCDAÑOANTERIOR, MAXROW - 1) & "',"
        sSQL = sSQL & "'" & ArrReporte(COL_VACACTOMADAS, MAXROW - 1) & "','" & ArrReporte(COL_RCDAÑOACTUAL, MAXROW - 1) & "','" & ArrReporte(COL_RCDACUMULADO, MAXROW - 1) & "'," & sSQLI & sSQLP & ArrReporte(MaxColTemp, MAXROW - 1) & ","
        sSQL = sSQL & ArrReporte(MaxColTemp + 1, MAXROW - 1) & "," & ArrReporte(MaxColTemp + 2, MAXROW - 1) & "," & ArrReporte(MaxColTemp + 3, MAXROW - 1) & "," & ArrReporte(MaxColTemp + 4, MAXROW - 1) & ",'" & FecFin & "',GETDATE(),'" & wuser & "','" & ArrReporte(MaxColTemp + 5, MAXROW - 1) & "',' ','" & Format(rs!fIngreso, "DD/MM/YYYY") & "')"

        cn.Execute (sSQL)
    End If
    rs.MoveNext
Loop
    Panelprogress.Caption = "Completo ..."
End If

End Sub
Private Function BuscaColumna(ByVal pCampo As String, ByVal pMaxcol As Integer) As Integer
Dim iRow As Integer
BuscaColumna = 0
For iRow = 3 To pMaxcol
    Debug.Print ArrReporte(iRow, 0)
    If ArrReporte(iRow, 0) = pCampo Then
        BuscaColumna = iRow
        Exit Function
    End If
Next

End Function

Private Function CalculaImpVaca(ByVal pFecha As String, ByVal pImporteRemun As String, ByVal pTipoTrab As String, Optional pCarga As Integer = 1) As String
Dim rs_CalculaImp As ADODB.Recordset
Dim ArrFecha As Variant
Dim I As Integer
Dim CalculaImpVacaTmp As Currency
Dim año As Currency, Mes As Currency, Dia As Currency

If pCarga = 1 Then

    ArrFecha = Split(pFecha, ".")
    
    If pTipoTrab = "01" Then
        'EMPLEADOS
        año = pImporteRemun
        If wGrupoPla <> "01" Then
                If wcia = "05" Then
                    Mes = pImporteRemun / (12 * 2)
                Else
                    Mes = pImporteRemun / (12)
                End If
        Else
            'IMPLEMENTACION GALLOS
            Mes = pImporteRemun / (12)
       End If
       'Dia = pImporteRemun / 365
       Dia = pImporteRemun / 360
    Else
        'OBREROS
        If wGrupoPla <> "01" Then
            If wcia = "05" Then
                Mes = Round((pImporteRemun * 30) / (12 * 2), 2)
                año = Round((pImporteRemun * 30) / 2, 2)
            Else
                Mes = Round((pImporteRemun * 30) / 12, 2)
                año = Round((pImporteRemun * 30), 2)
            End If
            'Dia = Round((pImporteRemun * 30) / 365, 2)
            Dia = Round((pImporteRemun * 30) / 360, 2)
        Else
            'IMPLEMENTACION GALLOS
            año = (pImporteRemun * 30)
            Mes = (pImporteRemun * 30) / 12
            'Dia = (pImporteRemun * 30) / 365
            Dia = (pImporteRemun * 30) / 360
        End If
        
    End If

    CalculaImpVacaTmp = 0
    For I = 0 To UBound(ArrFecha)
        Select Case I
        Case Is = 0
            ArrFecha(I) = IIf(wGrupoPla <> "01", año * Val(ArrFecha(I)), Round(año * Val(ArrFecha(I)), 2))
            CalculaImpVacaTmp = CalculaImpVacaTmp + ArrFecha(I)
        Case Is = 1
            If Mid(ArrFecha(1), 1, 1) = "-" Then
                ArrFecha(I) = IIf(wGrupoPla <> "01", Mes * Val(ArrFecha(I)), Round(Mes * Val(ArrFecha(I)), 2))
            Else
                ArrFecha(I) = IIf(wGrupoPla <> "01", Mes * Abs(Val(ArrFecha(I))), Round(Mes * Abs(Val(ArrFecha(I))), 2))
            End If
            CalculaImpVacaTmp = CalculaImpVacaTmp + ArrFecha(I)
        Case Is = 2
            If wGrupoPla <> "01" Then
               ArrFecha(I) = Format(Dia * Val(ArrFecha(I)), "0.00")
            Else
                ArrFecha(I) = Round(Dia * Val(ArrFecha(I)), 2)
            End If
            CalculaImpVacaTmp = CalculaImpVacaTmp + ArrFecha(I)
        End Select
    Next I
End If
CalculaImpVaca = IIf(wGrupoPla <> "01", Format(CalculaImpVacaTmp, "#0.00"), CalculaImpVacaTmp)

End Function
Private Sub ObtenerFechas(ByVal pPlacod As String, ByVal pFecIng As String, ByVal pFecIniProc As String, ByVal pFecproc As String, ByRef pRcdAcumuPasado As String, ByRef pVacacTomadas As String, ByRef pRcdActual As String, pRcdPerdida As String, ByRef pRcdAcumulado As String, ByRef pImpVacaAnterior As String)
Dim sSQL As String
Dim resultado As String
Dim FECINICIO As String
Dim FecIngTmp As String
Dim Saño As String, Smes As String, sDIA As String

FecIngTmp = pFecIng

'Vacaciones Perdidas
sSQL = "Select mes as Perdidas From Plavacaper where cia='" & wcia & "' and Ayo= " & Val(Txtano.Text) & " and Mes<= " & Cmbmes.ListIndex + 1 & " and Placod='" & Trim(pPlacod) & "' and  Status<>'*' "
If (fAbrRst(rs, sSQL)) Then
    resultado = " 1. 0. 0"
Else
    resultado = " 0. 0. 0"
End If
pRcdPerdida = resultado
rs.Close
'Fin Vaca Perdidas

sSQL = "SELECT recordacu FROM plaprovvaca WHERE PLACOD='" & Trim(pPlacod) & "' AND STATUS!='*' AND YEAR(fechaproceso)=" & Val(Txtano.Text) - 1 & " AND MONTH(fechaproceso)=12  And LEFT(CONVERT(VARCHAR,fechaproceso,112),8) >= '" & Format(pFecIng, "yyyymmdd") & "' "

If (fAbrRst(rs, sSQL)) Then
    resultado = rs(0)
Else
    resultado = " 0. 0. 0"
End If
pRcdAcumuPasado = resultado



sSQL = "SELECT provtotal FROM plaprovvaca WHERE PLACOD='" & Trim(pPlacod) & "' AND STATUS!='*' AND YEAR(fechaproceso)=" & Year(DateAdd("m", -1, pFecproc)) & " AND MONTH(fechaproceso)=" & Month(DateAdd("m", -1, pFecproc))
If (fAbrRst(rs, sSQL)) Then
    resultado = rs(0)
Else
    resultado = "0"
End If
pImpVacaAnterior = resultado

sSQL = "SELECT COUNT(*) FROM PLAHISTORICO WHERE placod='" & Trim(pPlacod) & "' and status!='*' and proceso='02' and  fecperiodovaca>='" & pFecIniProc & "' AND fecperiodovaca<='" & pFecproc & "' And LEFT(CONVERT(VARCHAR,fechaproceso,112),8) >= '" & Format(pFecIng, "yyyymmdd") & "' "

If (fAbrRst(rs, sSQL)) Then
    resultado = Space(2 - Len(Trim(rs(0)))) & Trim(rs(0)) & ". 0. 0"
Else
    resultado = " 0. 0. 0"
End If

pVacacTomadas = resultado

Dim CantDiasMesAños() As String

CantDiasMesAños = ObtenerFechasGallos(IIf(Year(pFecIng) = Year(pFecproc), pFecIng, DateAdd("d", -1, pFecIniProc)), pFecproc, pFecIng)

sDIA = CantDiasMesAños(0)
Smes = CantDiasMesAños(1)
Saño = CantDiasMesAños(2)

pRcdActual = Space(2 - Len(Trim(Saño))) & Saño & "." & Space(2 - Len(Trim(Smes))) & Abs(Smes) & "." & Space(2 - Len(Trim(sDIA))) & sDIA

If Len(pRcdActual) < 8 Then
    If Mid(pRcdActual, 2, 1) = "." Then
        pRcdActual = " " & pRcdActual
    End If
    If Mid(pRcdActual, 5, 1) = "." Then
        pRcdActual = Mid(pRcdActual, 1, 3) & " " & Mid(pRcdActual, 4, 7)
    End If
End If

Saño = "": Smes = "": sDIA = ""

Saño = Val(Mid(resultado, 1, 2)) + Val(Mid(pRcdActual, 1, 2))

Saño = Val(Mid(pRcdAcumuPasado, 1, 2)) + Val(Mid(pRcdActual, 1, 2)) - Val(Mid(pRcdPerdida, 1, 2))
Smes = Val(Mid(pRcdAcumuPasado, 4, 2)) + Val(Mid(pRcdActual, 4, 2)) - Val(Mid(pRcdPerdida, 4, 2))
sDIA = Val(Mid(pRcdAcumuPasado, 7, 2)) + Val(Mid(pRcdActual, 7, 2)) - Val(Mid(pRcdPerdida, 7, 2))

Dias:
If sDIA >= 30 Then
    Smes = Val(Smes) + 1
    sDIA = Val(sDIA) - 30
    GoTo Dias
End If

Mes:
If Smes >= 12 Then
    Saño = Val(Saño) + 1
    Smes = Val(Smes) - 12
    GoTo Mes
End If

If Val(Saño) < 0 Then
    Saño = 0
    Smes = Val(Smes) - 12
End If

resultado = Space(2 - Len(Trim(Saño))) & Saño & "." & Space(2 - Len(Trim(Smes))) & Smes & "." & Space(2 - Len(Trim(sDIA))) & sDIA

Saño = "": Smes = "": sDIA = ""
Saño = Val(Mid(resultado, 1, 2)) - Val(Mid(pVacacTomadas, 1, 2))

'-------------------------------------------
'Modificado el 03/06/2008
'Calculo de Provision Negativa
'para la Cia. Tripsa
'--------------------------------------------

If Saño = "-1" And Mid(pVacacTomadas, 1, 2) = " 1" Then ' And wcia = "03" Then

    If wGrupoPla <> "01" Then
        Smes = (Val(Mid(resultado, 4, 2)) - 12) '+ 1
        sDIA = 30 - Val(Mid(resultado, 7, 2))
        Saño = 0
    Else
        Saño = 0: sDIA = 0: Smes = 0
    End If
Else
    Smes = Val(Mid(resultado, 4, 2)) - Val(Mid(pVacacTomadas, 4, 2))
    sDIA = Val(Mid(resultado, 7, 2)) - Val(Mid(pVacacTomadas, 7, 2))
End If

masdiasACTUAL:
If sDIA >= 30 Then
    Smes = Val(Smes) + 1
    sDIA = Val(sDIA) - 30
    GoTo masdiasACTUAL
End If

masmesACTUAL:
If Smes >= 12 Then
    Saño = Val(Saño) + 1
    Smes = Val(Smes) - 12
    GoTo masmesACTUAL
End If

If Val(Saño) < 0 Then
    Saño = 0
    Smes = Val(Smes) - 12
End If

pRcdAcumulado = Space(2 - Len(Trim(Saño))) & Saño & "." & Space(2 - Len(Trim(Smes))) & Smes & "." & Space(2 - Len(Trim(sDIA))) & sDIA

End Sub
Private Function ObtenerFechasGallos(ByVal pFecInicio As String, ByVal pFecFinal As String, ByVal pFecIngreso As String) As String()

Dim mfecin As String
Dim mdia1, mdia2 As String
Dim manos, meses, mdias As String
Dim DiasMesesAños(3) As String
Dim mUltDiaFecIng As Integer

mfecin = pFecInicio

mdia1 = Day(mfecin)
mdia2 = Day(pFecFinal)

manos = Year(pFecFinal) - Year(mfecin)

meses = Month(pFecFinal) - Month(mfecin)

If Format(pFecInicio, "yyyy") = Format(pFecFinal, "yyyy") Then
   If Format(pFecInicio, "yyyymm") = Format(pFecFinal, "yyyymm") Then
        mdias = IIf(Day(pFecFinal) = Day(mfecin), 1, (DateDiff("d", mfecin, pFecFinal) + IIf(Day(pFecFinal) <> 31, 1, 0)))
    Else
        mUltDiaFecIng = Ultimo_Dia(Month(pFecInicio), Year(pFecInicio))
        mdias = IIf(mUltDiaFecIng = Day(pFecInicio), 1, mUltDiaFecIng - Day(pFecInicio) + IIf(mUltDiaFecIng <> 31, 1, 0))
    End If
Else
      mdias = 0
End If

If mdias < 0 Then
  meses = meses - 1
    If Month(pFecFinal) <> 2 Then
        mdias = mdias + 30
    Else
        If (Year(pFecFinal) Mod 4) = 0 Then
            mdias = mdias + 29
        Else
            mdias = mdias + 28
        End If
    End If
End If

'If mdias >= 29 Then
'   meses = meses + 1
'   mdias = 0
'End If

If meses > 12 Then
   meses = meses - 1
   manos = manos + 1
End If
'
If meses < 0 Then
    manos = manos - 1
    meses = meses + 12
End If

If manos < 0 Then
   manos = 0
End If

'If (mdias = 25 And Month(pFecFinal) = 2) Or (mdias > 24 And Month(pFecFinal) = 2) Then
'    meses = meses + 1
'    mdias = 0
'End If

DiasMesesAños(0) = mdias
DiasMesesAños(1) = meses
DiasMesesAños(2) = manos

ObtenerFechasGallos = DiasMesesAños
End Function



Private Function Otros_Pagos(ByVal PlaCod As String, ByVal FechaProceso As String, ByVal Tipo_Bol As String) As Double

Otros_Pagos = 0

Dim mCadOtPag As String
Dim mPer As Integer
Dim mDateBeginVac As String
Dim mDateEndVac As String
Dim mdia As Integer
Dim meses As Integer
Dim RX As ADODB.Recordset
Dim FecIni As String, FecFin As String
Dim mFecProceso As String
                   
mFecProceso = Ultimo_Dia(Month(DateAdd("m", 1, FechaProceso)), Year(CDate(FechaProceso))) & "/" & Format(Month(DateAdd("m", 1, FechaProceso)), "00") & "/" & Year(CDate(FechaProceso))
If Month(mFecProceso) = 7 Then
    FecIni = "01/01/" & Year(mFecProceso)
    FecFin = "07/01/" & Year(mFecProceso)
Else
    FecIni = "07/01/" & Year(mFecProceso)
    FecFin = "12/01/" & Year(mFecProceso)
End If

mPer = 0
      Sql = "select distinct(codinterno),factor,factor_divisionario from platasaanexo where cia='" & wcia & "' and tipomovimiento='" & VTipo & "' and modulo='01'  and basecalculo='16' and status<>'*'"
      If (fAbrRst(rs, Sql)) Then rs.MoveFirst: mPer = Val(rs(2)): meses = Val(rs(1))
      mCadOtPag = ""
      rs.MoveFirst
      Do While Not rs.EOF
        If Tipo_Bol = "03" Then
              Sql = "SET DATEFORMAT MDY SELECT COUNT(" & "i" & rs(0) & ") FROM plahistorico WHERE cia='" & wcia & "' and placod='" & PlaCod & "' and status!='*' and (fechaproceso>'" & FecIni & "' and fechaproceso<'" & FecFin & "') and " & "i" & rs(0) & ">0 AND proceso='01'"
              If (fAbrRst(RX, Sql)) Then
                  If RX(0) >= 3 Then
                       mCadOtPag = mCadOtPag & "i" & rs(0) & "+"
                  End If
              End If
        Else
           mCadOtPag = mCadOtPag & "i" & rs(0) & "+"
        End If
         rs.MoveNext
      Loop
      rs.Close
   

If InStr(1, mCadOtPag, "10") > 0 Or InStr(1, mCadOtPag, "11") > 0 Or InStr(1, mCadOtPag, "21") > 0 Then
    If InStr(1, mCadOtPag, "10") = 0 Then mCadOtPag = mCadOtPag & "i10+"
    If InStr(1, mCadOtPag, "11") = 0 Then mCadOtPag = mCadOtPag & "i11+"
    If InStr(1, mCadOtPag, "21") = 0 Then mCadOtPag = mCadOtPag & "i21+"
End If

Sql = ""

mDateBeginVac = Fecha_Promedios(mPer, mFecProceso)

If Val(Mid(mFecProceso, 4, 2)) = 1 Then
   mdia = Ultimo_Dia(12, Val(Mid(mFecProceso, 7, 4)) - 1)
   mDateEndVac = Format(mdia, "00") & "/12/" & Format(Val(Mid(mFecProceso, 7, 4)) - 1, "0000")
Else
   mdia = Ultimo_Dia(Val(Mid(mFecProceso, 4, 2) - 1), Val(Mid(mFecProceso, 7, 4)))
   mDateEndVac = Format(mdia, "00") & "/" & Format(Val(Mid(mFecProceso, 4, 2) - 1), "00") & "/" & Mid(mFecProceso, 7, 4)
End If

mDateBeginVac = "01/" & Format(Month(DateAdd("m", (meses - 1) * -1, mDateEndVac)), "00") & "/" & Year(DateAdd("m", -5, mDateEndVac))

If Trim(mCadOtPag) <> "" Then
    mCadOtPag = Mid(mCadOtPag, 1, Len(Trim(mCadOtPag)) - 1)
    mCadOtPag = "sum(" & mCadOtPag & ")"
    
    Dim FECINICIO As String
    Dim FECFINAL As String
    
    FECINICIO = Format(mDateBeginVac, "mm/dd/yyyy")
    FECFINAL = Format(mDateEndVac, "mm/dd/yyyy")
    
    Sql = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " select "
    Sql = Sql & "CASE WHEN dbo.FC_VALIDAPAGOS('" & wcia & "','" & PlaCod & "','" & FECINICIO & "','" & FECFINAL & "')>=3 THEN "
    Sql = Sql & mCadOtPag & "else 0 end from plahistorico"
End If

If Trim(Sql) <> "" Then
    Sql = Sql & " where cia='" & wcia & "' and fechaproceso Between '" & Format(mDateBeginVac, FormatFecha) & Space(1) & FormatTimei & "' and '" & Format(mDateEndVac, FormatFecha) & Space(1) & FormatTimef & "'"
    Sql = Sql & " and proceso='01' and placod='" & PlaCod & "' and status<>'*'"
    
    If (fAbrRst(rs, Sql)) Then rs.MoveFirst
    If Not IsNull(rs(0)) Then
        If VTipo = "01" Then
            Otros_Pagos = rs(0) / mPer
        Else
            Otros_Pagos = (rs(0) / 6) / 30
        End If
    End If
End If

End Function
Private Sub PROVISIONES_GRATI()
Dim sSQL As String
Dim FecFin As String
Dim FecIni As String

FecFin = Format(Format(Ultimo_Dia(Cmbmes.ListIndex + 1, Val(Txtano.Text)), "00") & "/" & Format(Cmbmes.ListIndex + 1, "00") & "/" & Txtano.Text, "DD/MM/YYYY")
FecIni = DateAdd("m", -5, CDate("01/" & Month(FecFin) & "/" & Year(FecFin)))

sSQL = "usp_Pla_Calcula_PRov_Grati '" & wcia & "','" & VTipo & "','" & FecIni & "','" & FecFin & "','" & wuser & "'"
cn.Execute (sSQL)

End Sub
Private Sub PROVISIONES_GRATI_ANTES()
Dim sSQL As String
Dim MAXROW As Long, MAXCOL As Integer, MaxColInicial As Integer
Dim rs As ADODB.Recordset, rsAux As ADODB.Recordset
Dim CantMes As String, Campo As String
Dim FecIni As String, FecFin As String, FecProceso As String
Dim I As Integer, MaxColTemp As Integer
Dim dblFactor As Currency, Cadena As String
Dim Factor_EsSalud As Currency, totaportes As Currency
Dim sCol As Integer, curfactor As Currency
Dim sSQLI As String, sSQLP As String
Dim PAÑO_PROCESO As Integer

'IMPLEMENTACION UNIFICADA
Dim CantDias As String
Dim nVecesPer As Integer

Dim NroTrans As Integer
On Error GoTo ErrorTrans
NroTrans = 0

CantDias = 0


Const COL_CODIGO = 0
Const COL_FECING = 1
Const COL_AREA = 2

MAXCOL = 2
I = MAXCOL + 1
FecProceso = Format(DateAdd("d", -1, Format(DateAdd("m", 1, "01/" & Cmbmes.ListIndex + 1 & "/" & Txtano.Text), "dd/mm/yyyy")), "dd/mm/yyyy")

FecFin = Format(Format(Ultimo_Dia(Cmbmes.ListIndex + 1, Val(Txtano.Text)), "00") & "/" & Format(Cmbmes.ListIndex + 1, "00") & "/" & Txtano.Text, "DD/MM/YYYY")
FecIni = DateAdd("m", -5, CDate("01/" & Month(FecFin) & "/" & Year(FecFin)))

'IMPLEMENTACION GALLOS

nVecesPer = 3

'OBTENEMOS LOS CONCEPTOS REMUNERATIVOS A TRABAJAR
  '-- HORAS EXTRAS --
  
If Month(FecFin) >= 1 And Month(FecFin) < 6 Then
   PAÑO_PROCESO = Val(Year(FecFin)) - 1
Else
   PAÑO_PROCESO = Val(Year(FecFin))
End If


sSQL = "select distinct factor from platasaanexo where status!='*' and tipomovimiento='02' and basecalculo=16 and cia='" & wcia & "'"

If (fAbrRst(rs, sSQL)) Then
    dblFactor = rs(0)
    rs.Close
End If

Erase ArrReporte

sSQL = "SELECT concepto,campo,sn_promedio,0 AS factor,sn_carga,formula FROM estructura_provisiones WHERE tipo='G' and "
sSQL = sSQL & " campo NOT IN (SELECT codinterno FROM platasaanexo WHERE cia='" & wcia & "' and tipomovimiento='02' and basecalculo='16' and status != '*')"
sSQL = sSQL & " UNION ALL SELECT 'P' As concepto,campo,sn_promedio,b.factor,sn_carga,formula FROM estructura_provisiones a INNER JOIN platasaanexo b ON"
sSQL = sSQL & " (b.cia='" & wcia & "' and b.tipomovimiento='02' and b.basecalculo='16' and b.codinterno=a.campo and status != '*') WHERE a.tipo='G'"

If (fAbrRst(rs, sSQL)) Then
    Do While Not rs.EOF
        Debug.Print rs(1) & Space(2) & MAXCOL
        MAXCOL = MAXCOL + 1
        rs.MoveNext
    Loop
    rs.MoveFirst
    MaxColTemp = MAXCOL + 1
    MAXCOL = MAXCOL + 5
    
    ReDim Preserve ArrReporte(0 To MAXCOL, 0 To MAXROW)
    Cadena = ""
    
    Dim FiltroFecIng As String
    
    Do While Not rs.EOF
        
        ArrReporte(I, MAXROW) = Trim(rs!concepto) & rs!Campo
        
        FiltroFecIng = " AND LEFT(CONVERT(VARCHAR,ph.fechaproceso,112),8) >= LEFT(CONVERT(VARCHAR,a.fingreso,112),8)"
        
        'CALCULO DE PROMEDIOS
        
        If CInt(rs!sn_carga) <> 0 Then
            If Trim(rs!Formula) & "" = "" Then
                Cadena = Cadena & IIf(CInt(rs!sn_promedio) = 0, "I", "P") & rs!Campo & "=CASE WHEN dbo.fc_cargapagos2('" & wcia & "',a.placod,'" & Val(Month(FecIni)) & "','" & Val(Month(FecFin)) & "','" & PAÑO_PROCESO & "','" & "B" & "')>=" & nVecesPer & " THEN (SELECT ISNULL(SUM(COALESCE(I" & Trim(rs!Campo) & ",0)),0) FROM plahistorico ph WHERE ph.cia='" & wcia & "' and ph.status!='*'  and  ph.FECHAPROCESO>='" & FecIni & "' " & _
                "AND ph.FECHAPROCESO<='" & FecFin & "'" & FiltroFecIng & " AND ph.placod=a.placod AND PH.PROCESO='01')"
            
                'Cadena = Cadena & IIf(CInt(rs!sn_promedio) = 0, "I", "P") & rs!Campo & "=CASE WHEN dbo.fc_validapagos('" & wcia & "',a.placod,'" & FecIni & "','" & FecFin & "','" & "B" & "')>=" & nVecesPer & " THEN (SELECT ISNULL(SUM(COALESCE(I" & Trim(rs!Campo) & ",0)),0) FROM plahistorico ph WHERE ph.cia='" & wcia & "' and ph.status!='*'  and  ph.FECHAPROCESO>='" & FecIni & "' " & _
                '"AND ph.FECHAPROCESO<='" & FecFin & "'" & FiltroFecIng & " AND ph.placod=a.placod AND PH.PROCESO='01')"
            Else
                Cadena = Cadena & IIf(CInt(rs!sn_promedio) = 0, "I", "P") & rs!Campo & "=CASE WHEN dbo.fc_cargapagos2('" & wcia & "',a.placod,'" & Val(Month(FecIni)) & "','" & Val(Month(FecFin)) & "','" & PAÑO_PROCESO & "','" & "E" & "')>=" & nVecesPer & " THEN (SELECT ISNULL(SUM(" & rs!Formula & "),0) FROM plahistorico ph WHERE ph.cia='" & wcia & "' and ph.status!='*'  and  ph.FECHAPROCESO>='" & FecIni & "' " & _
                "AND ph.FECHAPROCESO<='" & FecFin & "'" & FiltroFecIng & " AND ph.placod=a.placod AND PH.PROCESO='01')"
            
                'Cadena = Cadena & IIf(CInt(rs!sn_promedio) = 0, "I", "P") & rs!Campo & "=CASE WHEN dbo.fc_validapagos('" & wcia & "',a.placod,'" & FecIni & "','" & FecFin & "','" & "E" & "')>=" & nVecesPer & " THEN (SELECT ISNULL(SUM(" & rs!Formula & "),0) FROM plahistorico ph WHERE ph.cia='" & wcia & "' and ph.status!='*'  and  ph.FECHAPROCESO>='" & FecIni & "' " & _
                '"AND ph.FECHAPROCESO<='" & FecFin & "'" & FiltroFecIng & " AND ph.placod=a.placod AND PH.PROCESO='01')"
            End If
                
            
            If CInt(rs!sn_promedio) = -1 And Trim(rs!concepto) = "A" Then
                Cadena = Cadena & "/(SELECT aportacion FROM PLACONSTANTE WHERE tipomovimiento='03' AND cia='" & wcia & "' AND status!='*' AND codinterno='" & rs!Campo & "'),"
            
            'ElseIf CInt(rs!sn_promedio) = -1 And Trim(rs!concepto) = "I" And wGrupoPla <> "01" Then
            '    Cadena = Cadena & "/" & rs!factor & ","
            
            ElseIf CInt(rs!sn_promedio) = -1 Then 'And wGrupoPla = "01" Then
                'If Format(Cmbtipo.ItemData(Cmbtipo.ListIndex), "00") = "01" Then
                '    'EMPLEADOS
                '    Cadena = Cadena & "/" & dblFactor & " ELSE 0 END,"
                'Else
                '    'OBREROS
                    Cadena = Cadena & "/180" & " ELSE 0 END,"
                'End If
            Else
                Cadena = Cadena & ","
            End If
                

        Else
            Cadena = Cadena & "COALESCE((SELECT prb.importe/(factor_horas/" & hORAS_X_DIA & ") FROM plaremunbase prb WHERE cia='" & wcia & "' and placod=a.placod AND status!='*' and concepto='" & Trim(rs!Campo) & "'),0) as '" & Trim(rs!concepto) & Trim(rs!Campo) & "',"
        End If
        
        I = I + 1
        rs.MoveNext
    Loop
    
    'ArrReporte(i, MAXROW) = "I16"
    Cadena = Cadena & "0 AS I16"
    'i = i + 1
    'Cadena = Mid(Cadena, 1, Len(Trim(Cadena)) - 1)
    rs.Close
End If

Dim Fecha_CI As Date
Dim FechaCF As Date
Fecha_CI = DateSerial(Txtano.Text, (Cmbmes.ListIndex + 1) + 0, 1)
FechaCF = DateSerial(Txtano.Text, (Cmbmes.ListIndex + 1) + 1, 0)

sSQL = ""
sSQL = "SET DATEFORMAT DMY SELECT"


'IMPLEMENTACION GALLOS
'NO SE ECEPTUA AL TRABAJADOR SI SU FECHA DE INGRESO ES EN LA FECHA DE PROCESO
    
sSQL = sSQL & " a.placod,LTRIM(RTRIM(a.ap_pat))+' '+LTRIM(RTRIM(a.ap_mat))+' '+LTRIM(RTRIM(a.ap_cas))+' '+LTRIM(RTRIM(a.nom_1))+' '+LTRIM(RTRIM(a.nom_2)),"
sSQL = sSQL & " fingreso , fcese, a.area,a.tipotrabajador,"
sSQL = sSQL & Cadena
sSQL = sSQL & " FROM planillas a LEFT OUTER JOIN plahistorico ph ON (ph.cia='" & wcia & "' and ph.status!='*' "
sSQL = sSQL & " and  ph.FECHAPROCESO>='" & FecIni & "' AND ph.FECHAPROCESO<='" & FecFin & "' And ph.proceso='01' and ph.placod=a.placod)"

sSQL = sSQL & " WHERE a.cat_trab <> '04' and a.status!='*' AND (a.fcese is null OR  a.fcese>'" & FechaCF & "') "
sSQL = sSQL & "AND a.tipotrabajador='" & VTipo & "' and a.cia='" & wcia & "' "
sSQL = sSQL & "AND a.fingreso<='" & FechaCF & "' GROUP BY a.PLACOD , a.ap_pat, a.ap_mat, a.ap_cas, a.nom_1, a.nom_2, a.fingreso, a.fcese,a.area,a.tipotrabajador"

 MAXROW = MAXROW + 1
 
 Dim dBasico As Double
 Dim dAportacion As Double
 Dim dAuxiliar As Double
 
If (fAbrRst(rs, sSQL)) Then
    DoEvents
    Panelprogress.Caption = "Generando Provisión ..."
    
    cn.BeginTrans
    NroTrans = 1
    
    Barra.Min = 0
    Barra.Max = rs.RecordCount
    Do While Not rs.EOF
        Debug.Print rs!PlaCod
        Barra.Value = rs.AbsolutePosition
       
        ReDim Preserve ArrReporte(0 To MAXCOL, 0 To MAXROW)
        'If Trim(rs!PlaCod) = "RO378" Then Stop
        
        Dim CantDiasMes() As String
        CantDiasMes = CantidadDiasMesesCalculo(rs!fIngreso, FecProceso)
        CantDias = 0  'Se computa mes completo. Anteriormente para Gallos CantDiasMes(0)
        CantMes = CantDiasMes(1)
        
        If CantMes <> 0 Then
            ArrReporte(COL_CODIGO, MAXROW) = Trim(rs!PlaCod)
            ArrReporte(COL_FECING, MAXROW) = rs!fIngreso
            ArrReporte(COL_AREA, MAXROW) = rs!Area
            
            Factor_EsSalud = Trae_Porc_EsSalud_EPS(Trim(rs!PlaCod), Trim(rs!TipoTrabajador))
                    
            dBasico = 0: dAportacion = 0: totaportes = 0: dAuxiliar = 0
                            
            For I = 6 To rs.Fields.count - 1
                Debug.Print "Calculando =>" & rs.Fields(I).Name
                sCol = BuscaColumna(rs.Fields(I).Name, MAXCOL)
                If sCol > 0 Then
                    If Trim(rs!TipoTrabajador) = "01" Then
                        ArrReporte(sCol, MAXROW) = Round(rs.Fields(I).Value * DIAS_TRABAJO, 2)
                    Else
                        ArrReporte(sCol, MAXROW) = Round(rs.Fields(I).Value, 2)
                    End If
                    totaportes = totaportes + ArrReporte(sCol, MAXROW)
                End If
            Next
            
            
            For I = 3 To rs.Fields.count - 1
                If ArrReporte(I, 0) = "I16" Then
                    sCol = I
                    Exit For
                End If
            Next
            
            If sCol <> 0 Then
                'Son considerados en CALCULO DE PROMEDIOS
                ArrReporte(sCol, MAXROW) = 0
                'ArrReporte(sCol, MAXROW) = IIf(wGrupoPla <> "01", Otros_Pagos(Trim(rs!PlaCod), FecProceso, "03"), 0)
                totaportes = totaportes + ArrReporte(sCol, MAXROW)
            End If
            
            I = MaxColTemp
            ArrReporte(I, 0) = "P01"
            
            If wGrupoPla <> "01" Then
                'factor_essalud SE CAMBIA POR CERO, PARA TOMAR EN CUENTA NUEVAMENTE REEMPLAZAR EL CERO POR LA VARIABLE
                'factor_essalud SE CAMBIO POR 0
                ArrReporte(I, MAXROW) = Round(totaportes * (Factor_EsSalud / 100), 2)
            Else
                'IMPLEMENTACION GALLOS
                ArrReporte(I, MAXROW) = 0
            End If
            
            I = I + 1
            ArrReporte(I, MAXROW) = Round(totaportes + ArrReporte(I - 1, MAXROW), 2)
            totaportes = ArrReporte(I, MAXROW)
            I = I + 1
            
            If Cmbmes.ListIndex + 1 = 1 Or Cmbmes.ListIndex + 1 = 7 Then
                'PROVTOTAL
                Sql = "select provtotal from plaprovgrati where cia='" & wcia & "' and month(fechaproceso)=" & Cmbmes.ListIndex + 1 & " AND YEAR(FECHAPROCESO)=  " & Txtano.Text & " and placod='" & Trim(rs!PlaCod) & "' and status<>'*'"
            Else
                'PROVTOTAL
                Sql = "select provtotal from plaprovgrati where cia='" & wcia & "' and month(fechaproceso)=" & Cmbmes.ListIndex & " AND YEAR(FECHAPROCESO)=  " & Txtano.Text & " and placod='" & Trim(rs!PlaCod) & "' and status<>'*'"
            End If
            
            If (fAbrRst(rsAux, Sql)) Then
                ArrReporte(I, MAXROW) = rsAux(0)
                rsAux.Close
            Else
                ArrReporte(I, MAXROW) = 0
            End If
            
            I = I + 1
                
            If Trim(rs!TipoTrabajador) = "01" Then
                'EMPLEADOS
                ArrReporte(I, MAXROW) = Round((totaportes / 6) * CantMes + ((totaportes / 6) / 30 * CantDias), 2)
               'ArrReporte(I, MAXROW) = Round((totaportes / 6) * CantMes + IIf(CantMes = 0, (totaportes / 6) / 30 * CantDias, 0), 2)
            Else
                'OBREROS
                ArrReporte(I, MAXROW) = Round(((totaportes * 30) / 6) * CantMes + (totaportes / 6 * CantDias), 2)
                'ArrReporte(I, MAXROW) = Round(((totaportes * 30) / 6) * CantMes + IIf(CantMes = 0, totaportes / 6 * CantDias, 0), 2)
               
            End If
            
            
            I = I + 1
            ArrReporte(I, MAXROW) = Abs(ArrReporte(I - 1, MAXROW) - ArrReporte(I - 2, MAXROW))
    
            MAXROW = MAXROW + 1
            
            'RECORREMOS LOS CAMPOS DE INGRESO DE LA TABLA X(
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
            
            'RECORREMOS LOS CAMPOS DE PROMEDIO DE LA TABLA X(
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
            
            sSQL = ""
            sSQL = "SET DATEFORMAT DMY "
            sSQL = sSQL & "INSERT plaprovgrati VALUES ('" & wcia & "','" & ArrReporte(COL_CODIGO, MAXROW - 1) & "','" & Format(Cmbtipo.ItemData(Cmbtipo.ListIndex), "00") & "','',"
            sSQL = sSQL & " '','',''," & sSQLI & sSQLP & ArrReporte(MaxColTemp + 1, MAXROW - 1) & "," & ArrReporte(MaxColTemp + 2, MAXROW - 1) & ","
            sSQL = sSQL & ArrReporte(MaxColTemp + 2, MAXROW - 1) & "," & ArrReporte(MaxColTemp + 3, MAXROW - 1) & "," & ArrReporte(MaxColTemp + 4, MAXROW - 1) & ",'" & Format(FecProceso, "dd/mm/yyyy") & "',GETDATE(),'" & wuser & "','" & ArrReporte(COL_AREA, MAXROW - 1) & "',' ','" & Format(rs!fIngreso, "dd/mm/yyyy") & "')"
            
            cn.Execute (sSQL)
            totaportes = 0
        
        End If
        rs.MoveNext
    Loop
    DoEvents
    Panelprogress.Caption = "Completo ..."
    
    cn.CommitTrans
End If
    Carga_Prov_Grati
    
Exit Sub

ErrorTrans:

    If NroTrans = 1 Then
        cn.RollbackTrans
    End If
    
    MsgBox Err.Description, vbCritical, Me.Caption
End Sub
Private Sub PROVICIONES_GRATI_2011()
'Implementacion version 2011 Gallos-Roda
'Descontinuada

Dim sSQL As String
Dim MAXROW As Long, MAXCOL As Integer, MaxColInicial As Integer
Dim rs As ADODB.Recordset, rsAux As ADODB.Recordset
Dim CantMes As String, Campo As String
Dim FecIni As String, FecFin As String, FecProceso As String
Dim I As Integer, MaxColTemp As Integer
Dim dblFactor As Currency, Cadena As String
Dim Factor_EsSalud As Currency, totaportes As Currency
Dim sCol As Integer, curfactor As Currency
Dim sSQLI As String, sSQLP As String

'IMPLEMENTACION GALLOS
Dim CantDias As String
Dim nVecesPer As Integer

CantDias = 0


Const COL_CODIGO = 0
Const COL_FECING = 1
Const COL_AREA = 2

MAXCOL = 2
I = MAXCOL + 1
FecProceso = Format(DateAdd("d", -1, Format(DateAdd("m", 1, "01/" & Cmbmes.ListIndex + 1 & "/" & Txtano.Text), "dd/mm/yyyy")), "dd/mm/yyyy")

If Cmbmes.ListIndex + 1 < 7 Then
    FecIni = "01/01/" & Txtano.Text
    FecFin = Format(DateAdd("d", -1, "01/07/" & Txtano.Text), "DD/MM/YYYY")
Else
    FecIni = "01/07/" & Txtano.Text
    
    If wGrupoPla = "01" Then
        FecFin = Format(Format(Ultimo_Dia(Cmbmes.ListIndex + 1, Val(Txtano.Text)), "00") & "/" & Format(Cmbmes.ListIndex + 1, "00") & "/" & Txtano.Text, "DD/MM/YYYY")
    Else
        FecFin = Format(DateAdd("d", -1, CDate("01/01/" & Txtano.Text + 1)), "DD/MM/YYYY")
    End If
End If

'IMPLEMENTACION GALLOS
nVecesPer = IIf(wGrupoPla = "01", IIf(Format(FecProceso, "yyyymm") > "201112", "3", "1"), "3")

sSQL = "select distinct factor from platasaanexo where status!='*' and tipomovimiento='02' and basecalculo=16 and cia='" & wcia & "'"
If (fAbrRst(rs, sSQL)) Then
    dblFactor = rs(0)
    rs.Close
End If

Erase ArrReporte

If wGrupoPla = "01" Then
    sSQL = "SELECT concepto,campo,sn_promedio,0 AS factor,sn_carga,formula FROM estructura_provisiones WHERE tipo='G' and "
    sSQL = sSQL & " campo NOT IN (SELECT codinterno FROM platasaanexo WHERE cia='" & wcia & "' and tipomovimiento='02' and basecalculo='16' and status != '*')"
    sSQL = sSQL & " UNION ALL SELECT 'P' As concepto,campo,sn_promedio,b.factor,sn_carga,formula FROM estructura_provisiones a INNER JOIN platasaanexo b ON"
    sSQL = sSQL & " (b.cia='" & wcia & "' and b.tipomovimiento='02' and b.basecalculo='16' and b.codinterno=a.campo and status != '*') WHERE a.tipo='G'"
Else
    sSQL = "SELECT concepto,campo,sn_promedio,0 AS factor,sn_carga,formula FROM estructura_provisiones WHERE tipo='G' and "
    sSQL = sSQL & " campo NOT IN (SELECT codinterno FROM platasaanexo WHERE cia='" & wcia & "' and tipomovimiento='02' and basecalculo='16' and status != '*')"
    sSQL = sSQL & " UNION ALL SELECT concepto,campo,sn_promedio,b.factor,sn_carga,formula FROM estructura_provisiones a INNER JOIN platasaanexo b ON"
    sSQL = sSQL & " (b.cia='" & wcia & "' and b.tipomovimiento='02' and b.basecalculo='16' and b.codinterno=a.campo and status != '*') WHERE a.tipo='G'"
End If

If (fAbrRst(rs, sSQL)) Then
    Do While Not rs.EOF
        Debug.Print rs(1) & Space(2) & MAXCOL
        MAXCOL = MAXCOL + 1
        rs.MoveNext
    Loop
    rs.MoveFirst
    MaxColTemp = MAXCOL + 1
    MAXCOL = MAXCOL + 5
    
    ReDim Preserve ArrReporte(0 To MAXCOL, 0 To MAXROW)
    Cadena = ""
    
    Dim FiltroFecIng As String
    
    Do While Not rs.EOF
        
        ArrReporte(I, MAXROW) = Trim(rs!concepto) & rs!Campo
        FiltroFecIng = IIf(wGrupoPla = "01", " AND LEFT(CONVERT(VARCHAR,ph.fechaproceso,112),8) >= LEFT(CONVERT(VARCHAR,a.fingreso,112),8)", "")
        
        If CInt(rs!sn_carga) <> 0 Then
        
            If wGrupoPla = "01" Then
                'IMPLEMENTACION GALLOS
'                If Trim(rs!Formula) & "" = "" Then
'                    Cadena = Cadena & IIf(CInt(rs!sn_promedio) = 0, "I", "P") & rs!Campo & "=CASE WHEN dbo.fc_validapagos('" & wcia & "',a.placod,'" & FecIni & "','" & FecFin & "')>=" & nVecesPer & " THEN (SELECT ISNULL(SUM(COALESCE(I" & Trim(rs!Campo) & ",0)),0))"
'                   'Cadena = Cadena & IIf(CInt(rs!sn_promedio) = 0, "I", "P") & rs!Campo & "=SUM(COALESCE(i" & Trim(rs!Campo) & ",0))"
'                Else
'                    Cadena = Cadena & IIf(CInt(rs!sn_promedio) = 0, "I", "P") & rs!Campo & "=CASE WHEN dbo.fc_validapagos('" & wcia & "',a.placod,'" & FecIni & "','" & FecFin & "')>=" & nVecesPer & " THEN (SELECT ISNULL(SUM(COALESCE(" & Trim(rs!Formula) & ",0)),0))"
'                   'Cadena = Cadena & IIf(CInt(rs!sn_promedio) = 0, "I", "P") & rs!Campo & "=SUM(COALESCE(" & Trim(rs!Formula) & ",0))"
'                End If
                
                If Trim(rs!Formula) & "" = "" Then
                    Cadena = Cadena & IIf(CInt(rs!sn_promedio) = 0, "I", "P") & rs!Campo & "=CASE WHEN dbo.fc_validapagos('" & wcia & "',a.placod,'" & FecIni & "','" & FecFin & "')>=" & nVecesPer & " THEN (SELECT ISNULL(SUM(COALESCE(I" & Trim(rs!Campo) & ",0)),0) FROM plahistorico ph WHERE ph.cia='" & wcia & "' and ph.status!='*'  and  ph.FECHAPROCESO>='" & FecIni & "' " & _
                    "AND ph.FECHAPROCESO<='" & FecFin & "'" & FiltroFecIng & " AND ph.placod=a.placod AND PH.PROCESO='01')"
                Else
                    Cadena = Cadena & IIf(CInt(rs!sn_promedio) = 0, "I", "P") & rs!Campo & "=CASE WHEN dbo.fc_validapagos('" & wcia & "',a.placod,'" & FecIni & "','" & FecFin & "')>=" & nVecesPer & " THEN (SELECT ISNULL(SUM(" & rs!Formula & "),0) FROM plahistorico ph WHERE ph.cia='" & wcia & "' and ph.status!='*'  and  ph.FECHAPROCESO>='" & FecIni & "' " & _
                    "AND ph.FECHAPROCESO<='" & FecFin & "'" & FiltroFecIng & " AND ph.placod=a.placod AND PH.PROCESO='01')"
                End If
                
            Else
                Cadena = Cadena & IIf(CInt(rs!sn_promedio) = 0, "I", "P") & rs!Campo & "=SUM(COALESCE(" & Trim(rs!concepto) & Trim(rs!Campo) & ",0))"
            End If
            
            If CInt(rs!sn_promedio) = -1 And Trim(rs!concepto) = "A" Then
                Cadena = Cadena & "/(SELECT aportacion FROM PLACONSTANTE WHERE tipomovimiento='03' AND cia='" & wcia & "' AND status!='*' AND codinterno='" & rs!Campo & "'),"
            
            ElseIf CInt(rs!sn_promedio) = -1 And Trim(rs!concepto) = "I" And wGrupoPla <> "01" Then
                Cadena = Cadena & "/" & rs!factor & ","
            
            ElseIf CInt(rs!sn_promedio) = -1 And wGrupoPla = "01" Then
                'If Format(Cmbtipo.ItemData(Cmbtipo.ListIndex), "00") = "01" Then
                '    'EMPLEADOS
                '    Cadena = Cadena & "/" & dblFactor & " ELSE 0 END,"
                'Else
                '    'OBREROS
                    Cadena = Cadena & "/180" & " ELSE 0 END,"
                'End If
            Else
                Cadena = Cadena & ","
            End If
                

        Else
            Cadena = Cadena & "COALESCE((SELECT prb.importe/(factor_horas/" & hORAS_X_DIA & ") FROM plaremunbase prb WHERE cia='" & wcia & "' and placod=a.placod AND status!='*' and concepto='" & Trim(rs!Campo) & "'),0) as '" & Trim(rs!concepto) & Trim(rs!Campo) & "',"
        End If
        
        I = I + 1
        rs.MoveNext
    Loop
    
    'ArrReporte(i, MAXROW) = "I16"
    Cadena = Cadena & "0 AS I16"
    'i = i + 1
    'Cadena = Mid(Cadena, 1, Len(Trim(Cadena)) - 1)
    rs.Close
End If

Dim Fecha_CI As Date
Dim FechaCF As Date
Fecha_CI = DateSerial(Txtano.Text, (Cmbmes.ListIndex + 1) + 0, 1)
FechaCF = DateSerial(Txtano.Text, (Cmbmes.ListIndex + 1) + 1, 0)

sSQL = ""
sSQL = "SET DATEFORMAT DMY SELECT"

If wGrupoPla = "01" Then
    'IMPLEMENTACION GALLOS
    'NO SE ECEPTUA AL TRABAJADOR SI SU FECHA DE INGRESO ES EN LA FECHA DE PROCESO
    
    sSQL = sSQL & " a.placod,LTRIM(RTRIM(a.ap_pat))+' '+LTRIM(RTRIM(a.ap_mat))+' '+LTRIM(RTRIM(a.ap_cas))+' '+LTRIM(RTRIM(a.nom_1))+' '+LTRIM(RTRIM(a.nom_2)),"
    sSQL = sSQL & " fingreso , fcese, a.area,a.tipotrabajador,"
    sSQL = sSQL & Cadena
    sSQL = sSQL & " FROM planillas a LEFT OUTER JOIN plahistorico ph ON (ph.cia='" & wcia & "' and ph.status!='*' "
    sSQL = sSQL & " and  ph.FECHAPROCESO>='" & FecIni & "' AND ph.FECHAPROCESO<='" & FecFin & "' And ph.proceso='01' and ph.placod=a.placod)"
    
    sSQL = sSQL & " WHERE a.status!='*' AND (a.fcese is null OR  a.fcese>'" & FechaCF & "') "
    sSQL = sSQL & "AND a.tipotrabajador='" & VTipo & "' and a.cia='" & wcia & "' "
    sSQL = sSQL & "AND a.fingreso<='" & FechaCF & "' GROUP BY a.PLACOD , a.ap_pat, a.ap_mat, a.ap_cas, a.nom_1, a.nom_2, a.fingreso, a.fcese,a.area,a.tipotrabajador"

    'sSQL = sSQL & " WHERE a.status!='*' AND (a.fcese is null OR fcese>'" & "01/" & Format(Cmbmes.ListIndex + 1, "00") & "/" & Txtano.Text & "') "
    'sSQL = sSQL & "and a.tipotrabajador='" & VTipo & "' and a.cia='" & wcia & "' and (fcese >= '" & FechaCF & "' or fcese is null or fcese < '" & DateAdd("m", 1, "01/" & Format(Cmbmes.ListIndex + 1, "00") & "/" & Txtano.Text) & "')"
    'sSQL = sSQL & " And a.fingreso<='" & FechaCF & "' GROUP BY a.PLACOD , a.ap_pat, a.ap_mat, a.ap_cas, a.nom_1, a.nom_2, a.fingreso, a.fcese,a.area,a.tipotrabajador"



Else
    sSQL = sSQL & " a.placod,LTRIM(RTRIM(a.ap_pat))+' '+LTRIM(RTRIM(a.ap_mat))+' '+LTRIM(RTRIM(a.ap_cas))+' '+LTRIM(RTRIM(a.nom_1))+' '+LTRIM(RTRIM(a.nom_2)),"
    sSQL = sSQL & " fingreso , fcese, a.area,a.tipotrabajador,"
    sSQL = sSQL & Cadena
    sSQL = sSQL & " FROM planillas a LEFT OUTER JOIN plahistorico ph ON (ph.cia='" & wcia & "' and ph.status!='*' "
    sSQL = sSQL & " and  ph.FECHAPROCESO>='" & FecIni & "' AND ph.FECHAPROCESO<='" & FecFin & "' and ph.placod=a.placod)"
    sSQL = sSQL & " WHERE a.status!='*' AND (a.fcese is null OR fcese>'" & "01/" & Format(Cmbmes.ListIndex + 1, "00") & "/" & Txtano.Text & "') "
    sSQL = sSQL & "and fingreso <= '" & "01/" & Format(Cmbmes.ListIndex + 1, "00") & "/" & Txtano.Text & "' and a.tipotrabajador='" & VTipo & "' and a.cia='" & wcia & "' and (fcese >= '" & FechaCF & "' or fcese is null or fcese = '" & DateAdd("m", 1, "01/" & Format(Cmbmes.ListIndex + 1, "00") & "/" & Txtano.Text) & "')"
    'sSQL = sSQL & "and fingreso <= '" & "01/" & Format(Cmbmes.ListIndex + 1, "00") & "/" & Txtano.Text & "' and a.tipotrabajador='" & VTipo & "' and a.cia='" & wcia & "' and (fcese > '" & DateAdd("m", 1, "01/" & Format(Cmbmes.ListIndex + 1, "00") & "/" & Txtano.Text) & "' or fcese is null or fcese = '" & DateAdd("m", 1, "01/" & Format(Cmbmes.ListIndex + 1, "00") & "/" & Txtano.Text) & "')"
    sSQL = sSQL & " GROUP BY a.PLACOD , a.ap_pat, a.ap_mat, a.ap_cas, a.nom_1, a.nom_2, a.fingreso, a.fcese,a.area,a.tipotrabajador"

End If


 MAXROW = MAXROW + 1
 
 Dim dBasico As Double
 Dim dAportacion As Double
 Dim dAuxiliar As Double
 
If (fAbrRst(rs, sSQL)) Then
    Panelprogress.Caption = "Generando Provisión ..."
    Barra.Min = 0
    Barra.Max = rs.RecordCount
    Do While Not rs.EOF
        
        Barra.Value = rs.AbsolutePosition
       
        ReDim Preserve ArrReporte(0 To MAXCOL, 0 To MAXROW)
        'If Trim(rs!PlaCod) = "TO216" Then Stop
        If wGrupoPla <> "01" Then
            CantMes = CantidadMesesCalculo(rs!fIngreso)
        Else
            Dim CantDiasMes() As String
            CantDiasMes = CantidadDiasMesesCalculo(rs!fIngreso, FecProceso)
            CantDias = CantDiasMes(0)
            CantMes = CantDiasMes(1)
        End If
        
        ArrReporte(COL_CODIGO, MAXROW) = Trim(rs!PlaCod)
        ArrReporte(COL_FECING, MAXROW) = rs!fIngreso
        ArrReporte(COL_AREA, MAXROW) = rs!Area
        
        Factor_EsSalud = Trae_Porc_EsSalud_EPS(Trim(rs!PlaCod), Trim(rs!TipoTrabajador))
                
        dBasico = 0: dAportacion = 0: totaportes = 0: dAuxiliar = 0
                        
        For I = 6 To rs.Fields.count - 1
            Debug.Print "Calculando =>" & rs.Fields(I).Name
            sCol = BuscaColumna(rs.Fields(I).Name, MAXCOL)
            If sCol > 0 Then
                If Trim(rs!TipoTrabajador) = "01" Then
                    ArrReporte(sCol, MAXROW) = Round(rs.Fields(I).Value * DIAS_TRABAJO, 2)
                Else
                    ArrReporte(sCol, MAXROW) = Round(rs.Fields(I).Value, 2)
                End If
                If wGrupoPla <> "01" Then
                    If Left(rs.Fields(I).Name, 1) = "I" Then totaportes = totaportes + ArrReporte(sCol, MAXROW)
                Else
                    totaportes = totaportes + ArrReporte(sCol, MAXROW)
                End If
                
            End If
        Next
        
        
        For I = 3 To rs.Fields.count - 1
            If ArrReporte(I, 0) = "I16" Then
                sCol = I
                Exit For
            End If
        Next
        
        If sCol <> 0 Then
            ArrReporte(sCol, MAXROW) = IIf(wGrupoPla <> "01", Otros_Pagos(Trim(rs!PlaCod), FecProceso, "03"), 0)
            totaportes = totaportes + ArrReporte(sCol, MAXROW)
        End If
        
        I = MaxColTemp
        ArrReporte(I, 0) = "P01"
        
        If wGrupoPla <> "01" Then
            'factor_essalud SE CAMBIA POR CERO, PARA TOMAR EN CUENTA NUEVAMENTE REEMPLAZAR EL CERO POR LA VARIABLE
            'factor_essalud SE CAMBIO POR 0
            ArrReporte(I, MAXROW) = Round(totaportes * (Factor_EsSalud / 100), 2)
        Else
            'IMPLEMENTACION GALLOS
            ArrReporte(I, MAXROW) = 0
        End If
        
        I = I + 1
        ArrReporte(I, MAXROW) = Round(totaportes + ArrReporte(I - 1, MAXROW), 2)
        totaportes = ArrReporte(I, MAXROW)
        I = I + 1
        
        If Cmbmes.ListIndex + 1 = 1 Or Cmbmes.ListIndex + 1 = 7 Then
            'PROVTOTAL
            Sql = "select provtotal from plaprovgrati where cia='" & wcia & "' and month(fechaproceso)=" & Cmbmes.ListIndex + 1 & " AND YEAR(FECHAPROCESO)=  " & Txtano.Text & " and placod='" & Trim(rs!PlaCod) & "' and status<>'*'"
        Else
            'PROVTOTAL
            Sql = "select provtotal from plaprovgrati where cia='" & wcia & "' and month(fechaproceso)=" & Cmbmes.ListIndex & " AND YEAR(FECHAPROCESO)=  " & Txtano.Text & " and placod='" & Trim(rs!PlaCod) & "' and status<>'*'"
        End If
        
        If (fAbrRst(rsAux, Sql)) Then
            ArrReporte(I, MAXROW) = rsAux(0)
            rsAux.Close
        Else
            ArrReporte(I, MAXROW) = 0
        End If
        
        I = I + 1
        
        If wGrupoPla <> "01" Then
            
            If Trim(rs!TipoTrabajador) = "01" Then
                ArrReporte(I, MAXROW) = Round((totaportes / 6) * CantMes, 2)
            Else
                ArrReporte(I, MAXROW) = Round(((totaportes * 30) / 6) * CantMes, 2)
            End If
        Else
        
            If Trim(rs!TipoTrabajador) = "01" Then
                'EMPLEADOS
                ArrReporte(I, MAXROW) = Round((totaportes / 6) * CantMes + IIf(CantMes = 0, (totaportes / 6) / 30 * CantDias, 0), 2)
            Else
                'OBREROS
                ArrReporte(I, MAXROW) = Round(((totaportes * 30) / 6) * CantMes + IIf(CantMes = 0, totaportes / 6 * CantDias, 0), 2)
               
            End If
        
        
        End If
        
        I = I + 1
        ArrReporte(I, MAXROW) = Abs(ArrReporte(I - 1, MAXROW) - ArrReporte(I - 2, MAXROW))

        MAXROW = MAXROW + 1
        
        'RECORREMOS LOS CAMPOS DE INGRESO DE LA TABLA X(
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
        
        'RECORREMOS LOS CAMPOS DE PROMEDIO DE LA TABLA X(
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
        
        sSQL = ""
        sSQL = "SET DATEFORMAT DMY "
        sSQL = sSQL & "INSERT plaprovgrati VALUES ('" & wcia & "','" & ArrReporte(COL_CODIGO, MAXROW - 1) & "','" & Format(Cmbtipo.ItemData(Cmbtipo.ListIndex), "00") & "','',"
        sSQL = sSQL & " '','',''," & sSQLI & sSQLP & ArrReporte(MaxColTemp + 1, MAXROW - 1) & "," & ArrReporte(MaxColTemp + 2, MAXROW - 1) & ","
        sSQL = sSQL & ArrReporte(MaxColTemp + 2, MAXROW - 1) & "," & ArrReporte(MaxColTemp + 3, MAXROW - 1) & "," & ArrReporte(MaxColTemp + 4, MAXROW - 1) & ",'" & Format(FecProceso, "dd/mm/yyyy") & "',GETDATE(),'" & wuser & "','" & ArrReporte(COL_AREA, MAXROW - 1) & "',' ','" & Format(rs!fIngreso, "dd/mm/yyyy") & "')"
        
        
        cn.Execute (sSQL)
        
        
        totaportes = 0
        rs.MoveNext
    Loop
    Panelprogress.Caption = "Completo ..."
End If
    Carga_Prov_Grati
End Sub
Private Function CantidadMesesCalculo(ByVal pFecIngreso) As String
Dim mesestmp As String
If Year(pFecIngreso) < Txtano.Text Then
    If Cmbmes.ListIndex + 1 < 7 Then
        mesestmp = Cmbmes.ListIndex + 1
    Else
        mesestmp = (Cmbmes.ListIndex + 1) - 6
    End If
Else
    
    If Cmbmes.ListIndex + 1 >= Month(pFecIngreso) Then
        If Cmbmes.ListIndex + 1 < 7 Then
            mesestmp = Cmbmes.ListIndex + 1
            If Month(pFecIngreso) > 1 Then mesestmp = mesestmp - (Month(pFecIngreso) - 1)
            If Day(pFecIngreso) <> 1 Then mesestmp = mesestmp - 1
        Else
            mesestmp = Cmbmes.ListIndex + 1 - 6
            If Month(pFecIngreso) > 7 Then mesestmp = mesestmp - ((Month(pFecIngreso) - 6) - 1)
            If Month(pFecIngreso) > 6 Then
                If Day(pFecIngreso) <> 1 Then mesestmp = mesestmp - 1
            End If
        End If
    Else
        mesestmp = 0
    End If
    
End If
CantidadMesesCalculo = mesestmp

End Function
Private Function CantidadDiasMesesCalculo(ByVal pFecIngreso, ByVal pFecProceso) As String()
'IMPLEMENTACION GRUPO GALLOS

Dim DiasMeses(2) As String
Dim mdia As Integer
Dim mFec_Final As String


mdia = Ultimo_Dia(Month(pFecIngreso), Year(pFecIngreso))

'mdia = Ultimo_Dia(Month(CDate(pFecIngreso)), Year(CDate(pFecIngreso)))
'mdia = Ultimo_Dia(Val(Mid(CDate(pFecIngreso), 4, 2) - 1), Val(Mid(CDate(pFecIngreso), 7, 4)))

mFec_Final = Format(mdia, "00") & "/" & Format(Val(Mid(pFecIngreso, 4, 2)), "00") & "/" & Mid(pFecIngreso, 7, 4)

DiasMeses(0) = 0
DiasMeses(1) = 0
DiasMeses(2) = 0

If Month(pFecProceso) > 6 Then
    If Year(pFecIngreso) = Year(pFecProceso) Then
        'días
         DiasMeses(0) = IIf(Left(Format(pFecIngreso, "yyyymmdd"), 6) = Left(Format(pFecProceso, "yyyymmdd"), 6) And Day(pFecIngreso) <> 1, (DateDiff("d", pFecIngreso, mFec_Final) + IIf(Day(mFec_Final) <> 31, 1, 0)), IIf(Month(pFecIngreso) < 7 Or Day(pFecIngreso) = 1, 0, (DateDiff("d", pFecIngreso, mFec_Final) + IIf(Day(mFec_Final) <> 31, 1, 0))))
        'DiasMeses(0) = IIf(Left(Format(pFecIngreso, "yyyymmdd"), 6) = Left(Format(pFecProceso, "yyyymmdd"), 6), DateDiff("d", pFecIngreso, pFecProceso) + 1, IIf(Month(pFecIngreso) < 7, 0, (DateDiff("d", pFecIngreso, mFec_Final) + 1)))
        
        'meses
         DiasMeses(1) = IIf(Left(Format(pFecIngreso, "yyyymmdd"), 6) = Left(Format(pFecProceso, "yyyymmdd"), 6) And Day(pFecIngreso) = 1, 1, IIf(Month(pFecIngreso) < 7, Month(pFecProceso) - 6, Month(pFecProceso) - Month(pFecIngreso) + IIf(Day(pFecIngreso) = 1, 1, 0)))
        'DiasMeses(1) = IIf(Left(Format(pFecIngreso, "yyyymmdd"), 6) = Left(Format(pFecProceso, "yyyymmdd"), 6), 0, IIf(Month(pFecIngreso) < 7, Month(pFecProceso) - 6, Month(pFecProceso) - Month(pFecIngreso)))
        
        'En caso de que la f.ingreso sea = al inicio de mes en proceso
        If DiasMeses(1) = 0 And DiasMeses(0) >= mdia Then
            DiasMeses(0) = 0 'dias
            DiasMeses(1) = 1 'meses
        End If
        
    Else
        DiasMeses(0) = 0 'dias
        DiasMeses(1) = Month(pFecProceso) - 6
    End If
Else
    
    If Year(pFecIngreso) < Year(pFecProceso) Then
          DiasMeses(1) = Month(pFecProceso)
    ElseIf Year(pFecIngreso) = Year(pFecProceso) And Month(pFecIngreso) = 1 And Day(pFecIngreso) = 1 Then
          DiasMeses(1) = Month(pFecProceso)
    ElseIf Year(pFecIngreso) = Year(pFecProceso) And Month(pFecIngreso) = 6 And Day(pFecIngreso) = 1 Then
          DiasMeses(1) = 1
    ElseIf Year(pFecIngreso) = Year(pFecProceso) And Month(pFecIngreso) <> Month(pFecProceso) And Day(pFecIngreso) = 1 Then
        DiasMeses(1) = Month(pFecProceso) - Month(pFecIngreso) + 1
    ElseIf Year(pFecIngreso) = Year(pFecProceso) And Month(pFecIngreso) <> Month(pFecProceso) Then
           'días
            DiasMeses(0) = IIf(mdia = Day(pFecIngreso), 1, mdia - Day(pFecIngreso) + IIf(mdia <> 31, 1, 0))
            'meses
            DiasMeses(1) = Month(pFecProceso) - Month(pFecIngreso)
    ElseIf Year(pFecIngreso) = Year(pFecProceso) And Month(pFecIngreso) = Month(pFecProceso) And Day(pFecIngreso) = 1 Then
            DiasMeses(1) = 1
    ElseIf Year(pFecIngreso) = Year(pFecProceso) And Month(pFecIngreso) = Month(pFecProceso) And Day(pFecIngreso) <> 1 Then
            DiasMeses(0) = IIf(Day(pFecProceso) = Day(pFecIngreso), 1, (DateDiff("d", pFecIngreso, mFec_Final) + IIf(Day(mFec_Final) <> 31, 1, 0)))
     End If
     
    'En caso de que la f.ingreso sea = al inicio de mes en proceso
    
    If DiasMeses(1) = 0 And DiasMeses(0) >= mdia Then
        DiasMeses(0) = 0 'dias
        DiasMeses(1) = 1 'meses
    End If

End If

CantidadDiasMesesCalculo = DiasMeses



End Function
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
String, TipoTrabajador As String, CentroCosto As String, Area As String)
'    If Verifica_Existencia_Registro(CodCompañia, CodTrabajador, Año_Proceso, MesSeleccion, _
 '   SemProceso, Voucher, CtaContable, TipoBoleta) = False Then
        Select Case Opcion
            Case 1
                Call Graba_Asiento_Contable(CodCompañia, CodTrabajador, Año_Proceso, MesSeleccion, _
                SemProceso, Voucher, CtaContable, TipoAsiento, DescCta, 0, MontoInt, TipoBoleta, _
                TipoTrabajador, CentroCosto, Area)
            Case 2
                Call Graba_Asiento_Contable(CodCompañia, CodTrabajador, Año_Proceso, MesSeleccion, _
                SemProceso, Voucher, CtaContable, TipoAsiento, DescCta, MontoInt, 0, TipoBoleta, _
                TipoTrabajador, CentroCosto, Area)
        End Select
  '  End If
End Sub
Sub Codigo_Empresa_Starsoft()
'    'Call Recuperar_Codigo_Empresa_Starsoft(wcia)
'
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
Private Sub ReporteVaca()

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

Dim MArea  As String

Set xlApp1 = CreateObject("Excel.Application")
xlApp1.Workbooks.Add
Set xlApp2 = xlApp1.Application
Set xlBook = xlApp2.Workbooks(1)
Set xlSheet = xlApp2.Worksheets("HOJA1")

xlSheet.Range("A:A").ColumnWidth = 7
xlSheet.Range("B:B").ColumnWidth = 40
xlSheet.Range("C:C").ColumnWidth = 9.71
xlSheet.Range("H:Z").ColumnWidth = 14
xlSheet.Range("H:Z").NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "

xlSheet.Cells(1, 1).Value = Trim(Cmbcia.Text)
xlSheet.Cells(1, 1).Font.Bold = True
xlSheet.Cells(2, 2).Value = " REPORTE DE PROVISION DE VACACIONES " & Cmbtipo.Text & " - " & Cmbmes.Text & " " & Txtano.Text
xlSheet.Cells(2, 2).Font.Bold = True
xlSheet.Cells(2, 2).Font.Size = 12
xlSheet.Cells(2, 2).HorizontalAlignment = xlCenter

Fila = 4
xlSheet.Cells(Fila, 1).Value = "Codigo"
xlSheet.Cells(Fila, 2).Value = "Nombre Trabajador"
xlSheet.Cells(Fila, 3).Value = "DNI"
xlSheet.Cells(Fila, 4).Value = "F.Ing"

xlSheet.Cells(Fila, 5).Value = "Record Año"
xlSheet.Cells(Fila + 1, 5).Value = "Mes Dia"
xlSheet.Cells(Fila, 6).Value = "Vacaciones"
xlSheet.Cells(Fila + 1, 6).Value = "Tomadas"

xlSheet.Cells(Fila, 7).Value = "Record Año"
xlSheet.Cells(Fila + 1, 7).Value = "Actual"

xlSheet.Cells(Fila, 8).Value = "Vacaciones"
xlSheet.Cells(Fila + 1, 8).Value = "Perdidas"

xlSheet.Cells(Fila, 9).Value = "Record Acum"
xlSheet.Cells(Fila + 1, 9).Value = "Año Mes Dia"

Sql = "usp_Carga_Reporte_Vaca '" & wcia & "'," & Txtano.Text & "," & Cmbmes.ListIndex + 1 & ",'" & VTipo & "'"
If (fAbrRst(rs, Sql)) Then rs.MoveFirst

MArea = ""
Dim lCol As Integer
Dim rsTit As ADODB.Recordset
lCol = 1
If rs.RecordCount > 0 Then lCol = rs!nCol
For I = 1 To lCol
    
   Sql = "select des_cts from Pla_Cts_Titulos where cia='" & wcia & "' and codinterno='" & Mid(rs.Fields(I + 5).Name, 2, 2) & "'"
   If (fAbrRst(rsTit, Sql)) Then
      If I = 2 Then
         xlSheet.Cells(Fila, I + 9).Value = "AFP"
      Else
         xlSheet.Cells(Fila, I + 9).Value = Trim(rsTit(0) & "")
      End If
   Else
      xlSheet.Cells(Fila, I + 9).Value = UCase(rs.Fields(I + 5).Name)
   End If
   
   xlSheet.Range(xlSheet.Cells(Fila, I + 9), xlSheet.Cells(Fila + 1, I + 9)).Merge
   xlSheet.Range(xlSheet.Cells(Fila, I + 9), xlSheet.Cells(Fila + 1, I + 9)).WrapText = True
   xlSheet.Range(xlSheet.Cells(Fila, I + 9), xlSheet.Cells(Fila + 1, I + 9)).VerticalAlignment = xlTop
   xlSheet.Range(xlSheet.Cells(Fila, I + 9), xlSheet.Cells(Fila + 1, I + 9)).HorizontalAlignment = xlCenter
   rsTit.Close
Next
xlSheet.Cells(Fila, I + 9).Value = "Remun"
xlSheet.Cells(Fila + 1, I + 9).Value = "Vaca"
xlSheet.Cells(Fila, I + 10).Value = "Mes"
xlSheet.Cells(Fila + 1, I + 10).Value = "Anterior"
xlSheet.Cells(Fila, I + 11).Value = "Vaca"
xlSheet.Cells(Fila + 1, I + 11).Value = "Pagada"
xlSheet.Cells(Fila, I + 12).Value = "Vaca"
xlSheet.Cells(Fila + 1, I + 12).Value = "Perdida"
xlSheet.Cells(Fila, I + 13).Value = "Vacac."
xlSheet.Cells(Fila + 1, I + 13).Value = "Por Pagar"
xlSheet.Cells(Fila, I + 14).Value = "Prov."
xlSheet.Cells(Fila + 1, I + 14).Value = "Del Mes"

xlSheet.Range(xlSheet.Cells(Fila, I + 14), xlSheet.Cells(Fila + 1, I + 14)).HorizontalAlignment = xlCenter

Dim sLin As Integer
sLin = I + 14
If sLin < 17 Then sLin = 17

xlSheet.Range(xlSheet.Cells(4, 1), xlSheet.Cells(4, sLin)).Borders(xlEdgeTop).LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(5, 1), xlSheet.Cells(5, sLin)).Borders(xlEdgeBottom).LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(4, 1), xlSheet.Cells(5, I + sLin)).HorizontalAlignment = xlCenter

xlSheet.Range(xlSheet.Cells(2, 1), xlSheet.Cells(2, I + 14)).Merge

Fila = 7
Dim lNumTrab As Integer
lNumTrab = 0
Do While Not rs.EOF
   If MArea <> Trim(rs!ccosto) Then
      If MArea <> "" Then
         rs.MovePrevious
         xlSheet.Cells(Fila, I + 9).Value = rs!remtotal
         xlSheet.Cells(Fila, I + 9).Borders(xlEdgeTop).LineStyle = xlContinuous
         xlSheet.Cells(Fila, I + 10).Value = rs!provmesant
         xlSheet.Cells(Fila, I + 10).Borders(xlEdgeTop).LineStyle = xlContinuous
         xlSheet.Cells(Fila, I + 11).Value = rs!provpagadas
         xlSheet.Cells(Fila, I + 11).Borders(xlEdgeTop).LineStyle = xlContinuous
         xlSheet.Cells(Fila, I + 12).Value = rs!provperdida
         xlSheet.Cells(Fila, I + 12).Borders(xlEdgeTop).LineStyle = xlContinuous
         xlSheet.Cells(Fila, I + 13).Value = rs!provtotal
         xlSheet.Cells(Fila, I + 13).Borders(xlEdgeTop).LineStyle = xlContinuous
         xlSheet.Cells(Fila, I + 14).Value = rs!provmes
         xlSheet.Cells(Fila, I + 14).Borders(xlEdgeTop).LineStyle = xlContinuous
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
      xlSheet.Cells(Fila, 3).Value = "'" & Trim(rs!nro_doc)
      xlSheet.Cells(Fila, 4).Value = rs!fIngreso
      
      xlSheet.Cells(Fila, 5).Value = rs!recordanoant
      xlSheet.Cells(Fila, 6).Value = rs!recordvaca
      xlSheet.Cells(Fila, 7).Value = rs!recordanoact
      xlSheet.Cells(Fila, 8).Value = rs!recordperdida
      xlSheet.Cells(Fila, 9).Value = rs!recordacu
      lNumTrab = lNumTrab + 1
      
      For I = 1 To lCol
         xlSheet.Cells(Fila, I + 9).Value = rs(I + 5)
      Next
      xlSheet.Cells(Fila, I + 9).Value = rs!remtotal
      xlSheet.Cells(Fila, I + 10).Value = rs!provmesant
      xlSheet.Cells(Fila, I + 11).Value = rs!provpagadas
      xlSheet.Cells(Fila, I + 12).Value = rs!provperdida
      xlSheet.Cells(Fila, I + 13).Value = rs!provtotal
      xlSheet.Cells(Fila, I + 14).Value = rs!provmes
      Fila = Fila + 1
   ElseIf Trim(rs!ccosto) = "ZZZZZ" Then
      For I = 1 To lCol
         xlSheet.Cells(Fila, I + 9).Value = rs(I + 5)
      Next
      xlSheet.Cells(Fila, I + 9).Value = rs!remtotal
      xlSheet.Cells(Fila, I + 10).Value = rs!provmesant
      xlSheet.Cells(Fila, I + 11).Value = rs!provpagadas
      xlSheet.Cells(Fila, I + 12).Value = rs!provperdida
      xlSheet.Cells(Fila, I + 13).Value = rs!provtotal
      xlSheet.Cells(Fila, I + 14).Value = rs!provmes
      xlSheet.Range(xlSheet.Cells(Fila, I + 4), xlSheet.Cells(Fila, sLin)).Borders(xlEdgeTop).LineStyle = xlContinuous
      xlSheet.Range(xlSheet.Cells(Fila, I + 4), xlSheet.Cells(Fila, sLin)).Borders(xlEdgeBottom).LineStyle = xlContinuous
   End If
   rs.MoveNext
Loop
Fila = Fila + 1
xlSheet.Cells(Fila, 2).Value = "NUMERO DE TRABAJADORES"
xlSheet.Cells(Fila, 3).Value = lNumTrab

xlApp2.Application.ActiveWindow.DisplayGridLines = False
xlSheet.Range("A1:A1").Select
xlApp2.Application.Caption = "PROVISION DE VACACIONES"
xlApp2.ActiveWindow.Zoom = 80
xlApp2.Application.Visible = True

End Sub
Private Sub Reporte_Provision_Grati()
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

Dim MArea  As String

Set xlApp1 = CreateObject("Excel.Application")
xlApp1.Workbooks.Add
Set xlApp2 = xlApp1.Application
Set xlBook = xlApp2.Workbooks(1)
Set xlSheet = xlApp2.Worksheets("HOJA1")

xlSheet.Range("A:A").ColumnWidth = 7
xlSheet.Range("B:B").ColumnWidth = 40
xlSheet.Range("C:C").ColumnWidth = 9.71
xlSheet.Range("H:Z").ColumnWidth = 14
xlSheet.Range("F:Z").NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
xlSheet.Range("E:E").NumberFormat = "#,##0_ ;[Red]-#,##0 "

xlSheet.Cells(1, 1).Value = Trim(Cmbcia.Text)
xlSheet.Cells(1, 1).Font.Bold = True
xlSheet.Cells(2, 2).Value = " REPORTE DE PROVISION DE GRATIFICACIONES " & Cmbtipo.Text & " - " & Cmbmes.Text & " " & Txtano.Text
xlSheet.Cells(2, 2).Font.Bold = True
xlSheet.Cells(2, 2).Font.Size = 12
xlSheet.Cells(2, 2).HorizontalAlignment = xlCenter

Fila = 4
xlSheet.Cells(Fila, 1).Value = "Codigo"
xlSheet.Cells(Fila, 2).Value = "Nombre Trabajador"
xlSheet.Cells(Fila, 3).Value = "DNI"
xlSheet.Cells(Fila, 4).Value = "F.Ing"

xlSheet.Cells(Fila, 5).Value = "Meses"

Sql = "usp_Carga_Reporte_Grati '" & wcia & "'," & Txtano.Text & "," & Cmbmes.ListIndex + 1 & ",'" & VTipo & "'"
If (fAbrRst(rs, Sql)) Then rs.MoveFirst

MArea = ""
Dim lCol As Integer
Dim rsTit As ADODB.Recordset
lCol = 1
If rs.RecordCount > 0 Then lCol = rs!nCol
For I = 1 To lCol
    
   Sql = "select des_cts from Pla_Cts_Titulos where cia='" & wcia & "' and codinterno='" & Mid(rs.Fields(I + 5).Name, 2, 2) & "'"
   If (fAbrRst(rsTit, Sql)) Then
      If I = 2 Then
         xlSheet.Cells(Fila, I + 5).Value = "AFP"
      Else
         xlSheet.Cells(Fila, I + 5).Value = Trim(rsTit(0) & "")
      End If
   Else
      xlSheet.Cells(Fila, I + 5).Value = UCase(rs.Fields(I + 5).Name)
   End If
   
   xlSheet.Range(xlSheet.Cells(Fila, I + 5), xlSheet.Cells(Fila + 1, I + 5)).Merge
   xlSheet.Range(xlSheet.Cells(Fila, I + 5), xlSheet.Cells(Fila + 1, I + 5)).WrapText = True
   xlSheet.Range(xlSheet.Cells(Fila, I + 5), xlSheet.Cells(Fila + 1, I + 5)).VerticalAlignment = xlTop
   xlSheet.Range(xlSheet.Cells(Fila, I + 5), xlSheet.Cells(Fila + 1, I + 5)).HorizontalAlignment = xlCenter
   rsTit.Close
Next
xlSheet.Cells(Fila, I + 5).Value = "Aporta"
xlSheet.Cells(Fila + 1, I + 5).Value = "ción"
xlSheet.Cells(Fila, I + 6).Value = "Remun"
xlSheet.Cells(Fila + 1, I + 6).Value = "Grati."
xlSheet.Cells(Fila, I + 7).Value = "Mes"
xlSheet.Cells(Fila + 1, I + 7).Value = "Anterior"
xlSheet.Cells(Fila, I + 8).Value = "Gratif."
xlSheet.Cells(Fila + 1, I + 8).Value = "Por PAgar"
xlSheet.Cells(Fila, I + 9).Value = "Prov."
xlSheet.Cells(Fila + 1, I + 9).Value = "Del Mes"

xlSheet.Range(xlSheet.Cells(Fila, I + 9), xlSheet.Cells(Fila + 1, I + 9)).HorizontalAlignment = xlCenter

Dim sLin As Integer
sLin = I + 9
If sLin < 14 Then sLin = 14

xlSheet.Range(xlSheet.Cells(4, 1), xlSheet.Cells(4, sLin)).Borders(xlEdgeTop).LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(5, 1), xlSheet.Cells(5, sLin)).Borders(xlEdgeBottom).LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(4, 1), xlSheet.Cells(5, I + sLin)).HorizontalAlignment = xlCenter

xlSheet.Range(xlSheet.Cells(2, 1), xlSheet.Cells(2, I + 9)).Merge

Fila = 7
Dim lNumTrab As Integer
lNumTrab = 0
Do While Not rs.EOF
   If MArea <> Trim(rs!ccosto) Then
      If MArea <> "" Then
         rs.MovePrevious
         xlSheet.Cells(Fila, I + 6).Value = rs!remtotal
         xlSheet.Cells(Fila, I + 6).Borders(xlEdgeTop).LineStyle = xlContinuous
         xlSheet.Cells(Fila, I + 7).Value = rs!gratmesant
         xlSheet.Cells(Fila, I + 7).Borders(xlEdgeTop).LineStyle = xlContinuous
         xlSheet.Cells(Fila, I + 8).Value = rs!provtotal
         xlSheet.Cells(Fila, I + 8).Borders(xlEdgeTop).LineStyle = xlContinuous
         xlSheet.Cells(Fila, I + 9).Value = rs!gratmes
         xlSheet.Cells(Fila, I + 9).Borders(xlEdgeTop).LineStyle = xlContinuous
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
      xlSheet.Cells(Fila, 3).Value = "'" & Trim(rs!nro_doc)
      xlSheet.Cells(Fila, 4).Value = rs!fIngreso
      xlSheet.Cells(Fila, 5).Value = rs!provpagadas
      lNumTrab = lNumTrab + 1
      
      For I = 1 To lCol
         xlSheet.Cells(Fila, I + 5).Value = rs(I + 5)
      Next
         xlSheet.Cells(Fila, I + 5).Value = rs!aporta
         xlSheet.Cells(Fila, I + 6).Value = rs!remtotal
         xlSheet.Cells(Fila, I + 7).Value = rs!gratmesant
         xlSheet.Cells(Fila, I + 8).Value = rs!provtotal
         xlSheet.Cells(Fila, I + 9).Value = rs!gratmes
         Fila = Fila + 1
   ElseIf Trim(rs!ccosto) = "ZZZZZ" Then
      For I = 1 To lCol
         xlSheet.Cells(Fila, I + 5).Value = rs(I + 5)
      Next
      xlSheet.Cells(Fila, I + 5).Value = rs!aporta
      xlSheet.Cells(Fila, I + 6).Value = rs!remtotal
      xlSheet.Cells(Fila, I + 7).Value = rs!gratmesant
      xlSheet.Cells(Fila, I + 8).Value = rs!provtotal
      xlSheet.Cells(Fila, I + 9).Value = rs!gratmes
      xlSheet.Range(xlSheet.Cells(Fila, I), xlSheet.Cells(Fila, sLin)).Borders(xlEdgeTop).LineStyle = xlContinuous
      xlSheet.Range(xlSheet.Cells(Fila, I), xlSheet.Cells(Fila, sLin)).Borders(xlEdgeBottom).LineStyle = xlContinuous
   End If
   rs.MoveNext
Loop
Fila = Fila + 1
xlSheet.Cells(Fila, 2).Value = "NUMERO DE TRABAJADORES"
xlSheet.Cells(Fila, 3).Value = lNumTrab

xlApp2.Application.ActiveWindow.DisplayGridLines = False
xlSheet.Range("A1:A1").Select
xlApp2.Application.Caption = "PROVISION DE GRATIFICACIONES"
xlApp2.ActiveWindow.Zoom = 80
xlApp2.Application.Visible = True

End Sub

