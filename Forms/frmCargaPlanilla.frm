VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmCargaPlanilla 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Carga Automática de Planillas"
   ClientHeight    =   8160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7755
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8160
   ScaleWidth      =   7755
   Begin VB.ListBox LstObs 
      Appearance      =   0  'Flat
      Height          =   1005
      Left            =   240
      TabIndex        =   30
      Top             =   6840
      Width           =   7605
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   375
      Left            =   135
      TabIndex        =   27
      Top             =   90
      Width           =   7665
      Begin VB.ComboBox Cmbcia 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1185
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   0
         Width           =   6375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Compañia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   0
         TabIndex        =   28
         Top             =   60
         Width           =   840
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Datos de la Boleta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1740
      Left            =   120
      TabIndex        =   20
      Top             =   960
      Width           =   7560
      Begin VB.CheckBox asig_fam 
         Caption         =   "Asf.Fam."
         Height          =   255
         Left            =   3600
         TabIndex        =   40
         Top             =   720
         Width           =   975
      End
      Begin MSComCtl2.DTPicker FechPerVac 
         Height          =   315
         Left            =   6120
         TabIndex        =   9
         Top             =   1200
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CalendarTitleBackColor=   8388608
         CalendarTitleForeColor=   16777215
         Format          =   16842753
         CurrentDate     =   39609
      End
      Begin VB.ComboBox Cmbturno 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5280
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   720
         Width           =   2175
      End
      Begin VB.ComboBox Cmbtipo 
         Height          =   315
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   3015
      End
      Begin VB.TextBox Txtsemana 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2880
         TabIndex        =   5
         Top             =   735
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.ComboBox Cmbtipotrabajador 
         Height          =   315
         ItemData        =   "frmCargaPlanilla.frx":0000
         Left            =   5250
         List            =   "frmCargaPlanilla.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   2175
      End
      Begin MSComCtl2.DTPicker Cmbfecha 
         Height          =   315
         Left            =   720
         TabIndex        =   3
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   16842753
         CurrentDate     =   39456
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   315
         Left            =   3240
         TabIndex        =   4
         Top             =   720
         Visible         =   0   'False
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.DTPicker Cmbal 
         Height          =   315
         Left            =   2880
         TabIndex        =   8
         Top             =   1200
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CalendarTitleBackColor=   8388608
         CalendarTitleForeColor=   16777215
         Format          =   16842753
         CurrentDate     =   39456
      End
      Begin MSComCtl2.DTPicker Cmbdel 
         Height          =   315
         Left            =   720
         TabIndex        =   7
         Top             =   1200
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CalendarTitleBackColor=   8388608
         CalendarTitleForeColor=   16777215
         Format          =   16842753
         CurrentDate     =   39458
      End
      Begin VB.Label LblFecPerVac 
         Caption         =   "Fecha Periodo Vacacional"
         Height          =   435
         Left            =   4440
         TabIndex        =   31
         Top             =   1140
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Turno"
         Height          =   195
         Left            =   4800
         TabIndex        =   29
         Top             =   720
         Width           =   420
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Boleta"
         Height          =   435
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   675
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha"
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   780
         Width           =   450
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Semana"
         Height          =   195
         Left            =   2160
         TabIndex        =   24
         Top             =   780
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Del"
         Height          =   195
         Left            =   240
         TabIndex        =   23
         Top             =   1260
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Al"
         Height          =   195
         Left            =   2160
         TabIndex        =   22
         Top             =   1260
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Trabajador"
         Height          =   195
         Left            =   4095
         TabIndex        =   21
         Top             =   240
         Width           =   1125
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Contenido del Archivo a importar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   2250
      Left            =   120
      TabIndex        =   17
      Top             =   3960
      Width           =   7530
      Begin TrueOleDBGrid70.TDBGrid DGrd 
         Height          =   1980
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   7305
         _ExtentX        =   12885
         _ExtentY        =   3493
         _LayoutType     =   0
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   2
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).FetchRowStyle=   -1  'True
         Splits(0).DividerColor=   13160660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=2"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
         Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         ColumnFooters   =   -1  'True
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         MultipleLines   =   0
         CellTipsWidth   =   0
         DeadAreaBackColor=   13160660
         RowDividerColor =   13160660
         RowSubDividerColor=   13160660
         DirectionAfterEnter=   1
         MaxRows         =   250000
         ViewColumnCaptionWidth=   0
         ViewColumnWidth =   0
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
         _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bgcolor=&HD7D7D7&,.bold=0"
         _StyleDefs(14)  =   ":id=3,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
         _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41,.bgcolor=&HFF8000&"
         _StyleDefs(23)  =   ":id=11,.appearance=0"
         _StyleDefs(24)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(25)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(26)  =   "Splits(0).CaptionStyle:id=22,.parent=4,.bgcolor=&HFF8000&"
         _StyleDefs(27)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.bgcolor=&HFF8000&,.fgcolor=&HFFFFFF&"
         _StyleDefs(28)  =   ":id=14,.bold=0,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(29)  =   ":id=14,.fontname=MS Sans Serif"
         _StyleDefs(30)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(31)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(32)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(33)  =   "Splits(0).EditorStyle:id=17,.parent=7,.bgcolor=&HFF8000&"
         _StyleDefs(34)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(35)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(36)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(37)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(38)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(39)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(40)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(41)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(42)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(43)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(44)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(45)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(46)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(47)  =   "Named:id=33:Normal"
         _StyleDefs(48)  =   ":id=33,.parent=0"
         _StyleDefs(49)  =   "Named:id=34:Heading"
         _StyleDefs(50)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(51)  =   ":id=34,.wraptext=-1"
         _StyleDefs(52)  =   "Named:id=35:Footing"
         _StyleDefs(53)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(54)  =   "Named:id=36:Selected"
         _StyleDefs(55)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(56)  =   "Named:id=37:Caption"
         _StyleDefs(57)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(58)  =   "Named:id=38:HighlightRow"
         _StyleDefs(59)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(60)  =   "Named:id=39:EvenRow"
         _StyleDefs(61)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(62)  =   "Named:id=40:OddRow"
         _StyleDefs(63)  =   ":id=40,.parent=33"
         _StyleDefs(64)  =   "Named:id=41:RecordSelector"
         _StyleDefs(65)  =   ":id=41,.parent=34"
         _StyleDefs(66)  =   "Named:id=42:FilterBar"
         _StyleDefs(67)  =   ":id=42,.parent=33"
      End
   End
   Begin VB.CommandButton cmdImportar 
      Appearance      =   0  'Flat
      Caption         =   "Importar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5280
      TabIndex        =   13
      Top             =   6435
      Width           =   2355
   End
   Begin VB.Frame Frame2 
      Caption         =   "Origen del Archivo a importar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1155
      Left            =   120
      TabIndex        =   14
      Top             =   2760
      Width           =   7605
      Begin VB.OptionButton Quincena 
         Caption         =   "2 Da. Qna."
         Height          =   375
         Index           =   1
         Left            =   6120
         TabIndex        =   42
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton Quincena 
         Caption         =   "1 Era. Qna."
         Height          =   375
         Index           =   0
         Left            =   5160
         TabIndex        =   41
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox TxtRango 
         Height          =   315
         Left            =   3000
         TabIndex        =   10
         Text            =   "A6:U95"
         Top             =   330
         Width           =   1095
      End
      Begin VB.TextBox Txtarchivos 
         BackColor       =   &H8000000A&
         Height          =   285
         Left            =   1170
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   705
         Width           =   5370
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   495
         Left            =   6840
         TabIndex        =   11
         Top             =   600
         Width           =   615
         _Version        =   65536
         _ExtentX        =   1085
         _ExtentY        =   873
         _StockProps     =   78
         Picture         =   "frmCargaPlanilla.frx":0004
      End
      Begin VB.Label Label10 
         Caption         =   "Ejm: A1:G45"
         Height          =   255
         Left            =   4200
         TabIndex        =   36
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Rango de Datos Hoja de Excel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   35
         Top             =   360
         Width           =   2745
      End
      Begin VB.Label Label1 
         Caption         =   "Ubicación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   135
         TabIndex        =   16
         Top             =   720
         Width           =   945
      End
   End
   Begin MSAdodcLib.Adodc AdoData 
      Height          =   420
      Left            =   1260
      Top             =   1035
      Visible         =   0   'False
      Width           =   3240
      _ExtentX        =   5715
      _ExtentY        =   741
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   $"frmCargaPlanilla.frx":031E
      OLEDBString     =   $"frmCargaPlanilla.frx":0499
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "AdoData"
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
   Begin MSComDlg.CommonDialog Box 
      Left            =   225
      Top             =   990
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar Barra 
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   6480
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.CommandButton CmdSeteo 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Seteo de Campos equivalentes"
      Height          =   315
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   530
      Width           =   2655
   End
   Begin VB.Frame FraSeteo 
      Caption         =   "Seteo de Campos equivalentes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   7095
      Left            =   120
      TabIndex        =   34
      Top             =   960
      Visible         =   0   'False
      Width           =   7575
      Begin VB.CommandButton CmdSalir 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Salir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3840
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   6600
         Width           =   2415
      End
      Begin VB.CommandButton CmdGrabar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Grabar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   6600
         Width           =   2415
      End
      Begin TrueOleDBGrid70.TDBGrid DgrdSeteo 
         Height          =   6060
         Left            =   120
         TabIndex        =   37
         Top             =   240
         Width           =   7305
         _ExtentX        =   12885
         _ExtentY        =   10689
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Nombre Columna Archivo Externo"
         Columns(0).DataField=   "campo_xls"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   49
         Columns(1)._MaxComboItems=   10
         Columns(1).Caption=   "Nombre Columna tabla Plahistorico"
         Columns(1).DataField=   "campo_sql"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Descripción Concepto Plahistorico"
         Columns(2).DataField=   "descripcion"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   3
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).FetchRowStyle=   -1  'True
         Splits(0).DividerColor=   13160660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=3"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
         Splits(0)._ColumnProps(8)=   "Column(1).Button=1"
         Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(10)=   "Column(1).AutoDropDown=1"
         Splits(0)._ColumnProps(11)=   "Column(1).DropDownList=1"
         Splits(0)._ColumnProps(12)=   "Column(1).AutoCompletion=1"
         Splits(0)._ColumnProps(13)=   "Column(2).Width=6641"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=6562"
         Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowDelete     =   -1  'True
         AllowAddNew     =   -1  'True
         ColumnFooters   =   -1  'True
         DefColWidth     =   0
         HeadLines       =   3
         FootLines       =   1
         MultipleLines   =   0
         CellTipsWidth   =   0
         DeadAreaBackColor=   13160660
         RowDividerColor =   13160660
         RowSubDividerColor=   13160660
         DirectionAfterEnter=   1
         MaxRows         =   250000
         ViewColumnCaptionWidth=   0
         ViewColumnWidth =   0
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
         _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bgcolor=&HD7D7D7&,.bold=0"
         _StyleDefs(14)  =   ":id=3,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
         _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41,.bgcolor=&HFF8000&"
         _StyleDefs(23)  =   ":id=11,.appearance=0"
         _StyleDefs(24)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(25)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(26)  =   "Splits(0).CaptionStyle:id=22,.parent=4,.bgcolor=&HFF8000&"
         _StyleDefs(27)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.bgcolor=&HFF8000&,.fgcolor=&HFFFFFF&"
         _StyleDefs(28)  =   ":id=14,.bold=0,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(29)  =   ":id=14,.fontname=MS Sans Serif"
         _StyleDefs(30)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(31)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(32)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(33)  =   "Splits(0).EditorStyle:id=17,.parent=7,.bgcolor=&HFF8000&"
         _StyleDefs(34)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(35)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(36)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(37)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(38)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(39)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(40)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(41)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(42)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(43)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(44)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(45)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(46)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(47)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
         _StyleDefs(48)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(49)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(50)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(51)  =   "Named:id=33:Normal"
         _StyleDefs(52)  =   ":id=33,.parent=0"
         _StyleDefs(53)  =   "Named:id=34:Heading"
         _StyleDefs(54)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(55)  =   ":id=34,.wraptext=-1"
         _StyleDefs(56)  =   "Named:id=35:Footing"
         _StyleDefs(57)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(58)  =   "Named:id=36:Selected"
         _StyleDefs(59)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(60)  =   "Named:id=37:Caption"
         _StyleDefs(61)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(62)  =   "Named:id=38:HighlightRow"
         _StyleDefs(63)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(64)  =   "Named:id=39:EvenRow"
         _StyleDefs(65)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(66)  =   "Named:id=40:OddRow"
         _StyleDefs(67)  =   ":id=40,.parent=33"
         _StyleDefs(68)  =   "Named:id=41:RecordSelector"
         _StyleDefs(69)  =   ":id=41,.parent=34"
         _StyleDefs(70)  =   "Named:id=42:FilterBar"
         _StyleDefs(71)  =   ":id=42,.parent=33"
      End
   End
   Begin VB.Label Label9 
      BackColor       =   &H00800000&
      Caption         =   "  Carga Automática de Planillas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   32
      Top             =   480
      Width           =   7815
   End
   Begin VB.Label Label11 
      Caption         =   "Progreso de la Transferencia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   180
      TabIndex        =   18
      Top             =   6240
      Width           =   2595
   End
End
Attribute VB_Name = "frmCargaPlanilla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim VTipo As String
Dim VTipotrab As String
Dim VTipoPago As String
Dim VHorasBol As Integer
Dim VTurno As String
Dim Mstatus As String
Dim contador As Integer

Dim rsExport As New ADODB.Recordset
Dim rshoras As New Recordset
Dim rsCosto As New Recordset
Dim rspagadic As New Recordset
Dim rsdesadic As New Recordset
Dim RsCampoPlaHist As New ADODB.Recordset
Dim RsSeteo As New ADODB.Recordset
Dim mPrefijo As String

Public Sub Importar_Excel()

    'Referencia a la instancia de excel
    
    Dim xlApp2 As Excel.Application
    Dim xlApp1  As Excel.Application
    Dim xLibro  As Excel.Workbook
        
    On Error Resume Next
        
    'Chequeamos si excel esta corriendo
        
    Set xlApp1 = GetObject(, "Excel.Application")
    If xlApp1 Is Nothing Then
        'Si excel no esta corriendo, creamos una nueva instancia.
        Set xlApp1 = CreateObject("Excel.Application")
    End If
        
    'ACP On Error GoTo 0
    
    On Error GoTo ERR
    
    
    'Variable de tipo Aplicación de Excel
    
    Set xlApp2 = xlApp1.Application
    
    'Una variable de tipo Libro de Excel
    
    'Set xLibro = xlApp2.Workbooks(1)

    Dim Col As Integer, Fila As Integer
  
  
    'Usamos el método open para abrir el archivo que está _
     en el directorio del programa llamado archivo.xls
    
    Set xLibro = xlApp2.Workbooks.Open(Txtarchivos.Text)
  
    'Hacemos el Excel Invisible
    
    xlApp2.Visible = False
    
    Dim xTipoTrab As String
    xTipoTrab = fc_CodigoComboBox(Cmbtipotrabajador, 2)
    
    If xTipoTrab = "02" Then 'obrero verifica titulo de la semana
        Dim CadSem As String
        CadSem = xLibro.Sheets(1).Cells(2, 1).Value
        Dim pos As Integer
        pos = InStr(1, CadSem, Me.Txtsemana.Text)
        If pos = 0 Then
            MsgBox "El contenido del archivo no pertences a la Semana " & Me.Txtsemana.Text, vbCritical, Me.Caption
            GoTo Salir:
        End If
        
    End If
    
  
    'Eliminamos los objetos si ya no los usamos
    
    If Not xlApp2 Is Nothing Then Set xlApp2 = Nothing
    If Not xlApp1 Is Nothing Then Set xlApp1 = Nothing
    If Not xLibro Is Nothing Then Set xlBook = Nothing
    
    Dim conexion As ADODB.Connection, rs As ADODB.Recordset
  
    Set conexion = New ADODB.Connection
       
    conexion.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                  "Data Source=" & Txtarchivos.Text & _
                  ";Extended Properties=""Excel 8.0;HDR=Yes;"""

  'conexion.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Txtarchivos.Text & ";Extended Properties=""text;HDR=Yes;FMT=Delimited"";"""
         
    ' Nuevo recordset
    Set rsExport = New ADODB.Recordset
       
    With rsExport
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
    End With
         
If Quincena(0) = False And Quincena(1) = False Then
    rsExport.Open "SELECT * FROM [Hoja1$" & Trim(TxtRango.Text) & "]", conexion, , , adCmdText
Else

    'QUINCENAS

    If Quincena(0) = True Then
        rsExport.Open "SELECT * FROM [1Era quincena$" & Trim(TxtRango.Text) & "]", conexion, , , adCmdText
    Else
        rsExport.Open "SELECT * FROM [2da quincena$" & Trim(TxtRango.Text) & "]", conexion, , , adCmdText
    End If
    
    'rsExport.Filter = "placod LIKE '" & mPrefijo & "%'"

End If
       
rsExport.Filter = "placod LIKE '" & mPrefijo & "%'"
       
' Mostramos los datos en el datagrid
    
If rsExport.RecordCount <= 0 Then
    MsgBox "Codigo de trabajadores no corresponden a la compañia", vbCritical, Me.Caption
    GoTo Salir:
End If
    
Set DGrd.DataSource = rsExport
     
    
fc_SumaTotales rsExport, DGrd
    
Salir:

xLibro.Close
If Not xlApp2 Is Nothing Then Set xlApp2 = Nothing
If Not xlApp1 Is Nothing Then Set xlApp1 = Nothing
If Not xLibro Is Nothing Then Set xlBook = Nothing

Exit Sub

ERR:
    MsgBox ERR.Number & "-" & ERR.Description, vbCritical, Me.Caption
    Exit Sub
End Sub

Public Sub AbrirFile(pextension As String)
If Not Cuadro_Dialogo_Abrir(pextension) Then
    Txtarchivos.Text = ""
    Exit Sub
End If
Debug.Print Box.DefaultExt

If UCase(Right(Box.FileName, 3)) <> UCase(Right(pextension, 3)) Then
   MsgBox "La Extensión de archivo no concuerda con el formato elegido", vbCritical, "Archivo Inválido"
   'salir = True
   Exit Sub
End If
Txtarchivos.Text = Box.FileName
Txtarchivos.ToolTipText = Box.FileName

End Sub

Public Function Cuadro_Dialogo_Abrir(pextension As String) As Boolean
 On Error GoTo ErrHandler
   ' Establece los filtros.
   Box.CancelError = True
   Select Case pextension
    Case "*.txt"
        Box.Filter = "All Files (*.*)|*.*|Text Files (*.txt)|*.txt|"
    Case "*.dbf"
        Box.Filter = "All Files (*.*)|*.*|Tablas Files (*.dbf)|*.dbf|"
    Case "*.mdb"
        Box.Filter = "All Files (*.*)|*.*|BD Access (*.mdb)|*.mdb|"
    Case "*.csv"
        Box.Filter = "All Files (*.*)|*.*|Microsoft Excel (*.csv)|*.csv|"
    Case "*.xls"
        Box.Filter = "Excel files (*.xls)|*.xls|All files (*.*)|*.*"
        '"All Files (*.*)|*.*|Microsoft Excel 97/2000 (*.xls)|*.txt)"
   End Select
   ' Especifique el filtro predeterminado.
   Box.FilterIndex = 2
   'Box.FileName = "buenos.csv"
   Box.FileName = ""
   Box.InitDir = "U:\VPINTO\PLLA2008\" 'App.path
   ' Presenta el cuadro de diálogo Abrir.
   Box.ShowOpen
   ' Llamada al procedimiento para abrir archivo.
   Dim pos As String
   'CTA.CTE MN:
   'vNroBco
   
   
   
   Dim swExiste As Variant
   swExiste = InStr(1, UCase(Trim(Box.FileName)), UCase(xFile), vbTextCompare)
   If swExiste = 0 Then
      MsgBox "Archivo Elegido no es el correcto" & Chr(13) & "El Correcto es " & xFile, vbCritical, "Importacion"
      'salir = True
      Txtarchivos = ""
    Else
      Cuadro_Dialogo_Abrir = True
    End If
   Exit Function

ErrHandler:
   Cuadro_Dialogo_Abrir = False
   'El usuario hizo clic en el botón Cancelar.
   Exit Function
End Function

Private Sub Cmbal_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{tab}"
End Sub

Private Sub Cmbdel_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{tab}"
End Sub

Private Sub Cmbfecha_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{tab}"
End Sub

Private Sub Cmbtipo_Change()
Procesa_Cabeza_Boleta
End Sub
Private Sub CmbTipo_Click()
VTipo = fc_CodigoComboBox(Cmbtipo, 2)
Txtsemana.Text = "0"
Select Case VTipo
Case "01" 'NORMAL
    Cmbdel.Visible = False
    Cmbal.Visible = False
    Label5.Visible = False
    Label6.Visible = False
    LblFecPerVac.Visible = False
    FechPerVac.Visible = False
Case "02"
   LblFecPerVac.Visible = True
   Txtsemana.Visible = False
   UpDown1.Visible = False
   Label5.Visible = True
   Label6.Visible = True
   Cmbdel.Visible = True
   Cmbal.Visible = True
   Cmbdel.Enabled = True
   Cmbal.Enabled = True
   FechPerVac.Visible = True
   FechPerVac.Enabled = True
Case "03"
   Cmbdel.Visible = False
   Cmbal.Visible = False
   Label4.Visible = False
   Txtsemana.Visible = False
   UpDown1.Visible = False
   Label5.Visible = False
   Label6.Visible = False
   Cmbdel.Visible = False
   Cmbal.Visible = False
End Select

Cmbtipotrabajador_Click
Procesa_Cabeza_Boleta
End Sub

Private Sub Cmbcia_Click()
    wcia = Trim(Right("00" & Cmbcia.ItemData(Cmbcia.ListIndex), 2))
    Call fc_Descrip_Maestros2("01078", "", Cmbtipo)

    'ACP

   wciamae = Determina_Maestro("01055")
   
   Sql$ = "Select * from maestros_2 where flag1='" & IIf(wtipodoc = True, "02", "04") & "' and status<>'*'"
   Sql$ = Sql$ & wciamae
   
   cn.CursorLocation = adUseClient
   
   Set rs = New ADODB.Recordset
   Set rs = cn.Execute(Sql$, 64)
   
   If rs.RecordCount > 0 Then rs.MoveFirst
   
   Do While Not rs.EOF
      Cmbtipotrabajador.AddItem rs!DESCRIP
      Cmbtipotrabajador.ItemData(Cmbtipotrabajador.NewIndex) = Trim(rs!cod_maestro2)
      rs.MoveNext
   Loop
   
   rs.Close
   If Cmbtipotrabajador.ListCount >= 0 Then Cmbtipotrabajador.ListIndex = 0


If wtipodoc = True Then
   Procesa_Cabeza_Boleta
End If


'If wtipodoc = True Then
'   Call fc_Descrip_Maestros2("01055", "", Cmbtipotrabajador)
'   Procesa_Cabeza_Boleta
'Else
'   wciamae = Determina_Maestro("01055")
'   sql$ = "Select * from maestros_2 where flag1='04' and status<>'*'"
'   sql$ = sql$ & wciamae
'   cn.CursorLocation = adUseClient
'   Set RS = New ADODB.Recordset
'   Set RS = cn.Execute(sql$, 64)
'   If RS.RecordCount > 0 Then RS.MoveFirst
'   Do While Not RS.EOF
'      Cmbtipotrabajador.AddItem RS!Descrip
'      Cmbtipotrabajador.ItemData(Cmbtipotrabajador.NewIndex) = Trim(RS!COD_MAESTRO2)
'      RS.MoveNext
'   Loop
'   RS.Close
'   If Cmbtipotrabajador.ListCount >= 0 Then Cmbtipotrabajador.ListIndex = 0
'End If

End Sub

Private Sub Cmbtipo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{tab}"
End Sub

Private Sub Cmbtipotrabajador_Click()
VTipotrab = fc_CodigoComboBox(Cmbtipotrabajador, 2)
Dim wciamae As String
Dim wBeginMonth As String

VHorasBol = 0
VTipoPago = ""
wciamae = Determina_Maestro("01055")
Sql$ = "Select * from maestros_2 where cod_maestro2='" & VTipotrab & "' and status<>'*'"
Sql$ = Sql$ & wciamae
cn.CursorLocation = adUseClient
Set rs = New ADODB.Recordset
Set rs = cn.Execute(Sql$, 64)
If rs.RecordCount > 0 Then
   rs.MoveFirst
   VHorasBol = Val(rs!flag2)
   VTipoPago = Left(rs!flag1, 2)
End If
If Trim(VTipoPago) = "" Then Exit Sub

If VTipo = "01" Then
    Select Case Left(rs!flag1, 2)
           Case Is <> "02"
                Txtsemana.Text = ""
                Txtsemana.Visible = False
                UpDown1.Visible = False
                Label4.Visible = False
                Label5.Visible = False
                Label6.Visible = False
                Cmbdel.Visible = False
                Cmbal.Visible = False

                
                Sql$ = "select iniciomes from cia where cod_cia='" & wcia & "' and status<>'*'"
                If (fAbrRst(rs, Sql$)) Then
                   If IsNull(rs!iniciomes) Then wBeginMonth = "1" Else wBeginMonth = rs!iniciomes
                End If
                rs.Close
                
                If Trim(wBeginMonth) = "" Then
                    MsgBox "Ingrese el Inicio Del Mes", vbInformation, ""
                Exit Sub
                End If
                
                If Trim(wBeginMonth) <> "1" Then
                   Cmbfecha.Day = Val(wBeginMonth) - 1
                Else
                   Cmbfecha.Day = Val(fMaxDay(Month(Date), Year(Date)))
                End If
           Case Else
                Txtsemana.Visible = True
                UpDown1.Visible = True
                Label4.Visible = True
                asig_fam.Visible = True
'                Label5.Visible = True
'                Label6.Visible = True
'                Cmbdel.Visible = True
'                Cmbal.Visible = True
    End Select
End If
If rs.State = 1 Then rs.Close
Procesa_Cabeza_Boleta
End Sub

Private Sub CmbTipoTrabajador_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{tab}"
End Sub

Private Sub Cmbturno_Click()
    VTurno = Funciones.fc_CodigoComboBox(Cmbturno, 2)
End Sub

Private Sub CmdGrabar_Click()
If RsSeteo.RecordCount = 0 Then
    MsgBox "No existen registros a guardar", vbExclamation, Me.Caption
    Exit Sub
End If
With RsSeteo
    If .RecordCount > 0 Then
        .MoveFirst
        Do While Not .EOF
                
                If Trim(!campo_xls & "") = "" Then
                    MsgBox "Ingrese Nombre de columna del archivo externo", vbExclamation, Me.Caption
                    Me.DgrdSeteo.Col = 0
                    Me.DgrdSeteo.SetFocus
                    Exit Sub
                ElseIf Trim(!campo_sql & "") = "" Then
                    MsgBox "Ingrese Nombre de columna equivalente al Plahistorico", vbExclamation, Me.Caption
                    Me.DgrdSeteo.Col = 1
                    Me.DgrdSeteo.SetFocus
                    Exit Sub
                ElseIf Trim(!Descripcion & "") = "" Then
                    MsgBox "Ingrese Descripcion del Campo Plahistorico", vbExclamation, Me.Caption
                    Me.DgrdSeteo.Col = 2
                    Me.DgrdSeteo.SetFocus
                    Exit Sub
                End If
                
                cn.Execute Sql, 64
            .MoveNext
        Loop
    End If
End With

If MsgBox("Desea Guardar el Seteo", vbDefaultButton2 + vbYesNo + vbQuestion, "") = vbNo Then Exit Sub
Screen.MousePointer = 11
cn.Execute "begin transaction", 64
Sql = "update plaSeteoCampos_plahistorico set status='*' "
cn.Execute Sql, 64
With RsSeteo
    If .RecordCount > 0 Then
        .MoveFirst
        Do While Not .EOF
                Sql = "insert into plaSeteoCampos_plahistorico (camposql,campodbf,fec_crea,status,Descripcion_plahistorico,user_crea,fec_modi,user_modi)"
                Sql = Sql & " values('" & Trim(!campo_sql & "") & "','" & Trim(!campo_xls & "") & "',getdate(),'','" & Trim(!Descripcion) & "','" & Trim(wuser) & "',null,null)"
                cn.Execute Sql, 64
            .MoveNext
        Loop
    End If
End With

cn.Execute "commit transaction", 64
Me.CargaSeteo
Screen.MousePointer = 0
Exit Sub
ErrMsg:
    Screen.MousePointer = 0
    MsgBox ERR.Number & " - " & ERR.Description, vbCritical
End Sub

Private Sub cmdImportar_Click()
    Importar_Movimiento_Planilla
End Sub

Private Function ConfirmarProceso(ByVal pCodCia As String, ByVal pTipoBol As String, ByVal pFechaProceso As String, pNroSemana As String, pNumQuincena As String) As Boolean

Dim MAÑO As Integer
Dim mmes As Integer
Dim strMensaje As String

MAÑO = Val(Mid(pFechaProceso, 7, 4))
mmes = Val(Mid(pFechaProceso, 4, 2))

ConfirmarProceso = True

cn.CursorLocation = adUseClient

Set rs = New ADODB.Recordset

'Planilla de obreros

If pNumQuincena = "" Then
    
    Sql$ = "SELECT DISTINCT fechaproceso FROM plahistorico " & "WHERE cia='" & pCodCia & "' AND proceso='" & pTipoBol & "'" _
            & " AND semana='" & pNroSemana & "' AND YEAR(fechaproceso)=" & MAÑO _
            & " AND MONTH(fechaproceso) = " & mmes _
            & " AND status<>'*' "
   
   strMensaje = "Semana " & Trim(Str(pNroSemana)) & " ya esta cargada" & Chr(13) & "¿Desea anular la planilla y volver a cargar ?"
  
    
End If


'Planilla de empleados (Quincenas)

If pNumQuincena <> "" Then

    Select Case pNumQuincena
        Case "1"
        
            '1 Quincena
            
            Sql$ = "SELECT DISTINCT fechaproceso FROM plaquincena " & "WHERE cia='" & pCodCia & "'" _
                & " AND YEAR(fechaproceso)=" & MAÑO _
                & " AND MONTH(fechaproceso) = " & mmes _
                & " AND status<>'*' "
            
            strMensaje = "Primera quincena ya esta cargada" & Chr(13) & "¿Desea anular la quincena y volver a cargar ?"
        
        Case "2"
            
            '2 Quincena
            
            Sql$ = "SELECT DISTINCT fechaproceso FROM plahistorico " & "WHERE cia='" & pCodCia & "' AND proceso='" & pTipoBol & "'" _
                    & " AND YEAR(fechaproceso)=" & MAÑO _
                    & " AND MONTH(fechaproceso) = " & mmes & " AND SUBSTRING(placod,2,1)='E'" _
                    & " AND status<>'*' "
        
            strMensaje = "Segunda quincena ya esta cargada" & Chr(13) & "¿Desea anular la planilla y volver a cargar ?"
    End Select

End If


If (fAbrRst(rs, Sql$)) Then

ConfirmarProceso = IIf(MsgBox(strMensaje, vbYesNo + vbQuestion, TitMsg) = vbYes, True, False)
    
End If

End Function

Private Sub cmdsalir_Click()
FraSeteo.Visible = False
End Sub

Private Sub CmdSeteo_Click()
'Sql = "select cod_maestro2,descrip from maestros_2 where right(ciamaestro,3)='151' and status<>'*' and rtrim(isnull(codsunat,'')) IN ('21','22') order by DESCRIP"

If RsSeteo.State = 1 Then RsSeteo.Close
RsSeteo.Fields.Append "campo_xls", adChar, 15, adFldIsNullable
RsSeteo.Fields.Append "campo_sql", adChar, 15, adFldIsNullable
RsSeteo.Fields.Append "descripcion", adVarChar, 60, adFldIsNullable
RsSeteo.Open
Set DgrdSeteo.DataSource = RsSeteo


Sql = "SELECT left(COLUMN_NAME,25) as campo,"
Sql = Sql & " rtrim(isnull(left(COLUMN_NAME,25),''))+space(1)+isnull(case when  left(COLUMN_NAME ,1)='a' then (Select descripcion from placonstante where cia='" & wcia & "' and tipomovimiento='03' and codinterno=right(COLUMN_NAME ,2) and status<>'*')"
Sql = Sql & " when  left(COLUMN_NAME ,1)='d' then (Select descripcion from placonstante where cia='" & wcia & "' and tipomovimiento='03' and codinterno=right(COLUMN_NAME ,2) and status<>'*')"
Sql = Sql & " when  left(COLUMN_NAME ,1)='h' then (Select descrip from maestros_2 where status<>'*' and right(ciamaestro,3)= '077' and cod_maestro2=right(COLUMN_NAME ,2) and status<>'*')"
Sql = Sql & " when  left(COLUMN_NAME ,1)='i' then (Select descripcion from placonstante where cia='" & wcia & "' and tipomovimiento='02' and codinterno=right(COLUMN_NAME ,2)  and status<>'*')"
Sql = Sql & " end,'') As concepto"
Sql = Sql & " From information_schema.Columns"
Sql = Sql & " WHERE table_name = 'plahistorico'"
Sql = Sql & " and left(COLUMN_NAME ,1) in ('h','i','d','a')"
Sql = Sql & " ORDER BY ORDINAL_POSITION"

TrueDbgrid_CargarCombo Me.DgrdSeteo, 1, Sql, -1

'Sql = "SELECT left(COLUMN_NAME,25) as campo,"
'Sql = Sql & " isnull(case when  left(COLUMN_NAME ,1)='a' then (Select descripcion from placonstante where cia='" & wcia & "' and tipomovimiento='03' and codinterno=right(COLUMN_NAME ,2) and status<>'*')"
'Sql = Sql & " when  left(COLUMN_NAME,1)='d' then (Select descripcion from placonstante where cia='" & wcia & "' and tipomovimiento='03' and codinterno=right(COLUMN_NAME ,2) and status<>'*')"
'Sql = Sql & " when  left(COLUMN_NAME,1)='h' then (Select descrip from maestros_2 where status<>'*' and right(ciamaestro,3)= '077' and cod_maestro2=right(COLUMN_NAME ,2) and status<>'*')"
'Sql = Sql & " when  left(COLUMN_NAME,1)='i' then (Select descripcion from placonstante where cia='" & wcia & "' and tipomovimiento='02' and codinterno=right(COLUMN_NAME ,2)  and status<>'*')"
'Sql = Sql & " end,'') As concepto"
'Sql = Sql & " From information_schema.Columns"
'Sql = Sql & " WHERE table_name = 'plahistorico'"
'Sql = Sql & " and left(COLUMN_NAME ,1) in ('h','i','d','a')"
'Sql = Sql & " ORDER BY ORDINAL_POSITION"
'
'TrueDbgrid_CargarCombo Me.DgrdSeteo, 2, Sql, -1

CargaSeteo

DgrdSeteo.Refresh
FraSeteo.Visible = True
FraSeteo.ZOrder 0

End Sub

Private Sub Command1_Click()

End Sub

Private Sub DgrdSeteo_AfterColUpdate(ByVal ColIndex As Integer)
Select Case ColIndex
Case 1
    Dim xPos As String
    xPos = InStr(1, DgrdSeteo.Columns(1).Text, " ")
    DgrdSeteo.Columns(2).Text = Right(DgrdSeteo.Columns(1).Text, Len(DgrdSeteo.Columns(1).Text) - xPos)
    DgrdSeteo.Update
End Select

End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
Call rCarCbo(Cmbcia, Carga_Cia, "C", "00")
Call rUbiIndCmbBox(Cmbcia, wcia, "00")

Cmbfecha.Value = Date
Cmbdel.Value = Date
Cmbal.Value = Date

FechPerVac.Value = Date
Sql$ = "Select codturno,descripcion from platurno where cia='" & wcia & "' and status<>'*' order by codturno"
Cmbturno.Clear
VTurno = ""
If (fAbrRst(rs, Sql$)) Then
   If (Not rs.EOF) Then
      Do Until rs.EOF
         Cmbturno.AddItem rs(1)
         Cmbturno.ItemData(Cmbturno.NewIndex) = rs(0)
         rs.MoveNext
       Loop
       Cmbturno.ListIndex = 0
    End If
    If rs.State = 1 Then rs.Close
End If

'PARA VALIDAR CODIGO DE TRABAJADOR

Sql$ = "select prefijo from cia where cod_cia='" & wcia & "' and status<>'*'"
If (fAbrRst(rs, Sql$)) Then
    If IsNull(rs!prefijo) Then mPrefijo = "" Else mPrefijo = rs!prefijo
    If rs.State = 1 Then rs.Close
End If

Quincena(0).Visible = IIf(wtipodoc = True, False, True)
Quincena(1).Visible = IIf(wtipodoc = True, False, True)


End Sub

Private Sub SSCommand1_Click()

    If Quincena(0).Visible Then
        If (Quincena(0) = True Or Quincena(1) = True) Then
            wtipodoc = IIf(Quincena(0) = True, False, True)
        Else
            MsgBox "Debe indicar quincena a procesar ", vbCritical, TitMsg
            Exit Sub
        End If
    
    End If
    
    LstObs.Clear
    Set rsExport = Nothing
    LimpiarRsT rsExport, DGrd
            
    AbrirFile ("*.xls")
    If Trim(Txtarchivos.Text) <> "" Then Importar_Excel

End Sub
Public Sub Procesa_Cabeza_Boleta()
Dim mano As Integer
Dim mmes As Integer
On Error GoTo CORRIGE

If Trim(Txtsemana.Text) <> "" Then
   Sql$ = "select * from plasemanas where cia='" & wcia & "' and ano='" & Format(Cmbfecha.Year, "0000") & "' and semana='" & Format(Trim(Txtsemana.Text), "00") & "'  and status<>'*'"
   cn.CursorLocation = adUseClient
   Set rs = New ADODB.Recordset
   
   Set rs = cn.Execute(Sql$, 64)
   
   If rs.RecordCount > 0 Then
      Cmbdel.Value = Format(rs!fechai, "dd/mm/yyyy")
      Cmbal.Value = Format(rs!fechaf, "dd/mm/yyyy")
      Cmbfecha.Value = Format(rs!fechaf, "dd/mm/yyyy")
   End If
   
   If rs.State = 1 Then rs.Close
End If

If Trim(VTipoPago) = "" Or IsNull(VTipoPago) Then Exit Sub

   mano = Val(Mid(Cmbfecha.Value, 7, 4))
   mmes = Val(Mid(Cmbfecha.Value, 4, 2))

   If wtipodoc = True Then
      Select Case VTipoPago
        Case Is = "02"
             Sql$ = nombre()
             Sql$ = Sql$ & "a.placod,a.totneto,b.pagomoneda as moneda " _
             & " from plahistorico a,planillas b " _
             & " where a.cia='" & wcia & "' and a.proceso='" & VTipo & "' and b.tipotrabajador='" & VTipotrab & "' " _
             & " and a.semana='" & Txtsemana.Text & "' AND YEAR(fechaproceso)=" & Format(Cmbfecha.Year, "0000") & " " _
             & " and month(fechaproceso)=" & Format(Cmbfecha.Month, "00") & " " _
             & " and a.status='T' and a.placod=b.placod and a.cia=b.cia and b.status<>'*'"
        Case Is = "04"
             Sql$ = nombre()
             Sql$ = Sql$ & " a.placod, a.totneto, b.pagomoneda as moneda " _
             & " from plahistorico a,planillas b " _
             & " where a.cia='" & wcia & "' and a.proceso='" & VTipo & "' and b.tipotrabajador='" & VTipotrab & "' " _
             & " and year(a.fechaproceso)=" & mano & " and month(a.fechaproceso)=" & mmes & " " _
             & " and a.status='T' and a.placod=b.placod and a.cia=b.cia and b.status<>'*' " _
             & " order by a.placod"
      End Select
   
   Else
      Sql$ = nombre()
      Sql$ = Sql$ & "a.placod,a.totneto,b.pagomoneda as moneda " _
      & "from plaquincena a,planillas b " _
      & "where a.cia='" & wcia & "' and b.tipotrabajador='" & VTipotrab & "' and year(a.fechaproceso)=" & mano & " and month(a.fechaproceso)=" & mmes & " " _
      & "and a.status<>'*' and a.placod=b.placod and a.cia=b.cia and b.status<>'*' " & _
      " order by a.placod"
   End If

 cn.CursorLocation = adUseClient
 Screen.MousePointer = vbDefault
 Exit Sub
 
CORRIGE:
 MsgBox "Error :" & ERR.Description, vbCritical, Me.Caption
 
End Sub
Sub ProcesaFechas()
If Trim(Txtsemana.Text) <> "" Then
   Sql$ = "select * from plasemanas where cia='" & wcia & "' and ano='" & Format(Cmbfecha.Year, "0000") & "' and semana='" & Format(Trim(Txtsemana.Text), "00") & "'  and status<>'*'"
   cn.CursorLocation = adUseClient
   Set rs = New ADODB.Recordset
   
   Set rs = cn.Execute(Sql$, 64)
   
   If rs.RecordCount > 0 Then
      Cmbdel.Value = Format(rs!fechai, "dd/mm/yyyy")
      Cmbal.Value = Format(rs!fechaf, "dd/mm/yyyy")
      Cmbfecha.Value = Format(rs!fechaf, "dd/mm/yyyy")
      
      'Habilita la semana para calcular ASIGNACION FAMLIAR
      asig_fam.Value = IIf(rs!Calculo_Asigfam = "S", 1, 0)
   End If
   
   If rs.State = 1 Then rs.Close
End If
End Sub

Private Sub TxtRango_Change()
    Dim i As Integer
    TxtRango.Text = UCase(TxtRango.Text)
    i = Len(TxtRango.Text)
    TxtRango.SelStart = i
End Sub

Private Sub Txtsemana_Change()
    ProcesaFechas
End Sub

Public Sub Importar_File()
Dim i, J As Integer
Dim MqueryI As String
Dim MqueryCalD As String
Dim MqueryCalA As String
Dim rscampo As ADODB.Recordset
Dim rstrab As ADODB.Recordset
Dim rsbasico As ADODB.Recordset
Dim rsccosto As ADODB.Recordset
Dim sn_quinta As Boolean

Dim vbasico As Double
Dim VCCosto As String

If Trim(Cmbturno.Text) = "" Then MsgBox "Debe Indicar Turno", vbCritical, TitMsg: Cmbturno.SetFocus: Exit Sub
If Trim(Cmbtipo.Text) = "" Then MsgBox "Debe Indicar Tipo de Boleta", vbCritical, TitMsg: Cmbtipo.SetFocus: Exit Sub
If Trim(Cmbtipotrabajador.Text) = "" Then MsgBox "Debe Indicar Tipo de Trabajador", vbCritical, TitMsg: Cmbtipotrabajador.SetFocus: Exit Sub
If Trim(Txtarchivos.Text) = "" Then MsgBox "Debe Seleccionar archivo a cargar", vbCritical, TitMsg: SSCommand1.SetFocus: Exit Sub

If Cmbtipotrabajador.Text = "OBRERO" And Txtsemana.Text = "" Then MsgBox "Debe indicar semana a procesar ", vbCritical, TitMsg: Txtsemana.SetFocus: Exit Sub

Mstatus = "T"
BolDevengada = False
sn_quinta = True
rsExport.MoveFirst
contador = 0

Do While Not rsExport.EOF
    DoEvents
    P1.Min = 1: P1.Max = rsExport.RecordCount
    
    If EliminaBoleta(wtipodoc, Year(Cmbfecha.Value), Month(Cmbfecha.Value), Day(Cmbfecha.Value), Trim(rsExport!PlaCod), Trim(VTipoPago), Trim(VTipo), Val(Txtsemana.Text)) Then
    
        If VTipo = "01" Then
            If MUESTRA_CUENTACORRIENTE(Trim(rsExport!cia), Trim(rsExport!PlaCod), wtipodoc) Then
            End If
        End If
    
        'basico del trabajador
        '----------------------------------------------------------------------------
        Sql$ = "select importe from plaremunbase where cia='" & Trim(rsExport!cia) & "' " & _
               "and placod='" & Trim(rsExport!PlaCod) & "' and concepto='01' and status<>'*'"
        Set rsbasico = OpenRecordset(Sql$, cn)
        If rsbasico.RecordCount > 0 Then
            vbasico = rsbasico!importe
        End If
        '-----------------------------------------------------------------------------
                
        'datos del trabajador
        '-----------------------------------------------------------------------------
        Sql$ = Funciones.nombre()
        Sql$ = Sql$ & "codauxinterno,a.status,a.tipotrabajador,a.fingreso," & _
             "a.fcese,a.codafp,a.numafp,a.area,a.placod," & _
             "a.codauxinterno,b.descrip,a.tipotasaextra," & _
             "a.cargo,a.altitud,a.vacacion,a.area,a.fnacimiento," & _
             "a.fec_jubila,a.sindicato,a.ESSALUDVIDA,a.quinta " & _
             "from planillas a,maestros_2 b where a.status<>'*' "
             Sql$ = Sql$ & " and a.tipotrabajador=b.cod_maestro2 " _
             & "and cia='" & Trim(rsExport!cia) & "' AND placod='" & Trim(rsExport!PlaCod) & "' "
             Sql$ = Sql$ & " order by nombre"
        Set rstrab = OpenRecordset(Sql$, cn)
        '-----------------------------------------------------------------------------
                                                
        Sql$ = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
        Sql$ = Sql$ & "insert into platemphist(cia,placod," & _
               "codauxinterno,proceso,fechaproceso,fechavacai,fechavacaf,fecperiodovaca,semana," & _
               "fechaingreso,turno,codafp,status,fec_crea," & _
               "tipotrab,obra,numafp,basico,fec_modi," & _
               "user_crea,user_modi) values('" & Trim(rsExport!cia) & _
               "','" & Trim(rsExport!PlaCod) & "','" & _
               Trim(rstrab!codauxinterno) & "','" & Trim(VTipo) & "','" & _
               Format(Cmbfecha.Value, FormatFecha) & "','','', " & _
               "'','" & Trim(Txtsemana.Text) & "','" & Format(rstrab!fIngreso, FormatFecha) & _
               "','" & Format(VTurno, "0") & "','" & _
               rstrab!CodAfp & "','" & Mstatus & "'," & _
               FechaSys & ",'" & VTipotrab & "','', " & _
               "'" & rstrab!NUMAFP & "'," & _
               CCur(vbasico) & "," & FechaSys & _
               ",'" & wuser & "','" & wuser & "')"
        cn.Execute Sql
    
        'Ingreso de datos del archivo a exportar
        '----------------------------------------------------------------------------------
        Sql$ = "select * from maestros_2 where ciamaestro='" & Trim(rsExport!cia) & "044' and flag1='" & Left(Trim(rsExport!PLASEC1), 3) & "'"
        Set rsccosto = OpenRecordset(Sql$, cn)
        If rsccosto.RecordCount > 0 Then
            VCCosto = rsccosto!cod_maestro2
        Else
            LstObs.AddItem "No se identifico el centro de Costo para el trabajador " & rsExport!PlaCod & "-" & rstrab!nombre
            VCCosto = ""
        End If
        'If Trim(rsexport!PLACOD) = "TO132" Then Stop
        Sql$ = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
        Sql$ = Sql$ & "Update platemphist set " & _
        " ccosto1 = '" & Trim(VCCosto) & "', porc1 =100 , " & _
        " h14=" & IIf(IsNull(rsExport!diaslab) Or rsExport!diaslab = " ", 0, rsExport!diaslab) & " ,h01=" & IIf(IsNull(rsExport!plah01) Or rsExport!plah01 = " ", 0, rsExport!plah01) & ",h02=0,h03=" & IIf(IsNull(rsExport!plah03) Or rsExport!plah03 = " ", 0, rsExport!plah03) & ",h04=0,h06=0,h05=0,h07=0,h08=0,h09=0,h10=" & IIf(IsNull(rsExport!plah04) Or rsExport!plah04 = " ", 0, rsExport!plah04) & ", h17=" & IIf(IsNull(rsExport!plah10) Or rsExport!plah10 = " ", 0, rsExport!plah10) & ", h11=" & IIf(IsNull(rsExport!plah06) Or rsExport!plah06 = " ", 0, rsExport!plah06) & ",h12=0,h13=0,h15=0,h18=" & IIf(IsNull(rsExport!plah07) Or rsExport!plah07 = " ", 0, rsExport!plah07) & ", " & _
        " h19 = " & IIf(IsNull(rsExport!plah11) Or rsExport!plah11 = " ", 0, rsExport!plah11) & ", h24=" & IIf(IsNull(rsExport!plah08) Or rsExport!plah08 = " ", 0, rsExport!plah08) & ", h23=" & IIf(IsNull(rsExport!plah09) Or rsExport!plah09 = " ", 0, rsExport!plah09) & ", h25=" & IIf(IsNull(rsExport!diassub) Or rsExport!diassub = " ", 0, rsExport!diassub) & ", h26=" & IIf(IsNull(rsExport!DIASNL) Or rsExport!DIASNL = " ", 0, rsExport!DIASNL) & ", " & _
        " i03=" & IIf(IsNull(rsExport!plai07) Or rsExport!plai07 = " ", 0, rsExport!plai07) & ", i13=" & IIf(IsNull(rsExport!PLAI11) Or rsExport!PLAI11 = " ", 0, rsExport!PLAI11) & ",i16=0,i17=0,i18=0,i20=0,i26=0,i30=0,i31=0, i37=" & IIf(IsNull(rsExport!PLAI22) Or rsExport!PLAI22 = " ", 0, rsExport!PLAI22) & ",i39=0,i40=0,i41=0, " & _
        " d05=0 ,d06=0 ,d07=" & IIf(IsNull(rsExport!plad08) Or rsExport!plad08 = " ", 0, rsExport!plad08) & " ,d09=0 ,d12=0 ,d13=0 ,d15=0 " & _
        " where cia='" & Trim(rsExport!cia) & "' and placod='" & _
        Trim(rsExport!PlaCod) & "' and codauxinterno='" & _
        Trim(rstrab!codauxinterno) & "' and proceso='" & _
        Trim(VTipo) & "' and fechaproceso='" & _
        Format(Cmbfecha.Value, FormatFecha) & _
        "' and semana='" & Trim(Txtsemana.Text) & "' and status='" & _
        Mstatus & "'"
        cn.Execute Sql
        
        'Calculo de Ingresos
        '----------------------------------------------------------------------------------
        MqueryI = CalculoIngresos(Trim(rsExport!cia), Trim(rsExport!PlaCod), Mstatus, Year(Cmbfecha.Value), Month(Cmbfecha.Value), VTipotrab, rstrab!Cargo, rstrab!tipotasaextra, Trim(Txtsemana.Text), VTipo, wtipodoc)
        If MqueryI <> "" Then
           MqueryI = Mid(MqueryI, 1, Len(Trim(MqueryI)) - 1)
           Sql$ = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
           Sql$ = Sql$ & " Update platemphist set " & MqueryI
           Sql$ = Sql$ & " where cia='" & Trim(rsExport!cia) & "' and placod='" & Trim(rsExport!PlaCod) & "' and codauxinterno='" & Trim(rstrab!codauxinterno) & "' and proceso='" & VTipo & "' "
           Sql$ = Sql$ & "and fechaproceso='" & Format(Cmbfecha.Value, FormatFecha) & "' and semana='" & Trim(Txtsemana.Text) & "' and status='" & Mstatus & "'"
           cn.Execute Sql
        End If
        '------------------------------------------------------------------------------------
                
        'Calculo de Deducciones
        '-----------------------------------------------------------------------------------------
        MqueryCalD = CalculoDeducciones(Trim(rsExport!cia), Trim(rsExport!PlaCod), Mstatus, Trim(rstrab!CodAfp), rstrab!tipotasaextra, rstrab!Cargo, rstrab!essaludvida, rstrab!sindicato, Trim(rstrab!fec_jubila & ""), wtipodoc, VTipo, Year(Cmbfecha.Value), Month(Cmbfecha.Value), VHorasBol, Trim(Txtsemana.Text), sn_quinta)
        If MqueryCalD <> "" Then
           MqueryCalD = Mid(MqueryCalD, 1, Len(Trim(MqueryCalD)) - 1)
           Sql$ = " SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
           Sql$ = Sql$ & "Update platemphist set " & MqueryCalD
           Sql$ = Sql$ & " where cia='" & Trim(rsExport!cia) & "' and placod='" & Trim(rsExport!PlaCod) & "' and codauxinterno='" & rstrab!codauxinterno & "' and proceso='" & VTipo & "' "
           Sql$ = Sql$ & "and fechaproceso='" & Format(Cmbfecha.Value, FormatFecha) & "' and semana='" & Trim(Txtsemana.Text) & "' and status='" & Mstatus & "'"
           cn.Execute Sql
           
           Sql$ = " SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
           Sql$ = Sql$ & "Update platemphist set d11=d111+d112+d113+d114+d115"
           Sql$ = Sql$ & " where cia='" & Trim(rsExport!cia) & "' and placod='" & Trim(rsExport!PlaCod) & "' and codauxinterno='" & rstrab!codauxinterno & "' and proceso='" & VTipo & "' "
           Sql$ = Sql$ & "and fechaproceso='" & Format(Cmbfecha.Value, FormatFecha) & "' and semana='" & Trim(Txtsemana.Text) & "' and status='" & Mstatus & "'"
           cn.Execute Sql
        End If
        '-------------------------------------------------------------------------------------------
                
        'Calculo de Aportaciones
        '-------------------------------------------------------------------------------------------
        MqueryCalA = CalculoAportaciones(Trim(rsExport!cia), Trim(rsExport!PlaCod), rstrab!Area, VTipo, Year(Cmbfecha.Value), Month(Cmbfecha.Value), Mstatus)
        If MqueryCalA <> "" Then
           MqueryCalA = Mid(MqueryCalA, 1, Len(Trim(MqueryCalA)) - 1)
           Sql$ = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
           Sql$ = Sql$ & "Update platemphist set " & MqueryCalA
           Sql$ = Sql$ & " where cia='" & Trim(rsExport!cia) & "' and placod='" & Trim(rsExport!PlaCod) & "' and codauxinterno='" & rstrab!codauxinterno & "' and proceso='" & VTipo & "' "
           Sql$ = Sql$ & "and fechaproceso='" & Format(Cmbfecha.Value, FormatFecha) & "' and semana='" & Trim(Txtsemana.Text) & "' and status='" & Mstatus & "'"
           cn.Execute Sql
        End If
        '-------------------------------------------------------------------------------------------
        
        
        Dim mi As String, md As String, ma As String
'
'        mi = "i01+i02+i03+i04+i05+i06+i07+i08+i09+i10+i11+i12+i13+i14+i15+i16+i17+i18+i19+i20+i21+i22+i23+i24+i25+i26+i27+i28+i29+i30+i31+i32+i33+i34+i35+i36+i37+i38+i39+i40+i41+i42+i43+i44+i45+i46+i47+i48+i49+i50"
'        md = "d01+d02+d03+d04+d05+d06+d07+d08+d09+d10+d11+d12+d13+d14+d15+d16+d17+d18+d19+d20"
'        ma = "a01+a02+a03+a04+a05+a06+a07+a08+a09+a10+a11+a12+a13+a14+a15+a16+a17+a18+a19+a20"
'        Sql$ = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
'        Sql$ = Sql$ & "update platemphist set h14=round(h01+h02+h03/8,0),totaling=" & mi & "," _
'             & "totalded=" & md & "," _
'             & "totalapo=" & ma & "," _
'             & "totneto=(" & mi & ")-" & "(" & md & ")"
'           Sql$ = Sql$ & " where cia='" & Trim(rsexport!cia) & "' and " & _
'           " placod='" & Trim(rsexport!PLACOD) & "' and " & _
'           "codauxinterno='" & Trim(rstrab!codauxinterno) & "' and proceso='" & VTipo & "' "
'           Sql$ = Sql$ & "and fechaproceso='" & Format(Cmbfecha.Value, FormatFecha) & "' and semana='" & Trim(Txtsemana.Text) & "' and status='" & Mstatus & "'"
'        cn.Execute Sql
        
'Se Modifico para que no calcule el campo h14
        mi = "i01+i02+i03+i04+i05+i06+i07+i08+i09+i10+i11+i12+i13+i14+i15+i16+i17+i18+i19+i20+i21+i22+i23+i24+i25+i26+i27+i28+i29+i30+i31+i32+i33+i34+i35+i36+i37+i38+i39+i40+i41+i42+i43+i44+i45+i46+i47+i48+i49+i50"
        md = "d01+d02+d03+d04+d05+d06+d07+d08+d09+d10+d11+d12+d13+d14+d15+d16+d17+d18+d19+d20"
        ma = "a01+a02+a03+a04+a05+a06+a07+a08+a09+a10+a11+a12+a13+a14+a15+a16+a17+a18+a19+a20"
        Sql$ = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
        Sql$ = Sql$ & "update platemphist set totaling=" & mi & "," _
             & "totalded=" & md & "," _
             & "totalapo=" & ma & "," _
             & "totneto=(" & mi & ")-" & "(" & md & ")"
           Sql$ = Sql$ & " where cia='" & Trim(rsExport!cia) & "' and " & _
           " placod='" & Trim(rsExport!PlaCod) & "' and " & _
           "codauxinterno='" & Trim(rstrab!codauxinterno) & "' and proceso='" & VTipo & "' "
           Sql$ = Sql$ & "and fechaproceso='" & Format(Cmbfecha.Value, FormatFecha) & "' and semana='" & Trim(Txtsemana.Text) & "' and status='" & Mstatus & "'"
        cn.Execute Sql
        
        
        
'        Sql$ = "UPDATE platemphist set h14=round((COALESCE(h01,0)+COALESCE(h02,0)+COALESCE(h03,0))/8,0) where cia='" & Trim(rsexport!cia) & "' and "
'        Sql$ = Sql$ & " placod='" & Trim(rsexport!PLACOD) & "' and "
'        Sql$ = Sql$ & " codauxinterno='" & Trim(rstrab!codauxinterno) & "' and proceso='" & VTipo & "' "
'        Sql$ = Sql$ & " and fechaproceso='" & Format(Cmbfecha.Value, FormatFecha) & "' and semana='" & Trim(Txtsemana.Text) & "' and status='" & Mstatus & "'"
'        cn.Execute Sql
        
'        Sql$ = "UPDATE platemphist set h14=round((COALESCE(h01,0)+COALESCE(h02,0)+COALESCE(h03,0))/8,0) where cia='" & Trim(rsexport!cia) & "' and "
'        Sql$ = Sql$ & " placod='" & Trim(rsexport!PLACOD) & "' and "
'        Sql$ = Sql$ & " codauxinterno='" & Trim(rstrab!codauxinterno) & "' and proceso='" & VTipo & "' "
'        Sql$ = Sql$ & " and fechaproceso='" & Format(Cmbfecha.Value, FormatFecha) & "' and semana='" & Trim(Txtsemana.Text) & "' and status='" & Mstatus & "'"
'        cn.Execute Sql
        
        
        
        If wtipodoc = True Then
           Sql$ = "insert into plahistorico select * from platemphist"
        Else
           Sql$ = "insert into plaquincena select * from platemphist"
        End If
        Sql$ = Sql$ & " where cia='" & Trim(rsExport!cia) & "' and placod='" & Trim(rsExport!PlaCod) & "' and codauxinterno='" & Trim(rstrab!codauxinterno) & "' and proceso='" & VTipo & "' "
        Sql$ = Sql$ & "and fechaproceso='" & Format(Cmbfecha.Value, FormatFecha) & "' and semana='" & Trim(Txtsemana.Text) & "' and status='" & Mstatus & "'"
        cn.Execute Sql
        
        Sql$ = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
        Sql$ = Sql$ & "select cia,placod,proceso,fechaproceso,semana,tipotrab,d07 from platemphist"
        Sql$ = Sql$ & " where cia='" & Trim(rsExport!cia) & "' and placod='" & Trim(rsExport!PlaCod) & "' and codauxinterno='" & Trim(rstrab!codauxinterno) & "' and proceso='" & VTipo & "' "
        Sql$ = Sql$ & "and fechaproceso='" & Format(Cmbfecha.Value, Coneccion.FormatFecha) & "' and semana='" & Trim(Txtsemana.Text) & "' and status='" & Mstatus & "'"
        cn.Execute Sql
        
        If (fAbrRst(rs, Sql$)) Then
           If Descarga_ctaCte(Trim(rsExport!cia), Trim(rsExport!PlaCod), wtipodoc, Trim(rstrab!codauxinterno), VTipo, Format(Cmbfecha.Value, FormatFecha), Trim(Txtsemana.Text), VTipotrab, IIf(IsNull(rsExport!plad08) Or rsExport!plad08 = "", 0, rsExport!plad08)) Then
           End If
        End If
        
        Sql$ = "delete from platemphist where cia='" & Trim(rsExport!cia) & "' and placod='" & Trim(rsExport!PlaCod) & "'"
        Sql$ = Sql$ & " and cia='" & Trim(rsExport!cia) & "' and placod='" & Trim(rsExport!PlaCod) & "' and codauxinterno='" & Trim(rstrab!codauxinterno) & "' and proceso='" & VTipo & "' "
        Sql$ = Sql$ & "and fechaproceso='" & Format(Cmbfecha.Value, FormatFecha) & "' and semana='" & Trim(Txtsemana.Text) & "' and status='" & Mstatus & "'"
        cn.Execute Sql
                
    contador = contador + 1
    P1.Value = contador
    rsExport.MoveNext
    End If
Loop
MsgBox "Se Realizo la Carga Satisfactoriamente", vbInformation, "Mensaje"
Exit Sub

End Sub
Public Sub DepurarRegistros()
Dim intloop As Integer

intloop = 0

With rsExport

If .RecordCount > 0 Then
    For intloop = 0 To .Fields.count - 1

        If UCase(.Fields(intloop).Name) = UCase("placod") Then
            Exit For
        End If

    Next

    Do While Not .EOF

        If Left(Trim(.Fields(intloop).Value), 1) <> mPrefijo Then
            .Delete
        End If
        
        .MoveNext
    Loop
End If
End With
End Sub

Public Sub Importar_Movimiento_Planilla()
'If rshoras.State <> 1 Then
'    MsgBox "El archivo no se cargó correctamente", vbExclamation, Me.Caption
'    Exit Sub
'End If

Dim MsgError As String
Dim xTipoTrab As String
xTipoTrab = fc_CodigoComboBox(Cmbtipotrabajador, 2)
Dim xTipoBol As String
xTipoBol = fc_CodigoComboBox(Cmbtipo, 2)
Dim b_dominicial As Boolean
Dim B_AsifFam As Boolean
Dim xTurno As String
xTurno = fc_CodigoComboBox(Me.Cmbturno, 2)

If Trim(Cmbcia.Text) = "" Then
    MsgBox "Elija Compañia", vbExclamation, Me.Caption
    Exit Sub
ElseIf Trim(Cmbtipo.Text) = "" Then
    MsgBox "Debe Indicar Tipo de Boleta", vbCritical, Me.Caption
    Cmbtipo.SetFocus:
    Exit Sub
ElseIf Trim(Cmbtipotrabajador.Text) = "" Then
    MsgBox "Debe Indicar Tipo de Trabajador", vbCritical, Me.Caption
    Cmbtipotrabajador.SetFocus
    Exit Sub
ElseIf Trim(Cmbturno.Text) = "" Then
    MsgBox "Debe Indicar Turno", vbCritical, Me.Caption
    Cmbturno.SetFocus
    Exit Sub
ElseIf xTipoTrab = "02" And Trim(Txtsemana.Text) = "" Then 'obrero
    MsgBox "Debe indicar Nro de Semana a procesar ", vbCritical, Me.Caption: Txtsemana.SetFocus: Exit Sub
    
ElseIf xTipoTrab = "02" And Not IsNumeric(Txtsemana.Text) Then   'obrero
    MsgBox "Debe indicar Nro de Semana a procesar valor numérico ", vbCritical, Me.Caption: Txtsemana.SetFocus: Exit Sub
ElseIf xTipoTrab = "02" And Val(Txtsemana.Text) > 55 Then 'obrero
    MsgBox "Nro de Semana incorrecta", vbCritical, Me.Caption: Txtsemana.SetFocus: Exit Sub
ElseIf Trim(Txtarchivos.Text) = "" Then
    MsgBox "Debe Seleccionar archivo a cargar", vbCritical, Me.Caption:
    SSCommand1.SetFocus:
    Exit Sub
End If
Dim xNroSemana As String
If Trim(Txtsemana.Text) = "" Then
    xNroSemana = 0
Else
    xNroSemana = CInt(Txtsemana.Text)
End If

'Confirmar proceso de importación

Dim strNumQuincena As String

strNumQuincena = IIf(Quincena(0) = False And Quincena(1) = False, "", IIf(Quincena(0) = True, "1", "2"))

If Not ConfirmarProceso(wcia, xTipoBol, CStr(Cmbfecha.Value), xNroSemana, strNumQuincena) Then
    Exit Sub
End If

Dim ArrDsctoCTACTE() As Variant
Dim xFecIniVac As String
Dim xFecFinVac As String
Dim xFecPeriodoVac As String

If Not Me.Cmbdel.Visible Then
    xFecIniVac = ""
Else
    xFecIniVac = Me.Cmbdel.Value
End If

If Not Me.Cmbal.Visible Then
    xFecFinVac = ""
Else
    xFecFinVac = Me.Cmbal.Value
End If

If Not FechPerVac.Visible Then
    xFecPeriodoVac = ""
Else
    xFecPeriodoVac = FechPerVac.Value
End If

Dim B_ASIGFAM As Boolean

B_ASIGFAM = False

If asig_fam.Value = 0 Then
    B_ASIGFAM = False
Else
    B_ASIGFAM = True
End If

Screen.MousePointer = 11
                                
Dim intloop As Integer
With rsExport
    If .RecordCount > 0 Then
    
        '//*** Verifica si la estuctura es la correcta ***///
        Dim sw As Integer
        sw = 0
        intloop = 0
        Dim xValue As Variant
        xValue = ""
        For intloop = 0 To .Fields.count - 1
            'Debug.Print .Fields(intloop).Name
            'If UCase(.Fields(intloop).Name) = UCase("cia") Then sw = sw + 1
            If UCase(.Fields(intloop).Name) = UCase("placod") Then sw = sw + 1
            If UCase(.Fields(intloop).Name) = UCase("plah01") Then sw = sw + 1
        Next
        If sw < 2 Then
            MsgBox "El archivo a importar no contiene la estructura necesaria a importar", vbCritical, Me.Caption
            Screen.MousePointer = 0
            Exit Sub
        End If
        
        
        CargaSeteo_Equivalencia_Campos_Plahistorico
        
        Dim CampoPlaHistorico As String
        CampoPlaHistorico = ""
        Barra.Visible = True
        Barra.Max = .RecordCount
        Barra.Min = 1
        Dim i As Integer
        i = 1
        cn.Execute "begin transaction", 64
        .MoveFirst
        
        Dim HAcum As Double
        
        Dim Hdom As Double
        Dim HDesMed  As Double
        
        Dim intFactorCarga As Integer
        intFactorCarga = IIf(wtipodoc = True, 1, IIf(Quincena(0) = True, 2, 1))
        
        Do While Not .EOF
                If .Fields("PlaCod").Value = "RO041" Then Stop
                DoEvents
                Barra.Value = i
                
                i = i + 1
                
                'If Trim(!cia) <> wcia Then GoTo SIGUE:
                           
                 HAcum = 0 'HORAS ORDINARIAS
                 Hdom = 0  'HORAS DOMINICAL
                 HDesMed = 0   'HORAS DESCANSO
                 
                 b_dominicial = False
                 Crea_Rs
                        
                        
                        
'                        If Trim(rstrab!Area) <> "" Then
'                            rsCosto.AddNew
'                            rsCosto!Item = "1"
'                            rsCosto!codigo = Trim(rstrab!Area & "")
'                            rsCosto!Monto = "100"
'                        Else
'                            LstObs.AddItem "No se identifico el centro de Costo para el trabajador " & !PlaCod
'                        End If
                        
                        'Dim Rc As ADODB.Recordset
                        'Dim xCenCosto As String
                        
                        'Sql$ = "select * from maestros_2 where ciamaestro='01044' and flag1='" & Left(Trim(rsExport!PLASEC1), 3) & "'"
                        
'                        wciamae = Determina_Maestro("01044")
'                        Sql = "Select * from maestros_2 where flag1='" & Left(Trim(!PLASEC1), 3) & "' and status<>'*'"
'                        Sql = Sql & wciamae
'                        If fAbrRst(Rc, Sql) Then
'                            rsCosto.AddNew
'                            rsCosto!Item = "1"
'                            rsCosto!codigo = Trim(Rc!COD_MAESTRO2 & "")
'                            rsCosto!Monto = "100"
'                        Else
'                            LstObs.AddItem "No se identifico el centro de Costo para el trabajador " & !PlaCod
'                        End If
                                                                                                            
                        For intloop = 0 To .Fields.count - 1
                            
                            If Len(Trim(.Fields(intloop).Name)) = 2 And Left(Trim(.Fields(intloop).Name), 1) = "F" And IsNumeric(Right(Trim(.Fields(intloop).Name), 1)) Then GoTo SIGUE:
                            If Len(Trim(.Fields(intloop).Name)) = 3 And Left(Trim(.Fields(intloop).Name), 1) = "F" And IsNumeric(Right(Trim(.Fields(intloop).Name), 2)) Then GoTo SIGUE:
                                            
                            CampoPlaHistorico = Buscar_Campo_Equivalente(RsCampoPlaHist, Trim(.Fields(intloop).Name))
                            
                            If Trim(CampoPlaHistorico) = "" Then
                               MsgError = "No existe la equivalencia para el campo " & UCase(Trim(.Fields(intloop).Name)) & " en la tabla Plahistorico" & Chr(13) & "Se cancelará la carga"
                               GoTo Salir:
                            End If
                        
                             If Left(CampoPlaHistorico, 1) = "H" Then
                             
                                    If Not rshoras.EOF Then rshoras.MoveFirst
                                    rshoras.FIND "codigo='" & Right(Trim(CampoPlaHistorico), 2) & "'", 0, 1, 1
                                   
                                    If rshoras.EOF Then
                                        rshoras.AddNew
                                        rshoras!Codigo = Right(Trim(CampoPlaHistorico), 2)
                                        
                                        If Trim(.Fields(intloop).Value & "") <> "" Then
                                            rshoras!monto = Round(.Fields(intloop).Value / intFactorCarga, 2)
                                        Else
                                            rshoras!monto = 0
                                        End If
                                        
                                        'rshoras!Monto = IIf(Trim(.Fields(intloop).Value & "") = "", 0, Round(.Fields(intloop).Value / intFactorCarga, 2))
                                        
                                        If CampoPlaHistorico = "H01" And (IIf(Trim(.Fields(intloop).Value & "") = "", 0, .Fields(intloop).Value)) = 0 Then
                                            b_dominicial = True
                                        End If
                                    
                                        If CampoPlaHistorico = "H01" Or CampoPlaHistorico = "H09" Or CampoPlaHistorico = "H12" Or CampoPlaHistorico = "H23" Or CampoPlaHistorico = "H03" Then
                                           
                                            If Trim(.Fields(intloop).Value & "") <> "" Then
                                                HAcum = HAcum + IIf((IIf(Trim(.Fields(intloop).Value & "") = "", 0, .Fields(intloop).Value)) = 56, 0, (IIf(Trim(.Fields(intloop).Value & "") = "", 0, Round(.Fields(intloop).Value / intFactorCarga, 2))))
                                            End If
                                           
                                           Hdom = (8 / 48) * HAcum
                                           
                                           'DESCANSO MEDICO
                                           
                                           If CampoPlaHistorico = "H23" Then
                                               HDesMed = HDesMed + IIf(Trim(.Fields(intloop).Value & "") = "", 0, .Fields(intloop).Value)
                                              'HDesMed = HDesMed + IIf((IIf(Trim(.Fields(intloop).Value & "") = "", 0, .Fields(intloop).Value)) = 56, 0, (IIf(Trim(.Fields(intloop).Value & "") = "", 0, .Fields(intloop).Value)))
                                           End If
                                        
                                        
                                        End If
                                        
                                    Else
                                        If Trim(.Fields(intloop).Value & "") <> "" Then
                                            rshoras!monto = rshoras!monto + IIf(Trim(.Fields(intloop).Value & "") = "", 0, Round(.Fields(intloop).Value / intFactorCarga, 2))
                                        End If
                                    End If
                                                                  
                             End If

                             If Left(CampoPlaHistorico, 1) = "I" Then
                                    If Not rspagadic.EOF Then rspagadic.MoveFirst
                                    rspagadic.FIND "codigo='" & Right(Trim(CampoPlaHistorico), 2) & "'", 0, 1, 1
                                    If rspagadic.EOF Then
                                        rspagadic.AddNew
                                        rspagadic!Codigo = Right(Trim(CampoPlaHistorico), 2)
                                        If Trim(.Fields(intloop).Value & "") <> "" Then
                                            rspagadic!monto = IIf(Trim(.Fields(intloop).Value & "") = "", 0, Round(.Fields(intloop).Value / intFactorCarga, 2))
                                        Else
                                            rspagadic!monto = 0
                                        End If
                                    Else
                                        If Trim(.Fields(intloop).Value & "") <> "" Then
                                            rspagadic!monto = rspagadic!monto + IIf(Trim(.Fields(intloop).Value & "") = "", 0, Round(.Fields(intloop).Value / intFactorCarga, 2))
                                        End If
                                    End If
                             End If
                           
                            If Left(CampoPlaHistorico, 1) = "D" And Len(CampoPlaHistorico) = 3 Then
                                    If Not rsdesadic.EOF Then rsdesadic.MoveFirst
                                    rsdesadic.FIND "codigo='" & Right(Trim(CampoPlaHistorico), 2) & "'", 0, 1, 1
                                    If rsdesadic.EOF Then
                                        rsdesadic.AddNew
                                        rsdesadic!Codigo = Right(Trim(CampoPlaHistorico), 2)
                                        
                                        If Trim(.Fields(intloop).Value & "") <> "" Then
                                            rsdesadic!monto = IIf(Trim(.Fields(intloop).Value & "") = "", 0, Round(.Fields(intloop).Value / intFactorCarga, 2))
                                        Else
                                            rsdesadic!monto = 0
                                        End If
                                        
                                    Else
                                        If Trim(.Fields(intloop).Value & "") <> "" Then
                                            rsdesadic!monto = rsdesadic!monto + IIf(Trim(.Fields(intloop).Value & "") = "", 0, Round(.Fields(intloop).Value / intFactorCarga, 2))
                                        End If
                                    End If
                             End If
                             
                             
                                                          
SIGUE:
                          Next
                         ' Exit Do
                         
                        'CUANDO SEA CALCULO DEL MES NO EJECUTARSE EN QUINCENA
                        
                        If wtipodoc = True And Mid(!PlaCod, 2, 1) = "O" Then
                             If Hdom <> 0 Then
                             
                                If b_dominicial = False Then
                                   rshoras.AddNew
                                   rshoras!Codigo = "02"
                                   rshoras!monto = Hdom
                                End If
                              
                              End If
                        End If

                        
                        If Quincena(1) Then
                            '2 Quincena
                        
                           If rsdesadic.RecordCount > 0 Then rsdesadic.MoveFirst
                           
                           Dim MAÑO, mmes As Integer
                           
                           MAÑO = Val(Right(CStr(Cmbfecha.Value), 4))
                           mmes = Val(Mid(CStr(Cmbfecha.Value), 4, 2))
                           
                            Sql$ = "select totneto from plaquincena where cia='" & wcia & "' and placod='" & Trim(!PlaCod) & "'  and year(fechaproceso)=" & MAÑO & " and month(fechaproceso)=" & mmes & " and status<>'*' "
                                 
                            If (fAbrRst(rs, Sql$)) Then
                                
                                rsdesadic.FIND "codigo='09'", 0, 1, 1
                                
                                If rsdesadic.EOF Then
                                   rsdesadic.AddNew
                                   rsdesadic!Codigo = "09"
                                   rsdesadic!monto = rs(0)
                                End If
                                rs.Close
                            
                            End If
                           
                        End If
                           
                        Dim oBj As New ClsCalculaBoleta
                                                             
                               
                        If HAcum <> 0 Or HDesMed <> 0 Then
                                
                            If oBj.CalcularBoleta(wcia, False, rshoras, rsCosto, rspagadic, rsdesadic, Trim(!PlaCod) _
                                , xTipoTrab, xTipoBol, CStr(Cmbfecha.Value), xFecIniVac, xFecFinVac, xNroSemana, Format(xTurno, "0"), wtipodoc, "T", _
                                "", xFecPeriodoVac, False, False, ArrDsctoCTACTE(), True, B_ASIGFAM, strNumQuincena) Then
                            
                            Else
                            
                                MsgError = "Error en Calculo de boleta trabajador " & Trim(!PlaCod) & " - Se Cancelará la carga"
                                Set oBj = Nothing
                                Erase ArrDsctoCTACTE
                                GoTo Salir:
                                
                            End If
                                
                        End If
                                
                        Set oBj = Nothing
                        Erase ArrDsctoCTACTE
                                
                
            .MoveNext
        Loop
    End If
    
    Barra.Visible = False
    cn.Execute "commit transaction", 64
    Screen.MousePointer = 0
    MsgBox "Terminó la carga Correctamente " & .RecordCount & " Registros importados", vbInformation, Me.Caption
    
    
End With
Screen.MousePointer = 0

Exit Sub
Salir:
    Screen.MousePointer = 0
    cn.Execute "if @@trancount>0 begin transaction", 64
    MsgBox MsgError, vbCritical, Me.Caption

End Sub

Public Sub Crea_Rs()

    'Set pRsHoras = New adodb.Recordset
    If rshoras.State = 1 Then rshoras.Close
    rshoras.Fields.Append "codigo", adChar, 2, adFldIsNullable
    rshoras.Fields.Append "descripcion", adChar, 100, adFldIsNullable
    rshoras.Fields.Append "monto", adCurrency, 18, adFldIsNullable
    rshoras.Open
    
    'Set pRsCosto = New adodb.Recordset
    If rsCosto.State = 1 Then rsCosto.Close
    rsCosto.Fields.Append "codigo", adChar, 2, adFldIsNullable
    rsCosto.Fields.Append "descripcion", adChar, 45, adFldIsNullable
    rsCosto.Fields.Append "monto", adCurrency, 18, adFldIsNullable
    rsCosto.Fields.Append "item", adChar, 2, adFldIsNullable
    rsCosto.Open
    
    
    'Set pRsPagAdic = New adodb.Recordset
    If rspagadic.State = 1 Then rspagadic.Close
    rspagadic.Fields.Append "codigo", adChar, 2, adFldIsNullable
    rspagadic.Fields.Append "descripcion", adChar, 45, adFldIsNullable
    rspagadic.Fields.Append "monto", adCurrency, 18, adFldIsNullable
    rspagadic.Open
    
    'Set pRsDesAdic = New adodb.Recordset
    If rsdesadic.State = 1 Then rsdesadic.Close
    rsdesadic.Fields.Append "codigo", adChar, 2, adFldIsNullable
    rsdesadic.Fields.Append "descripcion", adChar, 45, adFldIsNullable
    rsdesadic.Fields.Append "monto", adCurrency, 18, adFldIsNullable
    rsdesadic.Open
    
End Sub


Public Sub CargaSeteo_Equivalencia_Campos_Plahistorico()
    If RsCampoPlaHist.State = 1 Then RsCampoPlaHist.Close
    RsCampoPlaHist.Fields.Append "campo_pla", adChar, 10, adFldIsNullable
    RsCampoPlaHist.Fields.Append "campo_xls", adChar, 10, adFldIsNullable
    RsCampoPlaHist.Fields.Append "descrip", adVarChar, 50, adFldIsNullable
    RsCampoPlaHist.Open
    
Dim Sql As String
Sql = "select * from plaSeteoCampos_plahistorico where status<>'*' order by campodbf"
Dim Rq As ADODB.Recordset
If fAbrRst(Rq, Sql) Then
With RsCampoPlaHist
    Do While Not Rq.EOF
            .AddNew
            !campo_pla = UCase(Trim(Rq!camposql & ""))
            !campo_xls = UCase(Trim(Rq!campodbf & ""))
            !DESCRIP = Trim(Rq!Descripcion_plahistorico & "")
            
        Rq.MoveNext
    Loop
End With
Else
End If
Rq.Close
Set Rq = Nothing
End Sub

Public Function Buscar_Campo_Equivalente(ByVal pRs As ADODB.Recordset, ByVal pDato As String) As String
Dim Rc As New ADODB.Recordset
Set Rc = pRs.Clone
With Rc
    If .RecordCount > 0 Then
          .MoveFirst
          .FIND "campo_xls='" & UCase(Trim(pDato)) & "'", 0, 1, 1
          If .EOF Then
            Buscar_Campo_Equivalente = ""
            Debug.Print "falta codigo de sunat=>" & Trim(pDato)
          Else
            Buscar_Campo_Equivalente = Trim(.Fields("campo_pla") & "")
'            If Trim(.Fields("codsunat") & "") = "" Then
'                Debug.Print "falta codigo de sunat=>" & Trim(pId) & ""
'            End If
          End If
    End If
End With
Rc.Close
Set Rc = Nothing
End Function

Private Sub Txtsemana_GotFocus()
ResaltarTexto Txtsemana
End Sub

Private Sub Txtsemana_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{tab}"
End Sub

Private Sub Txtsemana_KeyPress(KeyAscii As Integer)
Txtsemana.Text = Txtsemana.Text + fc_ValNumeros(KeyAscii)
End Sub

Private Sub UpDown1_DownClick()
If Val(Txtsemana.Text) > 1 Then
    Txtsemana.Text = Val(Txtsemana.Text) - 1
End If
End Sub

Private Sub UpDown1_UpClick()
If Val(Txtsemana.Text) < 56 Then
    Txtsemana.Text = Val(Txtsemana.Text) + 1
End If
End Sub
Private Sub fc_SumaTotales(ByRef pControl As ADODB.Recordset, ByRef Tdbgrid As TrueOleDBGrid70.Tdbgrid)

On Error GoTo ErrMsg:
Dim Rc As ADODB.Recordset
Set Rc = pControl.Clone

Rc.Filter = "placod LIKE '" & mPrefijo & "%'"
  
Dim Rt As New ADODB.Recordset
If Rt.State = 1 Then Rt.Close
Dim intloop  As Integer
intloop = 0
With Rc
        
        For intloop = 0 To .Fields.count - 1
            Debug.Print "campo " & .Fields(intloop).Name & "  tipo " & .Fields(intloop).Type
            Select Case .Fields(intloop).Type
               Case adCurrency, adNumeric, 5: xValue = 0#
                    Debug.Print "totalespor  " & .Fields(intloop).Name
                    Rt.Fields.Append .Fields(intloop).Name, adCurrency, , adFldIsNullable
            End Select
        Next
End With
Rt.Open
Rt.AddNew

intloop = 0
With Rc
    If .RecordCount > 0 Then
        .MoveFirst
        Do While Not .EOF
                intloop = 0
                For intloop = 0 To .Fields.count - 1
                   Select Case .Fields(intloop).Type
                    Case adCurrency, adNumeric, 5
                        Rt.Fields(.Fields(intloop).Name) = IIf(IsNull(Rt.Fields(.Fields(intloop).Name)), 0, Rt.Fields(.Fields(intloop).Name)) + IIf(IsNull(.Fields(intloop).Value), 0, .Fields(intloop).Value)
                    End Select
                Next
            .Update
            .MoveNext
        Loop
        .MoveFirst
    End If
End With


With Rt
    If .RecordCount > 0 Then
        .MoveFirst
        Do While Not .EOF
                For intloop = 0 To .Fields.count - 1
                   Select Case .Fields(intloop).Type
                    Case adCurrency, adNumeric, 5
                        Tdbgrid.Columns(.Fields(intloop).Name).FooterText = Format(IIf(IsNull(.Fields(intloop).Value), 0, .Fields(intloop).Value), "###,##0.00")
                        
                    End Select
                Next
                
                
            .MoveNext
        Loop
        .MoveFirst
    End If
End With

Rc.Close
Set Rc = Nothing
Rt.Close
Set Rt = Nothing
Exit Sub
ErrMsg:
    MsgBox ERR.Number & " - " & ERR.Description
End Sub



Public Sub TrueDbgrid_CargarCombo(ByRef Tdbgrd As TrueOleDBGrid70.Tdbgrid, ByVal Col As Integer, ByVal Strsql As String, ByVal default As Integer)
    Dim vItem As New TrueOleDBGrid70.ValueItem
    Dim vItems As TrueOleDBGrid70.ValueItems
    Dim AdoRs As New ADODB.Recordset
    
    Set vItems = Tdbgrd.Columns(Col).ValueItems
    vItems.Clear
    AdoRs.CursorLocation = adUseClient
    AdoRs.Open Strsql, cn, adOpenForwardOnly, adLockReadOnly
    Set AdoRs.ActiveConnection = Nothing
    
    If Not AdoRs.EOF Then
        Do While Not AdoRs.EOF
           vItem.Value = Trim(AdoRs(0))
           vItem.DisplayValue = Trim(AdoRs(1))
           vItems.Add vItem
           AdoRs.MoveNext
        Loop
        AdoRs.Close
    End If
    Set AdoRs = Nothing
    If default <> -1 Then vItems.DefaultItem = default
End Sub

Public Sub CargaSeteo()
Dim Rq As ADODB.Recordset
Sql = "select * from plaSeteoCampos_plahistorico where status<>'*' order by campodbf"
If fAbrRst(Rq, Sql) Then
    Do While Not Rq.EOF
            RsSeteo.AddNew
            RsSeteo!campo_xls = Trim(Rq!campodbf & "")
            RsSeteo!campo_sql = Trim(Rq!camposql & "")
            RsSeteo!Descripcion = Trim(Rq!Descripcion_plahistorico & "")
        Rq.MoveNext
    Loop
End If
Rq.Close
Set Rq = Nothing
End Sub
