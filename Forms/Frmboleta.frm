VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{92F3EE28-0DF4-4CB0-AAA5-E6578B12929B}#1.0#0"; "cbofacil.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form Frmboleta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "» Ingreso de Boletas «"
   ClientHeight    =   9330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   17070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9330
   ScaleWidth      =   17070
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   8895
      Begin VB.TextBox txtasigfam 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4080
         MaxLength       =   8
         TabIndex        =   78
         Top             =   1320
         Width           =   855
      End
      Begin Threed.SSPanel PnlActivos 
         Height          =   615
         Left            =   0
         TabIndex        =   65
         Top             =   1680
         Width           =   5055
         _Version        =   65536
         _ExtentX        =   8916
         _ExtentY        =   1085
         _StockProps     =   15
         BackColor       =   14215660
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.TextBox TxtAct4 
            Height          =   285
            Left            =   4080
            TabIndex        =   72
            Top             =   30
            Width           =   855
         End
         Begin VB.TextBox TxtAct3 
            Height          =   285
            Left            =   3160
            TabIndex        =   70
            Top             =   30
            Width           =   855
         End
         Begin VB.TextBox TxtHor4 
            Height          =   285
            Left            =   4080
            TabIndex        =   73
            Top             =   330
            Width           =   855
         End
         Begin VB.TextBox TxtHor1 
            Height          =   285
            Left            =   1440
            TabIndex        =   67
            Top             =   330
            Width           =   855
         End
         Begin VB.TextBox TxtHor2 
            Height          =   285
            Left            =   2325
            TabIndex        =   69
            Top             =   330
            Width           =   855
         End
         Begin VB.TextBox TxtHor3 
            Height          =   285
            Left            =   3165
            TabIndex        =   71
            Top             =   330
            Width           =   855
         End
         Begin VB.TextBox TxtAct2 
            Height          =   285
            Left            =   2320
            TabIndex        =   68
            Top             =   30
            Width           =   855
         End
         Begin VB.TextBox TxtAct1 
            Height          =   285
            Left            =   1440
            TabIndex        =   66
            Top             =   30
            Width           =   855
         End
         Begin VB.Label Label8 
            Caption         =   "Horas por Equipo"
            Height          =   255
            Left            =   120
            TabIndex        =   77
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label7 
            Caption         =   "Activos Asignados"
            Height          =   255
            Left            =   120
            TabIndex        =   74
            Top             =   45
            Width           =   1455
         End
      End
      Begin Threed.SSCommand BtnVerCalculo 
         Height          =   420
         Left            =   5160
         TabIndex        =   59
         Top             =   1800
         Visible         =   0   'False
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   741
         _StockProps     =   78
         Caption         =   "Ver Calculo"
      End
      Begin MSMask.MaskEdBox Txtvacaf 
         Height          =   255
         Left            =   3120
         TabIndex        =   3
         Top             =   1365
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Txtvacai 
         Height          =   255
         Left            =   1000
         TabIndex        =   2
         Top             =   1360
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox Cmbturno 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   120
         Width           =   1455
      End
      Begin MSMask.MaskEdBox Txtvaca 
         Height          =   255
         Left            =   7800
         TabIndex        =   4
         Top             =   1440
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtFecCese 
         Height          =   255
         Left            =   1005
         TabIndex        =   49
         Top             =   960
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin CboFacil.cbo_facil cbo_TipMotFinPer 
         Height          =   315
         Left            =   4200
         TabIndex        =   54
         Top             =   960
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   556
         NameTab         =   ""
         NameCod         =   ""
         NameDesc        =   ""
         Filtro          =   ""
         OrderBy         =   ""
         SetIndex        =   ""
         NameSistema     =   ""
         Mensaje         =   0   'False
         ToolTip         =   -1  'True
         Enabled         =   -1  'True
         Ninguno         =   -1  'True
         BackColor       =   -2147483643
         BeginProperty ComboFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox Txtcodpla 
         Height          =   285
         Left            =   3000
         MaxLength       =   8
         TabIndex        =   0
         Top             =   120
         Width           =   1095
      End
      Begin Threed.SSCommand CmbDiasAsigFam 
         Height          =   420
         Left            =   6960
         TabIndex        =   80
         Top             =   1800
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   741
         _StockProps     =   78
         Caption         =   "Dias Asignacion Fam"
      End
      Begin VB.Label Label13 
         BackColor       =   &H80000010&
         Height          =   375
         Left            =   5280
         TabIndex        =   79
         Top             =   1320
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H80000010&
         Caption         =   "Fecha Periodo Vac."
         Height          =   195
         Left            =   6360
         TabIndex        =   17
         Top             =   1440
         Width           =   1410
      End
      Begin VB.Label LblFechaIngreso 
         BackColor       =   &H80000010&
         Height          =   255
         Left            =   5040
         TabIndex        =   76
         Top             =   1440
         Width           =   2295
      End
      Begin VB.Label LblPlanta 
         Height          =   255
         Left            =   6960
         TabIndex        =   75
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label LbLSemBolVac 
         Height          =   255
         Left            =   8400
         TabIndex        =   64
         Top             =   0
         Width           =   375
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000010&
         Caption         =   "Motivo Fin de Periodo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2160
         TabIndex        =   53
         Top             =   1020
         Width           =   1935
      End
      Begin VB.Label LblPensiones 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4200
         TabIndex        =   52
         Top             =   600
         Width           =   4575
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000010&
         Caption         =   "Sistema de Pensiones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2160
         TabIndex        =   51
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000010&
         Caption         =   "Basico"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   50
         Top             =   600
         Width           =   480
      End
      Begin VB.Label lblAfpTipoComision 
         BackColor       =   &H80000010&
         Height          =   135
         Left            =   120
         TabIndex        =   48
         Top             =   840
         Width           =   375
      End
      Begin VB.Label LblCese 
         BackColor       =   &H80000010&
         Caption         =   "F. Cese"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   1020
         Width           =   855
      End
      Begin VB.Label LblID 
         AutoSize        =   -1  'True
         BackColor       =   &H80000010&
         Height          =   75
         Left            =   120
         TabIndex        =   46
         Top             =   360
         Width           =   285
      End
      Begin MSForms.Label LblCodTrabSunat 
         Height          =   135
         Left            =   8400
         TabIndex        =   45
         Top             =   1320
         Width           =   375
         BackColor       =   -2147483632
         Size            =   "661;238"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Lblbasico 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   690
         TabIndex        =   28
         Top             =   600
         Width           =   1425
      End
      Begin VB.Label Lblnumafp 
         BackColor       =   &H80000010&
         Height          =   135
         Left            =   5640
         TabIndex        =   27
         Top             =   600
         Visible         =   0   'False
         Width           =   15
      End
      Begin VB.Label Lblcargo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   6600
         TabIndex        =   26
         Top             =   240
         Width           =   45
      End
      Begin VB.Label Lblvaca1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000010&
         Caption         =   "F. Ret. Vac."
         Height          =   195
         Left            =   2160
         TabIndex        =   24
         Top             =   1360
         Width           =   855
      End
      Begin VB.Label Lblcodaux 
         AutoSize        =   -1  'True
         BackColor       =   &H80000010&
         Height          =   75
         Left            =   120
         TabIndex        =   20
         Top             =   480
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H80000010&
         Caption         =   "Turno"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   165
         Width           =   420
      End
      Begin VB.Label Lblvaca2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000010&
         Caption         =   "F. Ini. Vaca."
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   1360
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000010&
         Caption         =   "Trabajador"
         Height          =   195
         Index           =   0
         Left            =   2160
         TabIndex        =   14
         Top             =   165
         Width           =   765
      End
      Begin VB.Label Lblcodafp 
         BackColor       =   &H80000010&
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   720
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Lbltope 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   5640
         TabIndex        =   25
         Top             =   240
         Width           =   45
      End
      Begin VB.Label LblFingreso 
         BackColor       =   &H80000010&
         Height          =   375
         Left            =   5160
         TabIndex        =   23
         Top             =   1200
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Lblnombre 
         BackColor       =   &H80000014&
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
         Left            =   4200
         TabIndex        =   15
         Top             =   150
         Width           =   4575
      End
   End
   Begin VB.Frame frmImportar 
      Height          =   8775
      Left            =   9120
      TabIndex        =   32
      Top             =   120
      Width           =   7935
      Begin VB.ListBox LstObs 
         Appearance      =   0  'Flat
         Height          =   1005
         Left            =   120
         TabIndex        =   44
         Top             =   7560
         Width           =   7605
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
         TabIndex        =   42
         Top             =   7080
         Width           =   2355
      End
      Begin VB.Frame Frame7 
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
         Height          =   5250
         Left            =   120
         TabIndex        =   40
         Top             =   1680
         Width           =   7530
         Begin TrueOleDBGrid70.TDBGrid DGrd 
            Height          =   4905
            Left            =   120
            TabIndex        =   41
            Top             =   240
            Width           =   7305
            _ExtentX        =   12885
            _ExtentY        =   8652
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
      Begin VB.Frame Frame6 
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
         TabIndex        =   33
         Top             =   240
         Width           =   7605
         Begin VB.TextBox Txtarchivos 
            BackColor       =   &H8000000A&
            Height          =   285
            Left            =   1170
            Locked          =   -1  'True
            TabIndex        =   35
            Top             =   705
            Width           =   5370
         End
         Begin VB.TextBox TxtRango 
            Height          =   315
            Left            =   3000
            TabIndex        =   34
            Text            =   "A6:U95"
            Top             =   330
            Visible         =   0   'False
            Width           =   1095
         End
         Begin Threed.SSCommand cmdVerArchivo 
            Height          =   495
            Left            =   6840
            TabIndex        =   36
            Top             =   600
            Width           =   615
            _Version        =   65536
            _ExtentX        =   1085
            _ExtentY        =   873
            _StockProps     =   78
            Picture         =   "Frmboleta.frx":0000
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
            Index           =   1
            Left            =   135
            TabIndex        =   39
            Top             =   720
            Width           =   945
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
            TabIndex        =   38
            Top             =   360
            Visible         =   0   'False
            Width           =   2745
         End
         Begin VB.Label Label10 
            Caption         =   "Ejm: A1:G45"
            Height          =   255
            Left            =   4200
            TabIndex        =   37
            Top             =   360
            Visible         =   0   'False
            Width           =   975
         End
      End
      Begin MSComctlLib.ProgressBar BarraImporta 
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   7080
         Width           =   4965
         _ExtentX        =   8758
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin MSComDlg.CommonDialog Box 
         Left            =   1800
         Top             =   7200
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Descuentos Adicionales"
      Enabled         =   0   'False
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
      Height          =   2640
      Left            =   4680
      TabIndex        =   11
      Top             =   6615
      Width           =   4335
      Begin MSDataGridLib.DataGrid DgrdDesAdic 
         Height          =   2295
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   4048
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
            DataField       =   "descripcion"
            Caption         =   "Descripcion"
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
            DataField       =   "monto"
            Caption         =   "Importe"
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
            DataField       =   "codigo"
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
               Locked          =   -1  'True
               ColumnWidth     =   2204.788
            EndProperty
            BeginProperty Column01 
               Alignment       =   1
               ColumnWidth     =   1200.189
            EndProperty
            BeginProperty Column02 
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Pagos Adicionales"
      Enabled         =   0   'False
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
      Height          =   4335
      Left            =   4680
      TabIndex        =   10
      Top             =   2280
      Width           =   4335
      Begin MSDataGridLib.DataGrid DgrdPagAdic 
         Height          =   3975
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   7011
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
            DataField       =   "descripcion"
            Caption         =   "Descripcion"
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
            DataField       =   "monto"
            Caption         =   "Importe"
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
            DataField       =   "codigo"
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
               Locked          =   -1  'True
               ColumnWidth     =   1995.024
            EndProperty
            BeginProperty Column01 
               Alignment       =   1
               ColumnWidth     =   1395.213
            EndProperty
            BeginProperty Column02 
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
      Begin VB.Label Label2 
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "     Saldo Cta.Cte.                   S/."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   0
         TabIndex        =   63
         Top             =   4680
         Width           =   2655
      End
      Begin VB.Label Lblctacte 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2580
         TabIndex        =   62
         Top             =   4680
         Width           =   1455
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Centro de Costo"
      Enabled         =   0   'False
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
      Height          =   1695
      Left            =   120
      TabIndex        =   9
      Top             =   7600
      Width           =   4335
      Begin VB.ListBox LstCcosto 
         Appearance      =   0  'Flat
         Height          =   810
         Left            =   435
         TabIndex        =   21
         Top             =   675
         Visible         =   0   'False
         Width           =   2655
      End
      Begin MSDataGridLib.DataGrid Dgrdccosto 
         Height          =   1335
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   2355
         _Version        =   393216
         AllowUpdate     =   -1  'True
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
         ColumnCount     =   4
         BeginProperty Column00 
            DataField       =   "descripcion"
            Caption         =   "Centro de Costo"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "monto"
            Caption         =   "Porcentaje"
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
            DataField       =   "codigo"
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
         BeginProperty Column03 
            DataField       =   "item"
            Caption         =   "item"
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
               Button          =   -1  'True
               ColumnWidth     =   2399.811
            EndProperty
            BeginProperty Column01 
               Alignment       =   1
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column02 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column03 
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
      Begin VB.Label Lbltot 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   3150
         TabIndex        =   18
         Top             =   1740
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Constantes de Calculo"
      Enabled         =   0   'False
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
      Height          =   5295
      Left            =   120
      TabIndex        =   8
      Top             =   2280
      Width           =   4575
      Begin MSDataGridLib.DataGrid Dgrdhoras 
         Height          =   4935
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   8705
         _Version        =   393216
         Appearance      =   0
         HeadLines       =   1
         RowHeight       =   18
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
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   3
         BeginProperty Column00 
            DataField       =   "descripcion"
            Caption         =   "Descripcion"
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
         BeginProperty Column01 
            DataField       =   "Monto"
            Caption         =   "Horas"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "codigo"
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
            ScrollBars      =   2
            BeginProperty Column00 
               Locked          =   -1  'True
               ColumnWidth     =   2399.811
            EndProperty
            BeginProperty Column01 
               Alignment       =   1
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column02 
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
   End
   Begin Threed.SSPanel Panelprogress 
      Height          =   735
      Left            =   1680
      TabIndex        =   29
      Top             =   3240
      Visible         =   0   'False
      Width           =   5655
      _Version        =   65536
      _ExtentX        =   9975
      _ExtentY        =   1296
      _StockProps     =   15
      Caption         =   "Generando Vacaciones Devengadas"
      ForeColor       =   8388608
      BackColor       =   13160660
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   1
      Alignment       =   6
      Begin MSComctlLib.ProgressBar Barra 
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   360
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   0
      TabIndex        =   31
      Top             =   3720
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin Threed.SSPanel PnlPreView 
      Height          =   7575
      Left            =   120
      TabIndex        =   55
      Top             =   1680
      Visible         =   0   'False
      Width           =   8895
      _Version        =   65536
      _ExtentX        =   15690
      _ExtentY        =   13361
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
      BevelInner      =   1
      Begin MSComctlLib.ListView LstIngresos 
         Height          =   6135
         Left            =   240
         TabIndex        =   56
         Top             =   240
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   10821
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483624
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Concepto"
            Object.Width           =   4057
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Horas"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Importe"
            Object.Width           =   2117
         EndProperty
      End
      Begin MSComctlLib.ListView LstDeducciones 
         Height          =   3615
         Left            =   5040
         TabIndex        =   57
         Top             =   240
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   6376
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483637
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Concepto"
            Object.Width           =   4057
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Importe"
            Object.Width           =   2117
         EndProperty
      End
      Begin MSComctlLib.ListView LstAportaciones 
         Height          =   2415
         Left            =   5040
         TabIndex        =   58
         Top             =   3960
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   4260
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483626
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Concepto"
            Object.Width           =   4057
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Importe"
            Object.Width           =   2117
         EndProperty
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   420
         Left            =   240
         TabIndex        =   61
         Top             =   6945
         Width           =   8415
         _Version        =   65536
         _ExtentX        =   14843
         _ExtentY        =   741
         _StockProps     =   78
         Caption         =   "CERRAR"
      End
      Begin VB.Label LstNeto 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Height          =   375
         Left            =   240
         TabIndex        =   60
         Top             =   6480
         Width           =   8415
      End
   End
End
Attribute VB_Name = "Frmboleta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim VokDevengue As Boolean
Dim vDevengue As Boolean
Dim Mstatus As String
Dim vPlaCod As String
Dim BolDevengada As Boolean
Dim NumDev As Integer
Dim lItems As Integer
Dim TotalIngresos As Double
Dim WValor03 As Currency
Public w_error As Integer
Dim WValor32 As Currency
Dim WValor42 As Currency
Dim wValor43 As Currency
Dim t_horas_trabajadas As Currency
Public rshoras As New Recordset
Dim rsccosto As New Recordset
Dim rspagadic As New Recordset
Public rsdesadic As New Recordset
Dim VTipobol As String
Dim vTipoTra As String
Dim VTurno As String
Dim vItem As Integer
Dim VSemana As String
Dim VFProceso As String
Dim VPerPago As String
Dim Nueva_area As String
Dim VNewBoleta As Boolean
Dim Vano As Integer
Dim Vmes As Integer
Dim Vdia As Integer
Dim VHoras As Currency
Dim VHorasnormal As Currency
Dim VfDel As String
Dim VfAl As String
Dim mdiasfalta As String
Dim VAltitud As String
Dim VVacacion As String
Dim VArea As String
Dim macui As Currency
Dim macus As Currency
Dim macomi As Currency
Dim macuqtames As Currency
Dim mcancel As Boolean, MSINDICATO As Boolean
Dim VFechaNac As String
Dim VFechaJub As String
Dim VJubilado As String
Dim manos As Integer
Dim mHourDay As Currency
Dim lSubsidio As String

Dim ArrDsctoCTACTE() As Variant
Dim MAXROW As Long
Dim BoletaCargada As Boolean
Dim sn_essaludvida As Byte, sn_sindicato As Byte

Dim sn_quinta As Boolean
Dim sn_Domiciliado As Boolean
Dim Es_Cesado As Boolean
Dim importe As Currency
Dim bImportar As Boolean
'VARIABLES Y OBJ CARGA DESDE EXCEL
Dim rsExport As New ADODB.Recordset
Dim RsCampoPlaHist As New ADODB.Recordset
'Dim RsSeteo As New ADODB.Recordset
Dim mPrefijo As String
Dim RsLiquid As New ADODB.Recordset
Dim lCalculaLiqui As Boolean
Dim rsPlaSeteo As New ADODB.Recordset
Dim vMesesMinimo As Integer
Dim vImporteMinimo As Currency
Dim DiasAsignacion  As Integer

Private Sub BtnVerCalculo_Click()
Ver_Calculo
BtnVerCalculo.Visible = False
End Sub

Private Sub CmbDiasAsigFam_Click()
   DiasAsignacion = 7
   DiasAsignacion = InputBox("Introduzca los dias para Calculo de Asignacion Familiar en Dias ", "Dato 7 Dias es el valor por defecto", 7)
End Sub

Private Sub Cmbturno_Click()
VTurno = Funciones.fc_CodigoComboBox(Cmbturno, 2)
End Sub
Private Sub cmdImportar_Click()
    Importar_Movimiento_Planilla
End Sub
Private Sub Importar_Movimiento_Planilla()
   'Confirmar proceso de importación
   Dim strNumQuincena As String
   Dim NroTrans As Integer
   Dim pImportar As Boolean
   On Error GoTo Salir
   NroTrans = 0
   strNumQuincena = IIf(vTipoTra = "02", "", IIf(wtipodoc = True, "2", "1"))
   If Not ConfirmarProceso(wcia, VTipobol, VSemana, strNumQuincena) Then Exit Sub
   pImportar = True
   'REVISANDO
   Screen.MousePointer = 11
   Dim intloop As Integer
   With rsExport
    If .RecordCount > 0 Then
        '//*** Verifica si la estuctura es la correcta ***///
        
        CargaSeteo_Equivalencia_Campos_Plahistorico
        
        Dim CampoPlaHistorico As String
        CampoPlaHistorico = ""
        BarraImporta.Visible = True
        BarraImporta.Max = .RecordCount
        BarraImporta.Min = 0
        
        Dim I As Integer
        I = 1
        
       'Sql$ = wInicioTrans & " IMPORTAR_PLANILLA"
       'cn.Execute Sql$
        NroTrans = 1
        
        .MoveFirst
        
        Dim HAcum As Double
        
        Dim Hdom As Double
        Dim HDesMed  As Double
        
        
        Dim intFactorCarga As Integer
        'intFactorCarga = IIf(vTipoTra = "02", 1, IIf(wtipodoc = True, 1, 2))
        intFactorCarga = 1
        
        Dim nTotalHoras As Currency
        
        lItems = 0
        Do While Not .EOF
          bImportar = pImportar
          nTotalHoras = 0
          'If .Fields("PlaCod").Value = "RO054" Then Stop
'           If Trim(.Fields("PLAH01").Value & "") <> "" Then nTotalHoras = nTotalHoras + .Fields("PLAH01").Value
'           If Trim(.Fields("PLAH04").Value & "") <> "" Then nTotalHoras = nTotalHoras + .Fields("PLAH04").Value
'           If Trim(.Fields("PLAH07").Value & "") <> "" Then nTotalHoras = nTotalHoras + .Fields("PLAH07").Value
'           If Trim(.Fields("PLAH08").Value & "") <> "" Then nTotalHoras = nTotalHoras + .Fields("PLAH08").Value
'           If Trim(.Fields("PLAH09").Value & "") <> "" Then nTotalHoras = nTotalHoras + .Fields("PLAH09").Value
'           If Trim(.Fields("PLAH10").Value & "") <> "" Then nTotalHoras = nTotalHoras + .Fields("PLAH10").Value
'           If Trim(.Fields("PLAH11").Value & "") <> "" Then nTotalHoras = nTotalHoras + .Fields("PLAH11").Value
'           If Trim(.Fields("HSM").Value & "") <> "" Then nTotalHoras = nTotalHoras + .Fields("HSM").Value
'           If Trim(.Fields("HSE").Value & "") <> "" Then nTotalHoras = nTotalHoras + .Fields("HSE").Value
'           If Trim(.Fields("HPT").Value & "") <> "" Then nTotalHoras = nTotalHoras + .Fields("HPT").Value
'           If Trim(.Fields("MESES").Value & "") <> "" Then nTotalHoras = nTotalHoras + .Fields("MESES").Value
           
           'If Trim(.Fields("PlaCod").Value + "") <> "" And nTotalHoras <> 0 Then
           If Trim(.Fields("CODIGO").Value + "") <> "" Then
                DoEvents
                BarraImporta.Value = I
                I = I + 1
                HAcum = 0 'HORAS ORDINARIAS
                Hdom = 0  'HORAS DOMINICAL
               
      
                Txtcodpla.Text = Trim(!Codigo)
                'If Txtcodpla.Text = "O2137" Then Stop
                Elimina_Boleta
                Limpia_Boleta
                Txtcodpla.Text = Trim(!Codigo & "")
                txtCodPlaLostFocus
                If Rs.State = 1 Then Rs.Close: Set Rs = Nothing
                 
                         For intloop = 0 To .Fields.count - 1
                            CampoPlaHistorico = Buscar_Campo_Equivalente(RsCampoPlaHist, Trim(.Fields(intloop).Name))
                            
'                            If Trim(CampoPlaHistorico) = "" Then
'                               MsgError = "No existe la equivalencia para el campo " & UCase(Trim(.Fields(intloop).Name)) & " en la tabla Plahistorico" & Chr(13) & "Se cancelará la carga"
'                               GoTo Salir:
'                            End If
                        
                             If Left(CampoPlaHistorico, 1) = "H" Then
                                    CampoPlaHistorico = Mid(CampoPlaHistorico, 2, 21)
                                    If Not rshoras.EOF Then rshoras.MoveFirst
                                    rshoras.FIND "codigo='" & Right(Trim(CampoPlaHistorico), 2) & "'", 0, 1, 1
                                   
                                    If Not rshoras.EOF Then
                                        
                                        If Trim(.Fields(intloop).Value & "") <> "" Then
                                            rshoras!Monto = Round(.Fields(intloop).Value / intFactorCarga, 2)
                                        Else
                                            rshoras!Monto = 0
                                        End If
                                        
                                    
                                        If CampoPlaHistorico = "H01" Or CampoPlaHistorico = "H03" Or CampoPlaHistorico = "H05" Then
                                            If Trim(.Fields(intloop).Value & "") <> "" Then
                                                HAcum = HAcum + .Fields(intloop).Value
                                            End If
                                        End If
                                    End If
                                                                  
                             End If
    
                             If Left(CampoPlaHistorico, 1) = "I" Then
                                    CampoPlaHistorico = Mid(CampoPlaHistorico, 2, 21)
                                    If Not rspagadic.EOF Then rspagadic.MoveFirst
                                    rspagadic.FIND "codigo='" & Right(Trim(CampoPlaHistorico), 2) & "'", 0, 1, 1
                                    
                                    If Not rspagadic.EOF Then
                                        If Trim(.Fields(intloop).Value & "") <> "" Then
                                            rspagadic!Monto = IIf(Trim(.Fields(intloop).Value & "") = "", 0, Round(.Fields(intloop).Value / intFactorCarga, 2))
                                        Else
                                            rspagadic!Monto = 0
                                        End If
                                    End If
                             End If
    
                            'If Left(CampoPlaHistorico, 1) = "D" And Len(CampoPlaHistorico) = 3 Then
                            If Left(CampoPlaHistorico, 1) = "D" Then
                                    CampoPlaHistorico = Mid(CampoPlaHistorico, 2, 21)
                                    If Not rsdesadic.EOF Then rsdesadic.MoveFirst
                                    rsdesadic.FIND "codigo='" & Right(Trim(CampoPlaHistorico), 2) & "'", 0, 1, 1
    
                                    If Not rsdesadic.EOF Then
                                       If UCase(CampoPlaHistorico) <> "D09" Then
                                           If Trim(.Fields(intloop).Value & "") <> "" Then
                                               rsdesadic!Monto = IIf(Trim(.Fields(intloop).Value & "") = "", 0, Round(.Fields(intloop).Value / intFactorCarga, 2))
                                           Else
                                               rsdesadic!Monto = 0
                                           End If
                                        End If
                                    End If
                             End If
SIGUE:
                     Next
                     
                     
                    Me.Refresh
                    If strNumQuincena = "1" And wruc = "20100037689" Then 'En Agregados la quincena no se calcula se ingresa directamente
                       If Trim(Txtcodpla.Text & "") <> "" And IsNumeric(Trim(rsExport!Quincena & "")) Then
                          If rsExport!Quincena > 0 Then
                             Sql$ = "usp_Pla_Ingresa_Quincena_Comacsa '" & wcia & "','" & Txtcodpla.Text & "','" & VFProceso & "'," & rsExport!Quincena & ",'" & wuser & "'"
                             cn.Execute Sql$
                             lItems = lItems + 1
                          End If
                       End If
                    Else
                       If Trim(Txtcodpla.Text & "") <> "" Then Grabar_Boleta
                       If w_error = 1 Then MsgBox "No Procede Grabacion", vbInformation, Me.Caption
                       If Rs.State = 1 Then Rs.Close: Set Rs = Nothing
                    End If
                End If
 
            .MoveNext
        Loop
    End If
    BarraImporta.Visible = False

    'cn.CommitTrans
    'Sql$ = wFinTrans & " IMPORTAR_PLANILLA"
    'cn.Execute Sql$
    
    
    Screen.MousePointer = 0
    MsgBox "Terminó la carga Correctamente " & lItems & " Registros importados", vbInformation, Me.Caption
End With
Screen.MousePointer = 0

Exit Sub
Salir:
    If NroTrans = 1 Then
        'Sql$ = wCancelTrans & " IMPORTAR_PLANILLA"
        'cn.Execute Sql$
    End If
    Screen.MousePointer = 0
    MsgBox Err.Description, vbCritical, Me.Caption
    

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
            Buscar_Campo_Equivalente = Trim(.Fields("tipo") & "") & Trim(.Fields("campo_pla") & "")
'            If Trim(.Fields("codsunat") & "") = "" Then
'                Debug.Print "falta codigo de sunat=>" & Trim(pId) & ""
'            End If
          End If
    End If
End With
Rc.Close
Set Rc = Nothing
End Function
Private Sub CargaSeteo_Equivalencia_Campos_Plahistorico()
'IMPLEMENTACION GALLOS
    
    If RsCampoPlaHist.State = 1 Then RsCampoPlaHist.Close
    
    RsCampoPlaHist.Fields.Append "campo_pla", adVarChar, 20, adFldIsNullable
    RsCampoPlaHist.Fields.Append "campo_xls", adVarChar, 20, adFldIsNullable
    RsCampoPlaHist.Fields.Append "descrip", adVarChar, 200, adFldIsNullable
    RsCampoPlaHist.Fields.Append "tipo", adChar, 1, adFldIsNullable
    RsCampoPlaHist.Open
    
Dim Sql As String
Sql = "select * from plaSeteoCampos_plahistorico where status<>'*' order by campodbf"
Dim Rq As ADODB.Recordset
If fAbrRst(Rq, Sql) Then
    With RsCampoPlaHist
        Do While Not Rq.EOF
                .AddNew
                !campo_pla = UCase(Trim(Rq!camposql & ""))
                !campo_xls = UCase(Trim(Rq!campoexcel & ""))
                !DESCRIP = Trim(Rq!Descripcion & "")
                !tipo = Trim(Rq!tipo & "")
            Rq.MoveNext
        Loop
    End With
End If

Rq.Close
Set Rq = Nothing
End Sub

Private Sub cmdVerArchivo_Click()
'IMPLEMENTACION GALLOS

    LstObs.Clear
    Set rsExport = Nothing
    LimpiarRsT rsExport, DGrd
            
    AbrirFile ("*.xls")
    If Trim(Txtarchivos.Text) <> "" Then Importar_Excel

End Sub
Public Function Cuadro_Dialogo_Abrir(pextension As String) As Boolean
'IMPLEMENTACION GALLOS

 'On Error GoTo ErrHandler
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
Public Sub AbrirFile(pextension As String)
'IMPLEMENTACION GALLOS

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

Public Sub Importar_Excel()

'IMPLEMENTACION GALLOS

    'Referencia a la instancia de excel
    Dim xlApp2 As Excel.Application
    Dim xlApp1 As Excel.Application
    Dim xLibro As Excel.Workbook
        
    On Error Resume Next
        
    'Chequeamos si excel esta corriendo
        
    Set xlApp1 = GetObject(, "Excel.Application")
    If xlApp1 Is Nothing Then
        'Si excel no esta corriendo, creamos una nueva instancia.
        Set xlApp1 = CreateObject("Excel.Application")
    End If
        
    'ACP On Error GoTo 0
    
    On Error GoTo Err
        
    'Variable de tipo Aplicación de Excel
    
    Set xlApp2 = xlApp1.Application
    
    'Una variable de tipo Libro de Excel
    
    Dim Col As Integer, Fila As Integer
  
  
    'Usamos el método open para abrir el archivo que está _
     en el directorio del programa llamado archivo.xls
    
    Set xLibro = xlApp2.Workbooks.Open(Txtarchivos.Text)
  
    'Hacemos el Excel Invisible
    
    xlApp2.Visible = False
    
    Dim xTipoTrab As String
    xTipoTrab = vTipoTra
    
'    If xTipoTrab = "02" Then 'obrero verifica titulo de la semana
'        Dim CadSem As String
'        CadSem = xLibro.Sheets(1).Cells(2, 1).Value
'        Dim pos As Integer
'        pos = InStr(1, CadSem, VSemana)
'        If pos = 0 Then
'           MsgBox "El contenido del archivo no pertences a la Semana " & VSemana, vbCritical, Me.Caption
'            GoTo Salir:
'        End If
'    End If
    
  
    'Eliminamos los objetos si ya no los usamos
    
    If Not xlApp2 Is Nothing Then Set xlApp2 = Nothing
    If Not xlApp1 Is Nothing Then Set xlApp1 = Nothing
    If Not xLibro Is Nothing Then Set xlBook = Nothing
    
    Dim conexion As ADODB.Connection, Rs As ADODB.Recordset
  
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
         
         
Set rsPlaSeteo = Nothing
Sql$ = "select camposql, campoexcel from plaSeteoCampos_plahistorico where status !='*' order by tipo "
cn.CursorLocation = adUseClient
Set rsPlaSeteo = New ADODB.Recordset
Set rsPlaSeteo = cn.Execute(Sql$, 64)
         
If VTipobol = "03" Then
   rsExport.Open "SELECT * FROM [GRATI" & Format(Vmes, "00") & "$] Where Codigo<>'' Order by Codigo", conexion, , , adCmdText
Else
   If vTipoTra = "02" Then
      'CARGA SEMANAL
      'rsExport.Open "SELECT * FROM [Hoja1$" & Trim(TxtRango.Text) & "]", conexion, , , adCmdText
      If rsPlaSeteo.RecordCount > 0 Then
      Else
         MsgBox "Debe setear la columnas del archivo cargado con las columnas de la Boleta", vbCritical, Me.Caption
         GoTo Salir:
      End If
      rsExport.Open "SELECT * FROM [SEMANA" & Trim(VSemana) & "$] Where Codigo<>'' Order by Codigo", conexion, , , adCmdText
   Else
      'CARGA QUINCENAL
      If wtipodoc = False Then
         rsExport.Open "SELECT * FROM [1Era quincena$] Where Codigo<>'' Order by Codigo", conexion, , , adCmdText
      Else
         'rsExport.Open "SELECT * FROM [2da quincena$" & Trim(TxtRango.Text) & "]", conexion, , , adCmdText
         rsExport.Open "SELECT * FROM [2da quincena$] Where Codigo<>'' Order by Codigo", conexion, , , adCmdText
      End If
   End If
End If
If mPrefijo <> "" Then
    rsExport.Filter = "placod = '" & mPrefijo & "'"
End If
' Mostramos los datos en el datagrid
    
If rsExport.RecordCount <= 0 Then
    MsgBox "Codigo de trabajadores no corresponden a la compañia", vbCritical, Me.Caption
    GoTo Salir:
End If
    
Set DGrd.DataSource = rsExport
Dim I As Integer
For I = 0 To DGrd.Columns.count - 1
   DGrd.Columns(I).Visible = False
Next

If rsPlaSeteo.RecordCount > 0 Then rsPlaSeteo.MoveFirst
Do While Not rsPlaSeteo.EOF
   For I = 0 To DGrd.Columns.count - 1
      If UCase(Trim(DGrd.Columns(I).Caption & "")) = UCase(Trim(rsPlaSeteo!campoexcel & "")) Then DGrd.Columns(I).Visible = True: Exit For
   Next
   rsPlaSeteo.MoveNext
Loop
    
fc_SumaTotalesImportacion rsExport, DGrd
    
Salir:

xLibro.Close
If Not xlApp2 Is Nothing Then Set xlApp2 = Nothing
If Not xlApp1 Is Nothing Then Set xlApp1 = Nothing
If Not xLibro Is Nothing Then Set xlBook = Nothing
If rsPlaSeteo.State = 1 Then rsPlaSeteo.Close
Set rsPlaSeteo = Nothing
Exit Sub

Err:
    If rsPlaSeteo.State = 1 Then rsPlaSeteo.Close
    Set rsPlaSeteo = Nothing

    MsgBox Err.Number & "-" & Err.Description, vbCritical, Me.Caption
    Exit Sub
End Sub
Private Function ConfirmarProceso(ByVal pCodCia As String, ByVal pTipoBol As String, pNroSemana As String, pNumQuincena As String) As Boolean
'IMPLEMENTACION GALLOS

Dim strMensaje As String

ConfirmarProceso = True

cn.CursorLocation = adUseClient

Set Rs = New ADODB.Recordset

'Planilla de obreros

If pNumQuincena = "" And pTipoBol = "01" Then
    
    Sql$ = "SELECT DISTINCT fechaproceso FROM plahistorico " & "WHERE cia='" & pCodCia & "' AND proceso='" & pTipoBol & "'" _
            & " AND semana='" & pNroSemana & "' AND YEAR(fechaproceso)=" & Vano _
            & " AND MONTH(fechaproceso) = " & Vmes _
            & " AND TipoTrab='" & vTipoTra & "' and status<>'*' "
   
   strMensaje = "Semana " & Trim(pNroSemana) & " ya esta cargada" & Chr(13) & "¿Desea anular la planilla y volver a cargar ?"
End If

If pTipoBol = "03" Then
    
    Sql$ = "SELECT DISTINCT fechaproceso FROM plahistorico " & "WHERE cia='" & pCodCia & "' AND proceso='" & pTipoBol & "'" _
            & " AND YEAR(fechaproceso)=" & Vano _
            & " AND MONTH(fechaproceso) = " & Vmes _
            & " AND TipoTrab='" & vTipoTra & "' and status<>'*' "
   
   strMensaje = "Gratificación ya esta cargada" & Chr(13) & "¿Desea anular la planilla y volver a cargar ?"
End If

'Planilla de empleados (Quincenas)

If pNumQuincena <> "" Then

    Select Case pNumQuincena
        Case "1"
        
            '1 Quincena
            
            Sql$ = "SELECT DISTINCT fechaproceso FROM plaquincena " & "WHERE cia='" & pCodCia & "'" _
                & " AND YEAR(fechaproceso)=" & Vano _
                & " AND MONTH(fechaproceso) = " & Vmes _
                & " AND status<>'*' "
            
            strMensaje = "Primera quincena ya esta cargada" & Chr(13) & "¿Desea anular la quincena y volver a cargar ?"
        
        Case "2"
            
            '2 Quincena
            
            Sql$ = "SELECT DISTINCT fechaproceso FROM plahistorico " & "WHERE cia='" & pCodCia & "' AND proceso='" & pTipoBol & "'" _
                    & " AND YEAR(fechaproceso)=" & Vano _
                    & " AND MONTH(fechaproceso) = " & Vmes & " AND SUBSTRING(placod,2,1)='E'" _
                    & " AND status<>'*' "
        
            strMensaje = "Segunda quincena ya esta cargada" & Chr(13) & "¿Desea anular la planilla y volver a cargar ?"
    End Select

End If


If (fAbrRst(Rs, Sql$)) Then

ConfirmarProceso = IIf(MsgBox(strMensaje, vbYesNo + vbQuestion, TitMsg) = vbYes, True, False)
    
End If

End Function
Private Sub fc_SumaTotalesImportacion(ByRef pControl As ADODB.Recordset, ByRef Tdbgrid As TrueOleDBGrid70.Tdbgrid)
'IMPLEMENTACION GALLOS

On Error GoTo ErrMsg:
Dim Rc As ADODB.Recordset
Set Rc = pControl.Clone

If mPrefijo <> "" Then
    Rc.Filter = "placod = '" & mPrefijo & "'"
End If
Dim Rt As New ADODB.Recordset
If Rt.State = 1 Then Rt.Close
Dim intloop  As Integer
intloop = 0
With Rc
        
        For intloop = 0 To .Fields.count - 1
            Debug.Print "campo " & .Fields(intloop).Name & "  tipo " & .Fields(intloop).Type
            Select Case .Fields(intloop).Type
               Case adCurrency, adNumeric, adDouble, adDecimal: xValue = 0#
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
    MsgBox Err.Number & " - " & Err.Description
End Sub

Private Sub DGrd_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
Cancel = True
End Sub

Private Sub Dgrdccosto_AfterColEdit(ByVal ColIndex As Integer)
Select Case ColIndex
    Case Is = 1
         Dgrdccosto.Columns(1) = Format(Dgrdccosto.Columns(1), "###,###.00")
         Total_porcentaje ("N")
End Select
End Sub

Private Sub Dgrdccosto_AfterDelete()
Total_porcentaje ("S")
End Sub

Private Sub Dgrdccosto_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
If Dgrdccosto.Col = 0 Then
        KeyAscii = 0
        Cancel = True
        Dgrdccosto_ButtonClick (ColIndex)
End If
End Sub

Private Sub Dgrdccosto_ButtonClick(ByVal ColIndex As Integer)
Dim Y As Integer, xtop As Integer, xleft As Integer

Y = Dgrdccosto.Row
xtop = Dgrdccosto.Top + Dgrdccosto.RowTop(Y) + Dgrdccosto.RowHeight
Select Case ColIndex
Case 0:
       xleft = Dgrdccosto.Left + Dgrdccosto.Columns(0).Left
       With LstCcosto
       If Y < 8 Then
         .Top = xtop
       Else
         .Top = Dgrdccosto.Top + Dgrdccosto.RowTop(Y) - .Height
       End If
        .Left = xleft
        .Visible = True
        .SetFocus
        .ZOrder 0
       End With
End Select
End Sub
Private Sub DgrdDesAdic_AfterColEdit(ByVal ColIndex As Integer)
Select Case ColIndex
       Case Is = 1
'            If wGrupoPla <> "01" Then 'No Controlar Cta Cte por el momento en grupo GALLOS
'               If DgrdDesAdic.Columns(2) = "07" And CCur(Val(DgrdDesAdic.Columns(1))) > CCur(Format(Lblctacte.Caption, "########0.00")) Then
'                  MsgBox "El Importe no debe ser Mayor al Saldo", vbInformation, "Cuenta Corriente"
'                  DgrdDesAdic.Columns(1) = "0.00"
'               Else
'                   ProrrateaCtaCte CCur(Val(DgrdDesAdic.Columns(1)))
'               End If
'            End If
End Select
End Sub

Private Sub DgrdDesAdic_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
'**********DESHABILITAR PARA PPM****************************
If DgrdDesAdic.Columns(2) = "09" And VTipobol <> "05" Then KeyAscii = 0
If VTipobol = "10" Then Exit Sub
If DgrdDesAdic.Columns(2) = "05" Then 'Retencion judicial
   Dim lRetJud As Boolean
   lRetJud = False
   If VTipobol = "11" Then
      Sql$ = "select importe,status from pladeducper where cia='" & wcia & "' and placod='" & Trim(Txtcodpla.Text) & "' and concepto='05' and status<>'*' and placod in(select placod from PlaRetJudUti where cia='" & wcia & "' and status<>'*')"
   Else
      Sql$ = "select importe,status from pladeducper where cia='" & wcia & "' and placod='" & Trim(Txtcodpla.Text) & "' and concepto='05' and status<>'*'"
   End If
   If (fAbrRst(Rs, Sql$)) Then
      If Not IsNull(Rs(0)) Then If Rs(0) > 0 Then lRetJud = True
   End If
   If Rs.State = 1 Then Rs.Close
   If lRetJud Then Cancel = True
End If
If DgrdDesAdic.Columns(2) = "13" And sn_quinta = True Then Cancel = True

'***********************************************************


End Sub

Private Sub DgrdDesAdic_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
'**********DESHABILITAR PARA PPM****************************
If DgrdDesAdic.Columns(2) = "09" And VTipobol <> "05" Then Cancel = True
'***********************************************************
End Sub

Private Sub DgrdDesAdic_DblClick()
' Call Frmplacte.Show
' Frmplacte.Txtcodpla = Txtcodpla
' Frmplacte.proc = 1 'DE DONDE VIENE
'
' Call frmdesccta.Show
' frmdesccta.proc = 1
' frmdesccta.txtcod.Text = Me.Txtcodpla
' Call frmdesccta.Txtcod_KeyPress(13)
 
End Sub

Private Sub DgrdDesAdic_KeyPress(KeyAscii As Integer)
'**********DESHABILITAR PARA PPM GIOVANNI*******************
If Format(DgrdDesAdic.Columns(2) & "", "00") = "09" And VTipobol <> "05" Then KeyAscii = 0
'***********************************************************
End Sub

Private Sub DgrdHoras_AfterColEdit(ByVal ColIndex As Integer)
Dim Fila As Integer
Fila = Dgrdhoras.Row
Dim VALOR As Currency

If VTipobol <> "03" Then
   VALOR = Val(Dgrdhoras.Columns(1))
   If (VALOR < 0 Or VALOR > 7) And vTipoTra = "02" And Dgrdhoras.Columns(2) = "14" Then VALOR = 6: Dgrdhoras.Columns(1) = "6.00"
End If

If Dgrdhoras.Columns(2) = "01" Then
    t_horas_trabajadas = Dgrdhoras.Columns(1)
End If


'Dgrdhoras.Columns(1) = Format(Dgrdhoras.Columns(1), "###,###.00")
'If Not rshoras.EOF Then rshoras.MoveNext
'If vTipoTra = "02" And Dgrdhoras.Columns(2) = "14" Then
'    Dim rsTemp As New ADODB.Recordset
'    Set rsTemp = rshoras.Clone
'
'    rsTemp.MoveFirst
'    Do While Not rsTemp.EOF
'        ' add LFSA 11/03/2013
'        ' and VALOR <= 6
'        If rsTemp!codigo = "01" And VALOR <= 6 Then
'            rsTemp!Monto = VALOR * 8
'            rsTemp.Update
'            Exit Do
'        End If
'        rsTemp.MoveNext
'    Loop
'
'
'    rsTemp.MoveFirst
'
'    Do While Not rsTemp.EOF
'        If rsTemp!codigo = "02" And VALOR <= 6 Then
'            rsTemp!Monto = (VALOR / 6) * 8
'            Exit Do
'        End If
'        rsTemp.MoveNext
'    Loop
'
'    contar = 0
'    rsTemp.MoveFirst
'    Set rshoras = rsTemp.Clone
'    Do While Not rshoras.EOF
'        If Fila = conta Then
'            rshoras.MoveNext
'            Exit Do
'        Else
'            rshoras.MoveNext
'        End If
'        conta = conta + 1
'    Loop
'End If

If Not rshoras.EOF Then rshoras.MoveNext

'Dgrdhoras.Refresh
End Sub
'
'Private Sub Dgrdhoras_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
''If vTipoTra = "02" And rshoras!codigo = "01" Then Cancel = True
''If vTipoTra = "02" And rshoras!codigo = "02" Then Cancel = True
'End Sub

Private Sub Dgrdhoras_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
Dim VALOR  As Currency
Dim Fila As Integer
Dim contar As Integer

If VTipobol <> "03" Then
   If vTipoTra <> "02" Or rshoras!Codigo <> 14 Then Exit Sub
   VALOR = Dgrdhoras.Columns(1)
    Fila = Dgrdhoras.Row
   If VALOR < 0 Or VALOR > 7 Then Dgrdhoras.Columns(1) = OldValue: Cancel = 1: Exit Sub
End If
End Sub

Private Sub DgrdPagAdic_AfterColEdit(ByVal ColIndex As Integer)
        Select Case ColIndex
        Case 1
            With DgrdPagAdic
                If .Columns(2) = "46" And CCur(Val(.Columns(1))) > 0 Then
                    If rsdesadic.RecordCount > 0 Then
                        rsdesadic.MoveFirst
                        Do While Not rsdesadic.EOF
                            If rsdesadic("CODIGO") = "12" Then
                                rsdesadic("MONTO") = rsdesadic("MONTO") + CCur(Val(.Columns(1)))
                                rsdesadic.MoveFirst
                                DgrdDesAdic.Refresh
                                Exit Do
                            End If

                            rsdesadic.MoveNext
                        Loop
                    End If
                End If
                'GIRAO CAPTURO INGRESO SUBSIDIO
                
                If .Columns(2) = "03" Then  'MOVILIDAD
                    WValor03 = Format(CDec(.Columns(1)), "0#.00")
                End If
                If .Columns(2) = "32" Then 'ENFERMEDADES PAGADAS
                   'WValor32 = Format(CDec(.Columns(1)), "0#")
                    'add jcms 230921 corrige , para que no redondee
                    WValor32 = Format(CDec(.Columns(1)), "0#.00")
                End If
                If .Columns(2) = "42" Then 'MATERNIDAD
                    WValor42 = Format(CDec(.Columns(1)), "0#")
                End If
                If .Columns(2) = "43" Then  'SUBS ACC TRABAJO
                    'wValor43 = Format(CDec(.Columns(1)), "0#")
                    'add jcms 230921 corrige , para que no redondee
                    wValor43 = CDec(.Columns(1))
                End If
          
                      
            End With
    End Select
End Sub

Private Sub DgrdPagAdic_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
If DgrdPagAdic.Columns(2) <> "18" And VTipobol = "11" Then Cancel = True
End Sub

Private Sub DgrdPagAdic_BeforeUpdate(Cancel As Integer)
If DgrdPagAdic.Columns(2) <> "18" And VTipobol = "11" Then Cancel = True
End Sub

Private Sub Form_Activate()
If wtipodoc = True Then
   'Me.Caption = "Ingreso de Boletas"
   'Frmboleta.BackColor = &H80000001
Else
   Me.Caption = "Adelanto de Quincena"
   'Frmboleta.BackColor = &H808000
End If
Sql$ = "Select * From Pla_Basico_Minimo where cia='" & wcia & "' and tipotrab='" & vTipoTra & "' and status<>'*'"
If (fAbrRst(Rs, Sql$)) Then
    vMesesMinimo = Rs!meses: vImporteMinimo = Rs!importe
End If
If Rs.State = 1 Then Rs.Close

If wImportarXls And VTipobol = "11" And wImportaUti = True Then
   Importar_Utilidades
End If

End Sub

Private Sub Form_Load()
Dim wciamae As String
Me.Top = 0
Me.Left = 0
WValor03 = 0#
WValor32 = 0
WValor42 = 0
wValor43 = 0
'Quitar para calculo de liquidacion lCalculaLiqui=false
lCalculaLiqui = False
If wGrupoPla = "01" Then lCalculaLiqui = True
'IMPLEMENTACION GALLOS

Me.Width = IIf(wImportarXls, 17160, 9135)

TxtRango.Enabled = IIf(wImportarXls, True, False)
cmdVerArchivo.Enabled = IIf(wImportarXls, True, False)
cmdImportar.Enabled = IIf(wImportarXls, True, False)

Txtcodpla.Enabled = IIf(wImportarXls, False, True)
Cmbturno.Enabled = IIf(wImportarXls, False, True)

Me.Height = 9810

Crea_Rs

wciamae = Determina_Maestro("01076")
Sql$ = "Select flag2 from maestros_2 where cod_maestro2='01' and status<>'*'"
Sql$ = Sql$ & wciamae
mHourDay = 0
If (fAbrRst(Rs, Sql$)) Then mHourDay = Val(Rs!flag2)
Rs.Close


'PARA VALIDAR CODIGO DE TRABAJADOR EN LA CARGA DESDE EXCEL

Sql$ = "select prefijo from cia where cod_cia='" & wcia & "' and status<>'*'"
If (fAbrRst(Rs, Sql$)) Then
    If IsNull(Rs!prefijo) Then mPrefijo = "" Else mPrefijo = Rs!prefijo
    If Rs.State = 1 Then Rs.Close
End If


With cbo_TipMotFinPer
    .NameTab = "maestros_2"
    .NameCod = "cod_maestro2"
    .NameDesc = "descrip"
    .Filtro = "RIGHT(ciamaestro,3)='149' and status!='*' and rtrim(isnull(codsunat,''))<>''"
    .conexion = cn
    .Execute
End With

vMesesMinimo = 0: vImporteMinimo = 0


BoletaCargada = False
End Sub
Private Sub Crea_Rs()
    If rshoras.State = 1 Then rshoras.Close
    rshoras.Fields.Append "codigo", adChar, 2, adFldIsNullable
    rshoras.Fields.Append "descripcion", adChar, 100, adFldIsNullable
    rshoras.Fields.Append "monto", adCurrency, 18, adFldIsNullable
    'If Not VNewBoleta Then rshoras.LockType = adLockReadOnly
    rshoras.Open
    Set Dgrdhoras.DataSource = rshoras
    
    If rsccosto.State = 1 Then rsccosto.Close
    rsccosto.Fields.Append "codigo", adChar, 2, adFldIsNullable
    rsccosto.Fields.Append "descripcion", adChar, 45, adFldIsNullable
    rsccosto.Fields.Append "monto", adCurrency, 18, adFldIsNullable
    rsccosto.Fields.Append "item", adChar, 2, adFldIsNullable
    'If Not VNewBoleta Then rsccosto.LockType = adLockReadOnly
    rsccosto.Open
    Set Dgrdccosto.DataSource = rsccosto
    
    If rspagadic.State = 1 Then rspagadic.Close
    rspagadic.Fields.Append "codigo", adChar, 2, adFldIsNullable
    rspagadic.Fields.Append "descripcion", adChar, 45, adFldIsNullable
    rspagadic.Fields.Append "monto", adCurrency, 18, adFldIsNullable
    'If Not VNewBoleta Then rspagadic.LockType = adLockReadOnly
    rspagadic.Open
    Set DgrdPagAdic.DataSource = rspagadic
    
    If rsdesadic.State = 1 Then rsdesadic.Close
    rsdesadic.Fields.Append "codigo", adChar, 2, adFldIsNullable
    rsdesadic.Fields.Append "descripcion", adChar, 45, adFldIsNullable
    rsdesadic.Fields.Append "monto", adCurrency, 18, adFldIsNullable
    'If Not VNewBoleta Then rsdesadic.LockType = adLockReadOnly
    rsdesadic.Open
    Set DgrdDesAdic.DataSource = rsdesadic
End Sub
Private Sub Procesa()

'Pagos Adicionales
If VTipobol = "02" Then
    Sql$ = "Select * from placonstante where cia='" & Trim(wcia) & "' and tipomovimiento='02' and calculo='N' and status<>'*' "
      Sql$ = Sql$ & " UNION select * from placonstante where cia='" & Trim(wcia) & "' and tipomovimiento='02' and status<>'*' and codinterno ='29' " & _
                    "order by codinterno"
Else
    'Sql$ = "Select * from placonstante where cia='" & Trim(wcia) & "' and tipomovimiento='02' and calculo='N' and status<>'*' order by codinterno"
    
    Sql$ = "Select c.codinterno,c.descripcion,v.orden from placonstante c,plaseteoview v "
    Sql$ = Sql$ & "where c.cia='" & Trim(wcia) & "' and c.tipomovimiento='02' and c.calculo='N' and c.status<>'*' and v.tipo='I' and v.tipobol='" & VTipobol & "' and v.status<>'*' "
    Sql$ = Sql$ & "and c.cia=v.cia and v.codigo=c.codinterno "
    
    'MGIRAO AGREGAR SUBSIDIOS
    Sql$ = Sql$ & "UNION ALL Select c.codinterno,case when c.codinterno='43' then  'SUBS. ACC. TRAB.' else  CASE WHEN c.codinterno='32' then  'SUBSIDIO ENFERMEDAD' ELSE c.descripcion end END as descripcion,v.orden from placonstante c,plaseteoview v "
    Sql$ = Sql$ & "where c.cia='" & Trim(wcia) & "' and c.tipomovimiento='02' and c.codinterno in ('03','42','32','43') and c.status<>'*' and v.tipo='I' and v.tipobol='" & VTipobol & "' and v.status<>'*' "
    Sql$ = Sql$ & "and c.cia=v.cia and v.codigo=c.codinterno "
    Sql$ = Sql$ & "order by orden,codinterno"
End If

cn.CursorLocation = adUseClient
Set Rs = New ADODB.Recordset
Set Rs = cn.Execute(Sql$, 64)
If rspagadic.RecordCount > 0 Then
   rspagadic.MoveFirst
   Do While Not rspagadic.EOF
      rspagadic.Delete
      rspagadic.MoveNext
   Loop
End If

If Not Rs.RecordCount > 0 Then Exit Sub
Rs.MoveFirst
Do While Not Rs.EOF
   rspagadic.AddNew
   rspagadic!Codigo = Rs!codinterno
   rspagadic!Descripcion = Rs!Descripcion
   rspagadic!Monto = "0.00"
   Rs.MoveNext
Loop
Rs.MoveFirst
rspagadic.MoveFirst


If Rs.State = 1 Then Rs.Close

'Descuentos Adicionales
If wtipodoc = True Then
   'Activar despues de pruebas
   Sql$ = "Select * from placonstante where cia='" & wcia & "' and tipomovimiento='03' and adicional='S' and status<>'*' order by codinterno"
   'Sql$ = "Select * from placonstante where cia='" & wcia & "' and tipomovimiento='03' and status<>'*' order by codinterno"
Else
   Sql$ = "Select * from placonstante where cia='" & wcia & "' and tipomovimiento='03' and adicional='S' and status<>'*' and codinterno<>'09' order by codinterno"
End If

cn.CursorLocation = adUseClient
Set Rs = New ADODB.Recordset
Set Rs = cn.Execute(Sql$, 64)
If rsdesadic.RecordCount > 0 Then
   rsdesadic.MoveFirst
   Do While Not rsdesadic.EOF
      rsdesadic.Delete
      rsdesadic.MoveNext
   Loop
End If

If Not Rs.RecordCount > 0 Then Exit Sub
Dim RqComision As ADODB.Recordset
Rs.MoveFirst
Do While Not Rs.EOF
   rsdesadic.AddNew
   rsdesadic!Codigo = Rs!codinterno
   rsdesadic!Descripcion = Trim(Rs!Descripcion)
   rsdesadic!Monto = "0.00"
   
   'Mgirao
   'If Trim(rs!Descripcion) = "COMISION" Then
   '     Sql$ = "select importe from plaremunbase where cia='" & wcia & "' and placod='" & Codigo & "' and concepto='01' and status<>'*'"
   '     If fAbrRst(RqComision, Sql) Then
   '          rsdesadic!Monto = fCadNum(RqComision!valorcomision, "###,##0.00")
   '     Else
   '          rsdesadic!Monto = "0.00"
   '     End If
   'End If
   '----------------
   
   Rs.MoveNext
Loop

rsdesadic.MoveFirst
DgrdPagAdic.Row = 0
DgrdPagAdic.Col = 1

If VTipobol = "10" Or VTipobol = "04" Then
'    rsdesadic.AddNew
'    rsdesadic!codigo = "13"
'    rsdesadic!Descripcion = "QUINTA CATEGORIA"
'    rsdesadic!Monto = "0.00"
Else
    Dim rsTemporal As ADODB.Recordset
    'Cadena = "SP_CIA_QUINTA_CAT '" & wcia & "'"
    'Set rsTemporal = OpenRecordset(Cadena, cn)
    'If Not rsTemporal.EOF Then
    '    If MsgBox("Desea Agregar Quinta Categoría?", vbInformation + vbYesNo + vbDefaultButton2, "Planilla") = vbYes Then
    '        rsdesadic.AddNew
    '        rsdesadic!codigo = "13"
    '        rsdesadic!Descripcion = "QUINTA CATEGORIA"
    '        rsdesadic!Monto = "0.00"
    '    End If
    'End If
    
    'If Not rsTemporal Is Nothing Then
    '    If rsTemporal.State = adStateOpen Then rsTemporal.Close
    '    Set rsTemporal = Nothing
    'End If
End If
If Rs.State = 1 Then Rs.Close

End Sub

Private Sub Form_Unload(Cancel As Integer)
FrmCabezaBol.Procesa_Cabeza_Boleta
End Sub

Private Sub LstCcosto_Click()
Dim m As Integer

If LstCcosto.ListIndex > -1 Then
   m = Len(LstCcosto.Text) - 3
    
   Dgrdccosto.Columns(0) = Trim(Left(LstCcosto.Text, m))
   Dgrdccosto.Columns(2) = Format(Right(LstCcosto.Text, 2), "00")
   Dgrdccosto.Col = 1
   Dgrdccosto.SetFocus
   LstCcosto.Visible = False
End If
End Sub

Private Sub LstCcosto_LostFocus()
LstCcosto.Visible = False
End Sub

Private Sub SSCommand1_Click()
Limpia_Boleta
Txtcodpla.SetFocus
End Sub

Private Sub TxtAct1_Change()
TxtAct1.Text = StrConv(Me.TxtAct1.Text, vbUpperCase)
End Sub

Private Sub TxtAct1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Sendkeys "{tab}"
        KeyAscii = 0
    End If
End Sub

Private Sub TxtAct2_Change()
TxtAct2.Text = StrConv(Me.TxtAct2.Text, vbUpperCase)
End Sub

Private Sub TxtAct2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Sendkeys "{tab}"
        KeyAscii = 0
    End If
End Sub

Private Sub TxtAct3_Change()
TxtAct3.Text = StrConv(Me.TxtAct3.Text, vbUpperCase)
End Sub

Private Sub TxtAct3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Sendkeys "{tab}"
        KeyAscii = 0
    End If
End Sub

Private Sub TxtAct4_Change()
TxtAct4.Text = StrConv(Me.TxtAct4.Text, vbUpperCase)
End Sub

Private Sub TxtAct4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Sendkeys "{tab}"
        KeyAscii = 0
    End If
End Sub

Private Sub TxtFecCese_LostFocus()
If TxtFecCese.Text <> "__/__/____" Then
   If Not IsDate(TxtFecCese.Text) Then
      MsgBox "Ingrese Fecha Correctamente", vbCritical, TitMsg
      TxtFecCese.SetFocus
      Call ResaltarTexto(TxtFecCese)
    End If
End If
End Sub

Private Sub Txtcodpla_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 Then
    Frmgrdpla.Show
    Frmgrdpla.ZOrder 0
End If
End Sub

Public Sub Txtcodpla_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then If Cmbturno.Enabled Then Cmbturno.SetFocus
End Sub
Private Sub Txtcodpla_LostFocus()
txtCodPlaLostFocus
DiasAsignacion = 7
If LblId.Caption <> "" Then Exit Sub
If rshoras.RecordCount > o Then
   rshoras.MoveFirst
   Dgrdhoras.Row = 0
   Dgrdhoras.Col = 1
   Dgrdhoras.SetFocus
   DgrdPagAdic.Row = 0
   DgrdPagAdic.Col = 1
End If
End Sub
Private Sub Verifica_Jornal_Minimo(fing As String, fProc As String, Codigo As String, Basico As String, AuxInterno As String)
If vMesesMinimo = 0 Or vImporteMinimo = 0 Then Exit Sub

Dim Dif_Mes_Min As Integer
Dif_Mes_Min = DateDiff("m", CDate(fing), CDate(fProc))

If Day(CDate(fing)) > Day(CDate(fProc)) + 1 Then
 Dif_Mes_Min = Dif_Mes_Min - 1
End If


If Dif_Mes_Min >= 3 And Basico < vImporteMinimo Then
   Mgrab = MsgBox("Sueldo Minimo es menor a " & Str(vImporteMinimo) & Chr(13) & "Fecha de Ingreso: " & fing & Chr(13) & "Desea Actualizarlo", vbYesNo + vbQuestion, TitMsg)
   If Mgrab = 6 Then
      Sql$ = "Update plaremunbase set importe=" & vImporteMinimo & " where cia='" & wcia & "' and placod='" & Codigo & "' and codauxinterno='" & AuxInterno & "' and concepto='01' and status<>'*'"
      cn.Execute Sql$
   End If
   
   Dim Rq As ADODB.Recordset
   Sql$ = "select importe from plaremunbase where cia='" & wcia & "' and placod='" & Codigo & "' and concepto='01' and status<>'*'"
   If fAbrRst(Rq, Sql) Then LblBasico.Caption = fCadNum(Rq!importe, "###,##0.00") Else LblBasico.Caption = ""
    Sql$ = "select importe from plaremunbase where cia='" & wcia & "' and placod='" & Codigo & "' and concepto='02' and status<>'*'"
   If fAbrRst(Rq, Sql) Then txtasigfam.Text = fCadNum(Rq!importe, "###,##0.00") Else txtasigfam.Text = ""
   Rq.Close
End If
End Sub
Private Sub txtCodPlaLostFocus()
Dim xciamae As String
Dim cod As String
Dim rs2 As ADODB.Recordset
Dim RS3 As ADODB.Recordset

On Error GoTo CORRIGE

BolDevengada = False
cn.CursorLocation = adUseClient

Set Rs = New ADODB.Recordset
cod = "01055"
Txtcodpla.Text = UCase(Txtcodpla.Text)
If Trim(Txtcodpla.Text & "") = "" Then Limpia_Boleta: Exit Sub
vPlaCod = Trim(Txtcodpla.Text)
' TIPO DE TRABAJADOR
Sql$ = "SELECT GENERAL FROM maestros where " & _
"right(ciamaestro,3)='" & Right(cod, 3) & "' and " & _
"status<>'*' "

If (fAbrRst(Rs, Sql$)) Then
   If Rs!General = "S" Then
      xciamae = " and right(ciamaestro,3)= '" & Right(cod, 3) & "'"
   Else
      xciamae = " and ciamaestro= '" & wcia + Right(cod, 3) & "'"
   End If
End If

If Rs.State = 1 Then Rs.Close
'MGIRAO PRACTICANTES
If vTipoTra = "05" Then Frmboleta.DgrdPagAdic.Enabled = True Else Frmboleta.DgrdPagAdic.Enabled = True
'OBTENER NOMBRE DE EMPLEADO
Sql$ = Funciones.nombre()
Sql$ = Sql$ & "codauxinterno,a.status,a.tipotrabajador,a.fingreso," & _
     "a.fcese,a.mot_fin_periodo,a.codafp,a.afptipocomision,a.numafp,a.area,a.placod," & _
     "a.codauxinterno,b.descrip,a.tipotasaextra," & _
     "a.cargo,a.altitud,a.vacacion,a.area,a.fnacimiento," & _
     "a.fec_jubila,a.txtrajub,a.sindicato,a.ESSALUDVIDA,a.quinta,a.CodTipoTrabSunat,a.domiciliado,a.codsctr,a.planta,a.cod_area " & _
     "from planillas a,maestros_2 b where a.status<>'*' "
     Sql$ = Sql$ & xciamae
     Sql$ = Sql$ & " and a.tipotrabajador=b.cod_maestro2 " _
     & "and cia='" & wcia & "' AND placod='" & Trim(Txtcodpla.Text) & "' "
     Sql$ = Sql$ & " order by nombre"

cn.CursorLocation = adUseClient
Set Rs = New ADODB.Recordset
Set Rs = cn.Execute(Sql$)

'Verificar Factores para SCRT
If Trim(Txtcodpla.Text & "") <> "" Then
   If Trim(Rs!codsctr & "") <> "00" Then
      Dim rsVeriSCRT As New ADODB.Recordset
      Sql$ = "usp_Pla_Verifica_Factores_SCRT '" & wcia & "','" & (Mid(VFProceso, 7, 4)) + (Mid(VFProceso, 4, 2)) & "','" & Trim(Rs!codsctr & "") & "'"
      cn.CursorLocation = adUseClient
      Set rsVeriSCRT = New ADODB.Recordset
      Set rsVeriSCRT = cn.Execute(Sql$)
      If rsVeriSCRT.RecordCount > 0 Then
         If Trim(rsVeriSCRT!resultado & "") <> "OK" Then
            MsgBox "Debe Registrar Factores Para SCRT" & Chr(13) & "Trabajador => " & Trim(Txtcodpla.Text), vbInformation
            Limpia_Boleta
            rsVeriSCRT.Close: Set rsVeriSCRT = Nothing
            Exit Sub
        End If
      Else
         MsgBox "Debe Registrar Factores Para SCRT" & Chr(13) & "Trabajador => " & Trim(Txtcodpla.Text), vbInformation
         Limpia_Boleta
         rsVeriSCRT.Close: Set rsVeriSCRT = Nothing
         Exit Sub
      End If
      rsVeriSCRT.Close: Set rsVeriSCRT = Nothing
      End If
End If


Es_Cesado = False
If Not IsNull(Rs!fcese) Then Es_Cesado = True
If VTipobol = "04" Then
   If Rs.RecordCount > 0 Then
        If IsNull(Rs!fcese) Then
            If bImportar = False Then
              MsgBox "Trabajador no es Cesado", vbInformation
            Else
               LstObs.AddItem Trim(Txtcodpla.Text) & ": Trabajador no es Cesado"
            End If
            
            Limpia_Boleta
            Exit Sub
        End If
        If Month(Rs!fcese) <> Vmes Or Year(Rs!fcese) <> Vano Then
            If bImportar = False Then
                MsgBox "Trabajador Cesado en otro Periodo", vbInformation
            Else
               LstObs.AddItem Trim(Txtcodpla.Text) & ": Trabajador no es Cesado"
            End If
            
            Limpia_Boleta
            Exit Sub
        End If
    End If
End If

If Rs.RecordCount > 0 Then
   If Rs!TipoTrabajador <> vTipoTra Then
      If bImportar = False Then
         MsgBox "Trabajador no es del tipo seleccionado", vbExclamation, "Codigo N° => " & Txtcodpla.Text
      Else
         LstObs.AddItem Trim(Txtcodpla.Text) & ": Trabajador no es del tipo seleccionado"
      End If
      Txtcodpla.Text = ""
      Limpia_Boleta
      Exit Sub
   End If

   If Not IsNull(Rs!fcese) Then
      Dim mFechaC As Date
      mFechaC = DateAdd("d", -7, CDate(VFProceso))
      If CDate(mFechaC) > CDate(Rs!fcese) Then
          If bImportar = False Then
               MsgBox "Trabajador " & Trim(Txtcodpla.Text) & " ya fue Cesado", vbExclamation, "Con Fecha => " & Format(Rs!fcese, "dd/mm/yyyy")
          Else
               LstObs.AddItem Trim(Txtcodpla.Text) & ": Trabajador ya fue Cesado"
          End If
         Txtcodpla.Text = ""
         Limpia_Boleta
         Exit Sub
      End If
   End If
      
   LblFingreso.Caption = Format(Rs!fIngreso, "mm/dd/yyyy")
   LblFechaIngreso.Caption = Format(Rs!fIngreso, "dd/mm/yyyy")
   If Val(Right(LblFingreso.Caption, 4)) > Vano Then
      If bImportar = False Then
         MsgBox "Fecha de Ingreso del Trabajador es Superior", vbCritical, "Calculo de Boletas"
      Else
         LstObs.AddItem Trim(Txtcodpla.Text) & ": Fecha de Ingreso del Trabajador es Superior"
      End If
        
      Limpia_Boleta
      LblFingreso.Caption = ""
      Exit Sub
   ElseIf Val(Right(LblFingreso.Caption, 4)) = Vano And Val(Left(LblFingreso.Caption, 2)) > Vmes Then
      If bImportar = False Then
         MsgBox "Fecha de Ingreso del Trabajador es Superior", vbCritical, "Calculo de Boletas"
      Else
         LstObs.AddItem Trim(Txtcodpla.Text) & ": Fecha de Ingreso del Trabajador es Superior"
      End If
         
      Limpia_Boleta
      LblFingreso.Caption = ""
      Exit Sub
   ElseIf Val(Right(LblFingreso.Caption, 4)) = Vano And Val(Left(LblFingreso.Caption, 2)) = Vmes And Val(Mid(LblFingreso.Caption, 4, 2)) > Val(Left(VFProceso, 2)) Then
      If bImportar = False Then
         MsgBox "Fecha de Ingreso del Trabajador es Superior", vbCritical, "Calculo de Boletas"
      Else
         LstObs.AddItem Trim(Txtcodpla.Text) & ": Fecha de Ingreso del Trabajador es Superior"
      End If
         
      Limpia_Boleta
      LblFingreso.Caption = ""
      Exit Sub
   End If
      
   If VTipobol <> "04" Then If Not VerificaGrati(Txtcodpla.Text, Val(Mid(VFProceso, 4, 2))) Then Exit Sub

   'cargamos si la person esta afecto o no a la quinta categoria
   sn_quinta = True
   sn_Domiciliado = True
   If Trim(Rs!quinta & "") = "N" Or Trim(Rs!quinta & "") = "" Then sn_quinta = False
   If Not Rs!domiciliado Then sn_Domiciliado = False
   '***********************************************************
   If Not sn_quinta Then
      MsgBox "TRABAJADOR SETEADO PARA NO CALCULAR QUINTA CATEGORIA POR EL SISTEMA", vbInformation
   End If
   
   Lblnombre.Caption = Rs!nombre
   Lblcodaux.Caption = Rs!codauxinterno
   Lblcodafp.Caption = Rs!CodAfp
   lblAfpTipoComision.Caption = Rs!afptipocomision
   LblPlanta.Caption = Rs!Planta
   Nueva_area = Rs!cod_area
   'Nombre de AFP
   Sql$ = "select descrip from maestros_2 where ciamaestro='01069' and cod_maestro2='" & Rs!CodAfp & "'"
   If (fAbrRst(rs2, Sql$)) Then LblPensiones.Caption = rs2!DESCRIP Else LblPensiones.Caption = ""
   If rs2.State = 1 Then rs2.Close
      
   LblCodTrabSunat = Trim(Rs!CodTipoTrabSunat & "")
   Lblnumafp.Caption = Trim(Rs!NUMAFP)
   VFechaNac = Format(Rs!fnacimiento, "dd/mm/yyyy")
   VFechaJub = Format(Rs!fec_jubila, "dd/mm/yyyy")
   VJubilado = Trim(Rs!txtrajub)
   Lbltope.Caption = Rs!tipotasaextra
   If Trim(Rs!fcese & "") <> "" Then
      TxtFecCese.Text = Format(Day(Rs!fcese), "00") & "/" & Format(Month(Rs!fcese), "00") & "/" & Format(Year(Rs!fcese), "0000")
      cbo_TipMotFinPer.SetIndice Trim(Rs!mot_fin_periodo & "")
      TxtFecCese.Enabled = False
      cbo_TipMotFinPer.Enabled = False
   End If
   
   If vTipoTra = "05" Then Lblcargo.Caption = Rs!Cargo: VAltitud = Rs!altitud: VVacacion = Rs!vacacion
      
   VArea = Trim(Rs!Area)
   Frame2.Enabled = True
   Frame3.Enabled = True
   Frame4.Enabled = True
   Frame5.Enabled = True
      
   'Basico  de la persona
    txtasigfam.Text = "0.00"
    Sql$ = "select importe from plaremunbase where cia='" & wcia & "' and placod='" & Trim(Txtcodpla.Text) & "' and concepto='02' and status<>'*'"
   If fAbrRst(rs2, Sql$) Then txtasigfam.Text = rs2!importe Else txtasigfam.Text = "0.00"
   
   
   Sql$ = "select importe from plaremunbase where cia='" & wcia & "' " & _
   "and placod='" & Trim(Txtcodpla.Text) & "' and concepto='01' and status<>'*'"
      
   LblBasico.Caption = ""
   If (fAbrRst(rs2, Sql$)) Then LblBasico.Caption = fCadNum(rs2!importe, "###,##0.00")
   
 'add jcms 110122 se deshabilita por indicacion de O. Huancaya+ A. Gonzalez
'   Call Verifica_Jornal_Minimo(Format(rs!fIngreso, "dd/mm/yyyy"), VFProceso, Trim(Txtcodpla.Text), rs2!importe, Trim(rs!codauxinterno & ""))
'   If rs2.State = 1 Then rs2.Close
   
   'Centro de Costo
   If rsccosto.RecordCount > 0 Then rsccosto.MoveFirst
   Do While Not rsccosto.EOF
      rsccosto.Delete
      rsccosto.MoveNext
   Loop
   Dim RX As New ADODB.Recordset
   sn_essaludvida = 0: snmanual = 0
       
   If Trim(UCase(Rs("ESSALUDVIDA"))) = "S" And VTipobol <> "10" Then
      rsdesadic.MoveFirst
      sn_essaludvida = 1
      rsdesadic.FIND "CODIGO='06'", 1 'ESSALUDVIDA
      con = "select DEDUCCION,adicional from placonstante where cia='" & wcia & "' and tipomovimiento='03' and " & _
      "codinterno='06' and deduccion<>0 and status<>'*'"
      RX.Open con, cn, adOpenStatic, adLockReadOnly
      If RX.RecordCount > 0 Then If Not rsdesadic.EOF Then rsdesadic("MONTO") = RX("DEDUCCION"): snmanual = IIf(RX!adicional = "S", 0, 1)
      RX.Close
      Call Acumula_Mes("06", "D")
      If rsdesadic("MONTO") > 0 Then rsdesadic("MONTO") = rsdesadic("MONTO") - macui
   End If
    
   MSINDICATO = False
   sn_sindicato = 0
      
   If Trim(UCase(Rs("sindicato"))) = "S" Then
      rsdesadic.FIND "CODIGO='20'", 1 'SOLIC SINDICATO
      Dim VSemanaSindicato As String
      VSemanaSindicato = VSemana
      If Trim(vTipoTra) = "01" Then
          VSemanaSindicato = Vmes
      End If
      
      Sql$ = "usp_Pla_ConsultarImportePlaSolicitudSindicato '" & wcia & "','" & Trim(VTipobol)
      Sql$ = Sql$ & "','" & Trim(vTipoTra) & "','" & Trim(Txtcodpla.Text) & "','"
      Sql$ = Sql$ & Mid(VFProceso, 7, 4) & "','" & VSemanaSindicato & "'"
         
      RX.Open Sql$, cn, adOpenStatic, adLockReadOnly
      If RX.RecordCount > 0 Then If Not rsdesadic.EOF Then rsdesadic("MONTO") = RX("importe")
      MSINDICATO = True
      sn_sindicato = 1
      RX.Close
   End If
    
   Dim mNumCcosto As Integer
   Dim mTotCcosto As Double
   Dim mArecCosto As String

   mNumCcosto = 0: mTotCcosto = 0

'   Sql$ = "select a.ccosto as cod_maestro3,b.descrip,a.porc from planilla_ccosto a"
'   Sql$ = Sql$ & " inner join maestros_32 b"
'   Sql$ = Sql$ & "     on  (a.cia+'055')=b.ciamaestro and b.cod_maestro2='" & vTipoTra & "' and"
'   Sql$ = Sql$ & "          A.ccosto = cod_maestro3"
'   Sql$ = Sql$ & " where a.cia='" & wcia & "' and a.placod='" & Trim(Txtcodpla.Text) & "' and a.status !='*' and b.status !='*'"
'   Set rs = cn.Execute(Sql$)

   Sql$ = "select a.ccosto as cod_maestro3,b.descripcion as descrip,a.porc from planilla_ccosto a"
   Sql$ = Sql$ & " inner join PLA_CCOSTOS b"
   Sql$ = Sql$ & " on  a.cia=b.cia and "
   Sql$ = Sql$ & " A.ccosto = b.codigo "
   Sql$ = Sql$ & " where a.cia='" & wcia & "' and a.placod='" & Trim(Txtcodpla.Text) & "' and a.status !='*' and b.status !='*'"
   Set Rs = cn.Execute(Sql$)
    
   If rsccosto.RecordCount > 0 Then rsccosto.MoveFirst
   Do While Not rsccosto.EOF
      rsccosto.Delete
      rsccosto.MoveNext
   Loop
   Dim mSemana As String
   If Not VNewBoleta Then
      mSemana = FrmCabezaBol.Txtsemana.Text
      Dim Rs_tmp As New ADODB.Recordset
      Dim query As String
    
      query = "select ccosto1,PORC1,ccosto2,PORC2,ccosto3,PORC3,ccosto4,PORC4,ccosto5,PORC5 "
      query = query & "from plahistorico where id_boleta=" & Val(LblId.Caption) & ""

      Rs_tmp.Open query, cn, adOpenStatic, adLockOptimistic
      If Rs_tmp.RecordCount > 0 Then
         Dim J As Integer
         For J = 1 To 5
           Dim cc As String
            Dim por As Integer
            cc = Rs_tmp.Fields("ccosto" & J)
            por = Rs_tmp.Fields("porc" & J)
            If Trim(cc) <> "" And por <> 0 Then
               rsccosto.AddNew
               rsccosto!Codigo = cc
               Dim Qtmp As String
               Dim Rs_cc_t As New ADODB.Recordset
               'Qtmp = "Select cod_maestro3,descrip from maestros_32 where status<>'*'  "
               'Qtmp = Qtmp & "and ciamaestro= '" & wcia & "055'  and cod_maestro2='" & vTipoTra & "' and cod_maestro3='" & cc & "'  "
               
               Qtmp = "Select codigo as cod_maestro3,descripcion as descrip from PLA_CCOSTOS where status<>'*'  "
               Qtmp = Qtmp & "and cia= '" & wcia & "'  and codigo='" & cc & "'  "

               Set Rs_cc_t = cn.Execute(Qtmp)
               If Rs_cc_t.RecordCount > 0 Then rsccosto!Descripcion = Rs_cc_t("descrip") Else rsccosto!Descripcion = "XXX" & J 'UCase(rs!descrip)
               Rs_cc_t.Close: Set Rs_cc_t = Nothing
               rsccosto!Monto = por
               lbltot.Caption = "100.00"
               rsccosto!Item = vItem
               rsccosto.Update
            End If
         Next
      Else
Centro_Costo:
         'LLENA GRID CENTRO DE COSTOS
         If Rs.RecordCount > 0 Then
            Rs.MoveFirst
            Do While Not Rs.EOF
               rsccosto.AddNew
               rsccosto!Codigo = Trim(Rs!COD_MAESTRO3)
               rsccosto!Descripcion = UCase(Rs!DESCRIP)
               rsccosto!Monto = Rs!PORC
               mTotCcosto = mTotCcosto + Rs!PORC
               rsccosto.Update
               Dgrdccosto.Refresh
               mNumCcosto = mNumCcosto + 1
               Rs.MoveNext
            Loop
            lbltot.Caption = mTotCcosto
            rsccosto!Item = vItem
         End If
      End If
   Else
      If Rs.RecordCount > 0 Then
         Rs.MoveFirst
         Do While Not Rs.EOF
            rsccosto.AddNew
            rsccosto!Codigo = Trim(Rs!COD_MAESTRO3)
            rsccosto!Descripcion = UCase(Rs!DESCRIP)
            rsccosto!Monto = Rs!PORC
            mTotCcosto = mTotCcosto + Rs!PORC
            rsccosto.Update
            Dgrdccosto.Refresh
            mNumCcosto = mNumCcosto + 1
            Rs.MoveNext
         Loop
         lbltot.Caption = mTotCcosto
         rsccosto!Item = vItem
      End If
   End If
   For I = I To 5 - mNumCcosto
       If rsccosto.RecordCount < 5 Then rsccosto.AddNew
   Next I
    
   rsccosto.MoveFirst
   Dgrdccosto.Refresh
            
  'Activos Asignados solo para boleta Normal
   If VTipobol = "01" Then
      TxtAct1.Text = "": TxtAct2.Text = "": TxtAct3.Text = "": TxtAct4.Text = "": TxtHor1.Text = 0: TxtHor2.Text = 0: TxtHor3.Text = 0: TxtHor4.Text = 0
      query = "Select Activo,Activo2,Activo3,Activo4,HoraAsigAct1,HoraAsigAct2,HoraAsigAct3,HoraAsigAct4 From Pla_Trab_Activo where cia='" & wcia & "' and ayo=" & Vano & " and mes=" & Vmes & " and semana='" & VSemana & "' and placod='" & Trim(Txtcodpla.Text) & "' and status<>'*'"
      cn.CursorLocation = adUseClient
      Set Rs = New ADODB.Recordset
      Set Rs = cn.Execute(query)
      If Rs.RecordCount > o Then
         TxtAct1.Text = Trim(Rs!Activo & "")
         TxtAct2.Text = Trim(Rs!Activo2 & "")
         TxtAct3.Text = Trim(Rs!Activo3 & "")
         TxtAct4.Text = Trim(Rs!Activo4 & "")
         TxtHor1.Text = Trim(Rs!HoraAsigAct1 & "")
         TxtHor2.Text = Trim(Rs!HoraAsigAct2 & "")
         TxtHor3.Text = Trim(Rs!HoraAsigAct3 & "")
         TxtHor4.Text = Trim(Rs!HoraAsigAct4 & "")
      Else
         query = "Select Top 1 Activo,Activo2,Activo3,Activo4,HoraAsigAct1,HoraAsigAct2,HoraAsigAct3,HoraAsigAct4 From Pla_Trab_Activo where cia='" & wcia & "' and status<>'*' and placod='" & Trim(Txtcodpla.Text) & "' order by ayo,mes,semana desc"
         cn.CursorLocation = adUseClient
         Set Rs = New ADODB.Recordset
         Set Rs = cn.Execute(query)
         If Rs.RecordCount > o Then
            TxtAct1.Text = Trim(Rs!Activo & "")
            TxtAct2.Text = Trim(Rs!Activo2 & "")
            TxtAct3.Text = Trim(Rs!Activo3 & "")
            TxtAct4.Text = Trim(Rs!Activo4 & "")
            TxtHor1.Text = Trim(Rs!HoraAsigAct1 & "")
            TxtHor2.Text = Trim(Rs!HoraAsigAct2 & "")
            TxtHor3.Text = Trim(Rs!HoraAsigAct3 & "")
            TxtHor4.Text = Trim(Rs!HoraAsigAct4 & "")
         End If
      End If
   End If
   
   If Rs.State = 1 Then Rs.Close
   
   Txtcodpla.Enabled = False
       
ElseIf Trim(Txtcodpla.Text) <> "" Then
    If bImportar = False Then
          MsgBox "Codigo de Planilla No Existe ", vbExclamation, "Codigo N° => " & Txtcodpla.Text
    Else
          LstObs.AddItem Trim(Txtcodpla.Text) & ": Codigo de Planilla No Existe"
    End If
         
   Txtcodpla.Text = ""
   Limpia_Boleta
   Lblnombre.Caption = ""
   Lblctacte.Caption = "0.00"
   Lblcodaux.Caption = ""
   Lblcodafp.Caption = ""
   LblPensiones.Caption = ""
   lblAfpTipoComision.Caption = ""
   LblCodTrabSunat.Caption = ""
   Lblnumafp.Caption = ""
   LblBasico.Caption = ""
   LblFingreso.Caption = ""
   Lbltope.Caption = ""
   Lblcargo.Caption = ""
   VAltitud = ""
   VVacacion = ""
   VArea = ""
   VFechaNac = ""
   VFechaJub = ""
   Txtcodpla.SetFocus
   Frame2.Enabled = False
   Frame3.Enabled = False
   Frame4.Enabled = False
   Frame5.Enabled = False
End If
vItem = 0
If VTipobol <> "01" Then Dgrdhoras.Columns(1).Locked = False

Call Carga_Horas

If wGrupoPla <> "01" Then 'Calculos de promedios es detallado en Gallos
   'If VTipobol = "02" Or VTipobol = "03" Then Otros_Pagos_Vac
   If VTipobol = "02" Then Otros_Pagos_Vac (False)
End If
'---Incluye adelanto quincena Subsidio
If wtipodoc = True And (VTipobol = "01" Or VTipobol = "05") And Trim(Txtcodpla.Text & "") <> "" Then
   If rsdesadic.RecordCount > 0 Then rsdesadic.MoveFirst
   Do While Not rsdesadic.EOF
      If rsdesadic!Codigo = "09" Then
         Sql$ = "select totneto from plaquincena where cia='" & wcia & "' and placod='" & Txtcodpla.Text & "'  and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " and status<>'*' "
         '        Sql$ = "select 0"
         'mgirao
         If (fAbrRst(rs2, Sql$)) Then
            rsdesadic!Monto = rs2(0)
         End If
         rs2.Close
      End If
      rsdesadic.MoveNext
   Loop
End If



'AGREGADO A LA FUERZA 2011.03.09
Dim Cadena As String
Dim rsVal As ADODB.Recordset
Cadena = "select d06 as Contador from plahistorico where cia = '" & wcia & "' and year(fechaproceso) = " & Year(VFProceso) & " and month(fechaproceso) = " & Month(VFProceso) & " and placod = '" & Trim(Txtcodpla.Text) & "' and d06<>0 and status <> '*'"
Set rsVal = OpenRecordset(Cadena, cn)
If Not rsVal.EOF Then
If rsdesadic.RecordCount > 0 Then rsdesadic.MoveFirst
Do While Not rsdesadic.EOF
    If rsdesadic!Codigo = "06" Then
        If Not rsVal.EOF Then
            If rsVal!contador > 0 Then
                rsdesadic("MONTO") = 0
            End If
        End If
    End If
    rsdesadic.MoveNext
Loop
End If

'pruena
'Carga Saldo Cuenta Corriente
Sql = "Usp_Pla_Carga_Saldos_Ctacte '" & wcia & "','" & vTipoTra & "','" & Txtcodpla.Text & "','N'"
If (fAbrRst(Rs, Sql$)) Then
   If IsNull(Rs(0)) Then Lblctacte.Caption = "0.00" Else Lblctacte.Caption = Format(Rs!saldo, "###,###,###.00")
End If
Rs.Close

'Carga Cuenta Descuento Corriente
Dim lDctoCtaCte As Double
lDctoCtaCte = 0
Sql = "select Top 1 normal,vaca,grati from pla_dcto_ctacte where cia='" & wcia & "' and placod='" & Txtcodpla.Text & "' and fec_desde<='" & Format(VFProceso, FormatFecha) + " 00:00:00" & "'  and status<>'*' order by fec_crea desc"
If (fAbrRst(Rs, Sql$)) Then
   If VTipobol = "01" Then lDctoCtaCte = Rs!Normal
   If VTipobol = "02" Then lDctoCtaCte = Rs!vaca
   If VTipobol = "03" Then lDctoCtaCte = Rs!grati
   If lDctoCtaCte > CCur(Lblctacte.Caption) Then lDctoCtaCte = CCur(Lblctacte.Caption)
   
   If lDctoCtaCte > 0 Then
      If rsdesadic.RecordCount > 0 Then rsdesadic.MoveFirst
      Do While Not rsdesadic.EOF
         If rsdesadic("CODIGO") = "07" Then
            rsdesadic("MONTO") = lDctoCtaCte
            Exit Do
         End If
         rsdesadic.MoveNext
      Loop
   End If
End If
Rs.Close

If VTipobol = "02" And vDevengue = False Then
   Sql = "select i16 from plahistorico where cia='" & wcia & "' and proceso='02' and year(fechaproceso)<=" & Val(Mid(VFProceso, 7, 4)) & " and month(fechaproceso)=1 and status='D' and placod='" & Txtcodpla.Text & "' order by turno"
   If (fAbrRst(Rs, Sql$)) Then
      Rs.MoveFirst
      If rspagadic.RecordCount > 0 Then rspagadic.MoveFirst
      Do While Not rspagadic.EOF
         rspagadic.Delete
         rspagadic.MoveNext
      Loop
      If Rs!I16 <> 0 Then
         rspagadic.AddNew
         rspagadic!Codigo = "16"
         rspagadic!Monto = Rs!I16
         rspagadic!Descripcion = "OTROS PAGOS"
      End If
      MsgBox "Trabajador Tiene Vacaciones Devengadas" & Chr(13) & "No podra modificar Datos, Solo Grabar", vbInformation, "Vacaciones Devengadas"
      BolDevengada = True
      Frame2.Enabled = False
      Frame3.Enabled = False
      Frame4.Enabled = False
      Frame5.Enabled = False
   End If
End If
rspagadic.MoveFirst
DgrdPagAdic.Row = 0
DgrdPagAdic.Col = 1


If Verifica_Boleta(False) = True Then
   If VPerPago = "02" And VTipobol = "01" Then Carga_Marcacion
End If

LbLSemBolVac = ""
If VTipobol = "01" Then 'Carga Fecha de Boleta de vacaciones y determinar si la boleta de la semana es igual a la semana de la boleta de vacaciones
   Dim lFecVaca As String
   lFecVaca = ""
   Sql$ = "Select fechaproceso from plahistorico where cia='" & wcia & "' AND year(fechaproceso)=" & Val(Mid(VFProceso, 7, 4)) & " and proceso='02' and placod='" & Txtcodpla.Text & "' and status<>'*' order by fechaproceso desc"
   If (fAbrRst(Rs, Sql$)) Then lFecVaca = Format(Rs!FechaProceso, "dd/mm/yyyy")
   Rs.Close
   
   If lFecVaca <> "" Then
      SqlFec = "select semana from plasemanas where cia='" & wcia & "' and ano=" & Vano & " and status<>'*' and '" & Format(lFecVaca, FormatFecha) & "' between fechai and fechaf"
      If (fAbrRst(Rs, SqlFec)) Then LbLSemBolVac.Caption = Trim(Rs!semana & "")
   End If
   If LbLSemBolVac.Caption <> VSemana Then LbLSemBolVac.Caption = ""
End If

Exit Sub
CORRIGE:
  MsgBox "Error:" & Err.Description, vbCritical, Me.Caption
End Sub
Private Sub Carga_Marcacion()

Dim I As Integer
Dim lSinHoras As String
lSinHoras = "N"
Sql$ = "Usp_Pla_Traer_Marcacion '" & wcia & "'," & Vano & "," & Vmes & ",'" & VSemana & "','" & Txtcodpla.Text & "' "

cn.CursorLocation = adUseClient
Set Rs = New ADODB.Recordset
Set Rs = cn.Execute(Sql$, 64)
If Not Rs.EOF Then
   If rshoras.RecordCount > 0 Then rshoras.MoveFirst
   Do While Not rshoras.EOF
      For I = 0 To Rs.Fields.count - 1
          If Left(Rs.Fields(I).Name, 1) = "H" Then
             If Mid(Rs.Fields(I).Name, 2, 2) = rshoras!Codigo Then
                rshoras!Monto = Rs.Fields(I)
             End If
          End If
      Next
      If Rs!h01 = 0 And rshoras!Codigo = "01" Then lSinHoras = "S"
      rshoras.MoveNext
   Loop
Else
  'mgirao cambio para solo quinta
   'MsgBox "Trabajador No registra marcaciones", vbInformation
   I = I
   
End If
If Rs.RecordCount > o Then Rs.MoveFirst
If Not Rs.EOF Then
   If rspagadic.RecordCount > 0 Then rspagadic.MoveFirst
   Do While Not rspagadic.EOF
      For I = 0 To Rs.Fields.count - 1
          If Left(Rs.Fields(I).Name, 1) = "I" Then
             If Mid(Rs.Fields(I).Name, 2, 2) = rspagadic!Codigo Then
                rspagadic!Monto = Rs.Fields(I)
             End If
          End If
      Next
      rspagadic.MoveNext
   Loop
End If

If Rs.RecordCount > o Then Rs.MoveFirst
If Not Rs.EOF Then
   If rsdesadic.RecordCount > 0 Then rsdesadic.MoveFirst
   Do While Not rsdesadic.EOF
      For I = 0 To Rs.Fields.count - 1
          If Left(Rs.Fields(I).Name, 1) = "D" Then
             If Mid(Rs.Fields(I).Name, 2, 2) = rsdesadic!Codigo Then
                rsdesadic!Monto = Rs.Fields(I)
             End If
          End If
      Next
      rsdesadic.MoveNext
   Loop
End If

If Rs.State = 1 Then Rs.Close
If lSinHoras = "S" Then MsgBox "Trabajador not tiene horas trabajadas", vbInformation
End Sub
Public Sub Carga_Boleta(Codigo As String, tipo As String, Nuevo As Boolean, semana As String, fproce As String, Tipot As String, perpago As String, horas As Integer, mdel As String, mal As String, obra As String, devengue As Boolean, lId As String)
Dim MField As String
Dim wciamae As String
Load Frmboleta
LblId.Caption = lId
VTipobol = tipo
LblPlanta.Caption = obra
vTipoTra = Tipot
VSemana = semana
VFProceso = fproce
VPerPago = perpago
VNewBoleta = Nuevo
VHoras = horas
Vano = Val(Mid(VFProceso, 7, 4))
Vmes = Val(Mid(VFProceso, 4, 2))
Vdia = Val(Mid(VFProceso, 1, 2))
VfDel = mdel
VfAl = mal
LstCcosto.Clear
'wciamae = Funciones.Determina_Maestro("01044")
'Sql$ = "Select cod_maestro2,descrip from maestros_2 where  status<>'*'"
'Sql$ = Sql$ & wciamae

'Sql$ = "Select cod_maestro3 as COD_MAESTRO2,descrip from maestros_32 where ciamaestro='" & wcia & "055' and "
'Sql$ = Sql$ & "cod_maestro2='" & Trim(vTipoTra) & "' and status<>'*' ORDER BY 2"

Sql$ = "select codigo,descripcion From Pla_ccostos where cia='" & wcia & "' and status<>'*' order by descripcion"
WValor03 = 0#
WValor32 = 0#
WValor42 = 0#
wValor43 = 0#
t_horas_trabajadas = 0#
cn.CursorLocation = adUseClient
Set Rs = New ADODB.Recordset
Set Rs = cn.Execute(Sql$, 64)
If Rs.RecordCount > 0 Then Rs.MoveFirst
Do Until Rs.EOF
   'LstCcosto.AddItem rs!DESCRIP & Space(100) & rs!COD_MAESTRO2
   LstCcosto.AddItem Rs!Descripcion & Space(100) & Rs!Codigo
   Rs.MoveNext
Loop
If Rs.State = 1 Then Rs.Close

'Calculo se Retencion Judicial
Dim lRetJud As Boolean
lRetJud = False

If VTipobol = "11" Then
   Sql$ = "select importe,status from pladeducper where cia='" & wcia & "' and placod='" & Trim(Codigo) & "' and concepto='05' and status<>'*' and placod in(select placod from PlaRetJudUti where cia='" & wcia & "' and status<>'*')"
Else


   Sql$ = "select importe,status from pladeducper where cia='" & wcia & "' and placod='" & Trim(Codigo) & "' and concepto='05' and status<>'*'"
End If
If (fAbrRst(Rs, Sql$)) Then
   If Not IsNull(Rs(0)) Then If Rs(0) > 0 Then lRetJud = True
End If
If Rs.State = 1 Then Rs.Close

'Turno
Sql$ = "Select codturno,descripcion from platurno where cia='" & wcia & "' and status<>'*' order by codturno"
Cmbturno.Clear
VTurno = ""

If (fAbrRst(Rs, Sql$)) Then
   If (Not Rs.EOF) Then
      Do Until Rs.EOF
         Cmbturno.AddItem Rs(1)
         Cmbturno.ItemData(Cmbturno.NewIndex) = Rs(0)
         Rs.MoveNext
       Loop
       Cmbturno.ListIndex = 0
    End If
    If Rs.State = 1 Then Rs.Close
End If
If VTipobol = "01" Then PnlActivos.Visible = True Else PnlActivos.Visible = False
Select Case VTipobol 'VACACIONES
       Case Is = "02"
            Lblvaca1.Visible = True
            Txtvacai.Visible = True
            Lblvaca2.Caption = "F. Inicio Vaca."
            Txtvacaf.Visible = True
            Label4.Visible = True
            TxtVaca.Visible = True
            Lblvaca2.Visible = True
            'Lblcese.Visible = False
            'Txtcese.Visible = False
       Case Is = "04"
            Lblvaca1.Visible = False
            Txtvacai.Visible = False
            'Txtvacaf.Visible = False
            Label4.Visible = False
            TxtVaca.Visible = False
            Lblvaca2.Visible = False
            Lblcese.Visible = True
            TxtFecCese.Visible = True
       Case Else
            Lblvaca1.Visible = False
            Txtvacai.Visible = False
            'Txtvacaf.Visible = False
            Label4.Visible = False
            TxtVaca.Visible = False
            'Lblcese.Visible = False
End Select
Procesa
If VNewBoleta = False Then
   Txtcodpla.Text = Codigo
   Txtcodpla_LostFocus
   If rshoras.RecordCount > 0 Then rshoras.MoveFirst
   Do While Not rshoras.EOF
      rshoras!Monto = "0"
      rshoras.MoveNext
   Loop
   If VNewBoleta = False Then MsgBox "Los datos solo se mostraran como consulta " & Chr(13) & "Si desea modificar algun dato debera anular la boleta", vbInformation, "Sistema de Planilla"
   Sql$ = ""
   If wtipodoc = True Then
'      Select Case VPerPago
'             Case Is = "02"
'                  Sql$ = "select * from plahistorico " _
'                     & "where cia='" & wcia & "' and proceso='" & VTipobol & "' and semana='" & VSemana & "' and year(fechaproceso)=" & Vano & " " _
'                     & "and status<>'*' and placod='" & Txtcodpla.Text & "'"
'             Case Is = "04"
'                  Sql$ = "select * from plahistorico " _
'                     & "where cia='" & wcia & "' and proceso='" & VTipobol & "' and year( fechaproceso) = " & Vano & " And Month(fechaproceso) = " & Vmes & " " _
'                     & "and status<>'*' and placod='" & Txtcodpla.Text & "'"
'      End Select
      Sql$ = "select * from plahistorico where id_boleta=" & Val(LblId.Caption) & ""
   Else
      Sql$ = "select * from plaquincena " _
         & "where cia='" & wcia & "' and year( fechaproceso) = " & Vano & " And Month(fechaproceso) = " & Vmes & " " _
         & "and status<>'*' and placod='" & Txtcodpla.Text & "'"
       '        Sql$ = "select 0"
       '  mgirao
   End If
   cn.CursorLocation = adUseClient
   Set Rs = New ADODB.Recordset
   Set Rs = cn.Execute(Sql$, 64)
   If Rs.RecordCount > 0 Then
      Call rUbiIndCmbBox(Cmbturno, Format(Rs!turno, "00"), "00")
     
      If rshoras.RecordCount > 0 Then rshoras.MoveFirst
      Do While Not rshoras.EOF
         MField = "h" & rshoras!Codigo
'         If rshoras!codigo <> "14" Then
            rshoras!Monto = Rs.Fields(MField)
'         Else
'            rshoras!Monto = rs.Fields(MField) / 8
'         End If
         rshoras.MoveNext
      Loop
      If rspagadic.RecordCount > 0 Then rspagadic.MoveFirst
      Do While Not rspagadic.EOF
         MField = "i" & rspagadic!Codigo
         rspagadic!Monto = Rs.Fields(MField)
         rspagadic.MoveNext
      Loop
      
      If rsdesadic.RecordCount > 0 Then rsdesadic.MoveFirst
         Do While Not rsdesadic.EOF
            If rsdesadic!Codigo = "05" And lRetJud = True Then 'Judicial se calcula
                rsdesadic!Monto = 0
            Else
               MField = "d" & rsdesadic!Codigo
               rsdesadic!Monto = Rs.Fields(MField)
            End If
            rsdesadic.MoveNext
         Loop
   End If
End If
'Frame2.Enabled = False
'Frame3.Enabled = False
'Frame4.Enabled = False
'Frame5.Enabled = False

'If VTipobol = "02" Then Dgrdhoras.Enabled = False
vDevengue = devengue
If devengue = True Then Calcula_Devengue_Vaca
'If VTipobol = "01" Then MUESTRA_CUENTACORRIENTE
If LblId.Caption <> "" Then BtnVerCalculo.Visible = True
End Sub
Private Sub Total_porcentaje(mdele As String)
Dim mcolf As Integer
Dim mtotalf As Currency
Dim rST As New Recordset


If rsccosto.RecordCount <= 0 Then lbltot.Caption = "0.00": Exit Sub
If Not rsccosto.EOF And mdele <> "S" Then mi = rsccosto!Item
Set rST = rsccosto.Clone
rST.AbsolutePosition = 1
With rST
    mtotalf = 0
    mtotali = 0
    If rST.RecordCount > 0 Then
       mcolf = Dgrdccosto.Row
       rST.MoveFirst
        
       Do While Not rST.EOF
          If Not IsNull(rST!Monto) Then
             mtotalf = mtotalf + rST!Monto
          End If

          rST.MoveNext
       Loop
    End If
    lbltot.Caption = Format(mtotalf, "#,###,###.00")
End With

If CCur(lbltot.Caption) > 100 Then
   MsgBox "Total Porcentaje no puede exceder a 100%", vbCritical, TitMsg
   rsccosto!Monto = Format(CCur(rsccosto!Monto) - (CCur(lbltot.Caption) - 100), "###,###.00")
   lbltot.Caption = "100.00"
End If
Dgrdccosto.Refresh
Dgrdccosto.MarqueeStyle = 6
If Dgrdccosto.Enabled = True Then Dgrdccosto.SetFocus
End Sub
Public Sub Grabar_Boleta()

If Cierre_Planilla(Year(VFProceso), Month(VFProceso)) Then Exit Sub

If PnlPreView.Visible = True Then Exit Sub

Dim MqueryH As String
Dim MqueryP As String
Dim MqueryD As String
Dim MqueryI As String
Dim MqueryCalD As String
Dim MqueryCalA As String
Dim mtoting As Currency
Dim itemcosto As Integer
Dim mcad As String
Dim Quincena As Currency
Dim EsNegativo As Double
Dim NroTrans As Integer
Dim DiasTrabV As Boolean
On Error GoTo ErrorTrans
NroTrans = 0
DiasTrabV = False
EsNegativo = False
If vDevengue = True Then
   Mstatus = "D"
   VTurno = Format(NumDev, "00")
Else
   Mstatus = "T"
End If

If Trim(LblBasico.Caption) = "" Then
    'If wGrupoPla = "01" Then 'En Gallos no se permite grabar si el trabajador no tiene sueldo Basico se modifico para las otras por disposicion de JCB (Destajo) 16/02/2012
    
    If bImportar = False Then
       MsgBox "Trabajador No Registra Sueldo Basico", vbInformation, "Boletas de Pago"
    Else
       LstObs.AddItem Trim(Txtcodpla.Text) & ": Trabajador No Registra Sueldo Basico"
    End If
    LblBasico.Caption = ""
    VokDevengue = False
    Exit Sub
    'Else
    '   Lblbasico.Caption = "0.00"
    'End If
End If

mcancel = False
mtoting = 0
Total_porcentaje ("S")

If VTipobol = "03" Then
   If Not Verifica_Meses_Grati Then Exit Sub
End If


If Trim(Cmbturno.Text) = "" Then
    If bImportar = False Then
       MsgBox "Debe Indicar Turno", vbCritical, TitMsg
    Else
       LstObs.AddItem Trim(Txtcodpla.Text) & ": Debe Indicar Turno"
    End If
    
    Cmbturno.SetFocus: VokDevengue = False
    Exit Sub
End If

If CCur(lbltot.Caption) <> 100 Then
    If bImportar = False Then
       MsgBox "Total Porcentaje de Centro de Costos debe ser 100%", vbCritical, TitMsg
    Else
       LstObs.AddItem Trim(vPlaCod) & ": Total Porcentaje de Centro de Costos debe ser 100%"
    End If
    
    VokDevengue = False
    Exit Sub
End If

'Verificar Activos a signados al trabajador
If VTipobol = "01" Then
   If Not Verifica_Activo Then Exit Sub
End If

'Activar para cuenta corriente
If Not Verifica_Cuenta_Corriente Then
    If bImportar = False Then
       MsgBox "Monto a Descontar por Cta.Cte no debe exceder el saldo", vbCritical, TitMsg
    Else
       LstObs.AddItem Trim(vPlaCod) & ": Monto a Descontar por Cta.Cte no debe exceder el saldo"
    End If
    Exit Sub
End If

'Fecha de Cese
If Trim(TxtFecCese.Text) <> "__/__/____" And cbo_TipMotFinPer.ReturnCodigo = -1 Then
    If bImportar = False Then
       MsgBox "Debe elegir el motivo fin de periodo", vbCritical, TitMsg
    Else
       LstObs.AddItem Trim(vPlaCod) & ": Debe elegir el motivo fin de periodo"""
    End If
    Exit Sub
End If

If VTipobol = "02" And BolDevengada = True And vDevengue = False Then
   Grabar_Devengada
   Exit Sub
End If

'para dejar pasar boleta temporalmente
If Trim(LblId.Caption & "") = "" Then
   If Verifica_Boleta(True) = False Then
      If wtipodoc = True And VTipobol <> "02" Then
         MsgBox "Ya existe boleta generada con el mismo periodo, Debe eliminarla para poder generar otra", vbExclamation, TitMsg: VokDevengue = False: Exit Sub
      ElseIf VTipobol <> "02" Then
         MsgBox "Ya existe Adelanto de Quincena con el mismo periodo, Debe eliminar para poder generar otra", vbExclamation, TitMsg: VokDevengue = False: Exit Sub
      End If
   End If
End If

If vDevengue <> True Then
    If Not wImportarXls Then
       If Trim(LblId.Caption & "") = "" Then
          Mgrab = MsgBox("Seguro de Grabar Boleta", vbYesNo + vbQuestion, TitMsg)
       Else
          Mgrab = MsgBox("Seguro de Modificar Boleta", vbYesNo + vbQuestion, TitMsg)
       End If
       If Mgrab <> 6 Then Exit Sub
    Else
        Mgrab = 6
    End If
End If

            
Screen.MousePointer = vbArrowHourglass

'If Not wImportarXls Then
   Sql$ = wInicioTrans & " GRABA_BOLETA"
   cn.Execute Sql$
'End If

NroTrans = 1
'Limpiamos Tabla temporal
Sql$ = "delete from platemphist where cia='" & wcia & "' and placod='" & Trim(Txtcodpla.Text) & "'"
Sql$ = Sql$ & " and cia='" & wcia & "' and placod='" & Txtcodpla.Text & "' and codauxinterno='" & Lblcodaux.Caption & "' and proceso='" & VTipobol & "' "
Sql$ = Sql$ & "and fechaproceso='" & Format(VFProceso, FormatFecha) & "' and semana='" & VSemana & "' and status='" & Mstatus & "'"
cn.Execute Sql


manos = perendat(VFProceso, VFechaNac, "a")
'VOLVER AQUI

If Trim(LblId.Caption & "") <> "" Then
   Sql$ = "update plahistorico set status='*',fec_modi=" & FechaSys & ",user_modi='" & wuser & "' where Id_Boleta=" & LblId.Caption & ""
   cn.Execute Sql$
   
   If VTipobol = "04" And lCalculaLiqui Then
      Sql$ = "Update plahisliquid set status='*' where Id_Boleta=" & LblId.Caption & ""
      cn.Execute Sql$
   End If
End If

'Calculo Liquidacion
If VTipobol = "04" And lCalculaLiqui Then

   If Mid(LblFingreso.Caption, 4, 7) <> Mid(TxtFecCese.Text, 4, 7) Then
      Dim lProvV As Boolean
      Dim lProvC As Boolean
      lProvV = False
      lProvC = False
   
      Sql$ = "usp_Verfica_Prov_anterior '" & wcia & "'," & Val(Mid(TxtFecCese.Text, 7, 4)) & "," & Val(Mid(TxtFecCese.Text, 4, 2)) & ",'" & Txtcodpla.Text & "'"
      If fAbrRst(RsLiquid, Sql) Then RsLiquid.MoveFirst
      Do While Not RsLiquid.EOF
         If RsLiquid!tipo = "V" Then lProvV = True
         If RsLiquid!tipo = "C" Then lProvC = True
         RsLiquid.MoveNext
      Loop
      RsLiquid.Close: Set RsLiquid = Nothing
      If Not lProvV Then
            'cn.RollbackTrans
            Sql$ = wCancelTrans & " GRABAR_BOLETA"
            cn.Execute Sql$

            MsgBox "No se encontro provisión de Vacaciones del periodo anterior", vbInformation
            
            Exit Sub
      End If
      If Not lProvC Then
            'cn.RollbackTrans
            Sql$ = wCancelTrans & " GRABAR_BOLETA"
            cn.Execute Sql$
            
            MsgBox "No se encontro provisión de CTS del periodo anterior", vbInformation
            Exit Sub
      End If
   End If
   
   If Trim(LblId.Caption) <> "" Then
      Sql$ = "update plahisliquid set status='*' where Id_Boleta=" & LblId.Caption & ""
      cn.Execute Sql$
   End If
   Sql = "usp_Liq_Fechas '" & wcia & "'," & Val(Mid(TxtFecCese.Text, 7, 4)) & "," & Val(Mid(TxtFecCese.Text, 4, 2)) & "," & Val(Mid(TxtFecCese.Text, 1, 2)) & ",'" & Txtcodpla.Text & "'"
   If fAbrRst(RsLiquid, Sql) Then RsLiquid.MoveFirst

   Sql$ = "delete from platmphisliquid"
   cn.Execute Sql
   Call Agrega_Promedios("02", "S")

   Sql$ = "update platmphisliquid set totaling=i01+i02+i03+i04+i05+i06+i07+i08+i09+i10+i11+i12+i13+i14+i15+i16+i17+i18+i19+i20+i21+i22+i23+i24+i25+i26+i27+i28+i29+i30+i31+i32+i33+i34+i35+i36+i37+i38+i39+i40+i41+i42+i43+i44+i45+i46+i47+i48+i49+i50"
   cn.Execute Sql

   lBasico = 0

   Dim RsBase As New ADODB.Recordset
   Sql$ = "select concepto,round((importe/factor_horas),10) as basico from plaremunbase "
   Sql$ = Sql$ & "where cia='" & Trim(wcia) & "' and status<>'*' and placod='" & Trim(Txtcodpla.Text) & "' "
   Sql$ = Sql$ & "and concepto in('01','02')"
   If fAbrRst(RsBase, Sql$) Then RsBase.MoveFirst
   Do While Not RsBase.EOF
      If RsBase!concepto = "01" Then
         Sql$ = "update platmphisliquid set sueldo=" & Round(RsBase!Basico * 8, 2) & ""
         cn.Execute Sql
      End If
      If RsBase!concepto = "02" Then
         Sql$ = "update platmphisliquid set asigfam=" & Round(RsBase!Basico * 8, 2) & ""
         cn.Execute Sql
      End If
      RsBase.MoveNext
   Loop
   RsBase.Close: Set RsBase = Nothing

   Sql$ = "update platmphisliquid set promedios=Round(totaling/30,2)"
   cn.Execute Sql
   
   Sql$ = "select top 1 totaling-i30 as Grati from plahistorico where cia='" & wcia & "' and proceso='03' and placod='" & Trim(Txtcodpla.Text) & "' and status<>'*' order by fechaproceso desc"
   If fAbrRst(RsBase, Sql$) Then
      Sql$ = "update platmphisliquid set promgrati= " & Round((RsBase!grati / 6) / 30, 2) & ""
      cn.Execute Sql
   End If
   RsBase.Close: Set RsBase = Nothing
   
   Sql$ = "update platmphisliquid set "
   Sql$ = Sql$ & "ivano=Round(((sueldo+asigfam+promedios) * 30) * vano,2),"
   Sql$ = Sql$ & "ivmes=Round((((sueldo+asigfam+promedios) * 30)/12) * vmes,2),"
   Sql$ = Sql$ & "ivdia=Round((((sueldo+asigfam+promedios) * 30)/360) * vdia,2),"
   Sql$ = Sql$ & "igmes=Round((((sueldo+asigfam+promedios) * 30)/6) * gmes,2),"
   Sql$ = Sql$ & "igdia=Round((((sueldo+asigfam+promedios) * 30)/180) * gdia,2),"
   Sql$ = Sql$ & "icano=Round(((sueldo+asigfam+promedios+promgrati) * 30) * cano,2),"
   Sql$ = Sql$ & "icmes=Round((((sueldo+asigfam+promedios+promgrati) * 30)/12) * cmes,2),"
   Sql$ = Sql$ & "icdia = Round((((sueldo + asigfam + promedios + promgrati) * 30) / 360) * cdia, 2)"
   cn.Execute Sql
End If


Sql$ = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
 'Sql$ = Sql$ & "insert into platemphist (cia,placod,codauxinterno,proceso,fechaproceso,semana,fechacese,fechainicial,fechafinal,fechavacai,fechavacaf,fecperiodovaca" & _
 '       ",fechaingreso,turno,h01,h02,h03,h04,h05,h06,h07,h08,h09,h10,h11,h12,h13,h14,h15,h16,h17,h18,h19,h20,h21,h22,h23,h24,h25,h26,h27,h28,h29,h30" & _
 '       ",i01,i02,i03,i04,i05,i06,i07,i08,i09,i10,i11,i12,i13,i14,i15,i16,i17,i18,i19,i20,i21,i22,i23,i24,i25,i26,i27,i28,i29,i30" & _
 '       ",i31,i32,i33,i34,i35,i36,i37,i38,i39,i40,i41,i42,i43,i44,i45,i46,i47,i48,i49,i50,d01,d02,d03,d04,d05,d06,d07,d08,d09,d10,d11,d12" & _
 '       ",d13,d14,d15,d16,d17,d18,d19,d20,d21,a01,a02,a03,a04,a05,a06,a07,a08,a09,a10,a11,a12,a13,a14,a15,a16,a17,a18,a19,a20,a21,saldoantcte" & _
 '       ",prestamo,totalapo,totalded,totaling,totneto,fec_crea,codafp,status,d111,d112,d113,d114,d115,tipotrab,obra,ccosto1,ccosto2,ccosto3,ccosto4,ccosto5" & _
 '       ",porc1,porc2,porc3,porc4,porc5,numafp,basico,fec_modi,user_modi,user_crea,tottardanza,Area) Values "
 Sql$ = Sql$ & "insert into platemphist (cia,placod," & _
        "codauxinterno,proceso,fechaproceso,fechavacai,fechavacaf,fecperiodovaca,semana," & _
        "fechaingreso,turno,codafp,status,fec_crea," & _
        "tipotrab,obra,numafp,basico,fec_modi," & _
        "user_crea,user_modi,Area,d21,a21) values('" & wcia & _
        "','" & Trim(Txtcodpla.Text) & "','" & _
        Lblcodaux.Caption & "','" & VTipobol & "','" & _
        Format(VFProceso, FormatFecha) & "'," & IIf(Txtvacai.Text <> "__/__/____", "'" & Format(Txtvacai.Text, FormatFecha) & "'", "NULL") & "," & IIf(Txtvacaf.Text <> "__/__/____", "'" & Format(Txtvacaf.Text, FormatFecha) & "'", "NULL") & "," & _
        IIf(TxtVaca.Text <> "__/__/____", "'" & Format(TxtVaca.Text, FormatFecha) & "'", "NULL") & ",'" & VSemana & "','" & Format(LblFingreso.Caption, FormatFecha) & _
         "','" & Format(VTurno, "0") & "','" & _
         Lblcodafp.Caption & "','" & Mstatus & "'," & _
         FechaSys & ",'" & vTipoTra & "','" & LblPlanta.Caption & _
         "','" & Lblnumafp.Caption & "'," & _
        CCur(LblBasico.Caption) & "," & FechaSys & _
        ",'" & wuser & "','" & wuser & "','" & Nueva_area & "',0,0)"
 
cn.Execute Sql

lSubsidio = "N"
Dim lHorSub As Double
Dim lHorOtr As Double
lHorSub = 0: lHorOtr = 0
w_error = 0
'Horas
If rshoras.RecordCount > 0 Then
   rshoras.MoveFirst
   MqueryH = ""

   Do While Not rshoras.EOF
      Select Case rshoras!Codigo
            Case "14"
               'If vTipoTra = "02" Then
               '   If IsNumeric(rshoras!Monto) And rshoras!Monto > 7 Then w_error = 1:   MsgBox "Se Cancela la grabacion", vbCritical, "Dias erradas": Exit Sub
               'End If
               'If vTipoTra = "01" Then
               '   If IsNumeric(rshoras!Monto) And rshoras!Monto > 30 Then w_error = 1:   MsgBox "Se Cancela la grabacion", vbCritical, "Dias erradas": Exit Sub
               'End If
               DiasTrabV = True
            Case "01"
               ' If vTipoTra = "01" Then
               '     If IsNumeric(rshoras!Monto) And rshoras!Monto > 240 Then w_error = 1:   MsgBox "Se Cancela la grabacion", vbCritical, "horas erradas": Exit Sub
               ' End If
               ' If vTipoTra = "02" Then
               '     If IsNumeric(rshoras!Monto) And rshoras!Monto > 56 Then w_error = 1:   MsgBox "Se Cancela la grabacion", vbCritical, "horas erradas": Exit Sub
               ' End If
                DiasTrabV = True
             Case "06", "27", "30"
                  If IsNumeric(rshoras!Monto) Then lHorSub = lHorSub + rshoras!Monto
             Case "08", "25"
                  lHorSub = lHorSub + 0
             Case "14"
                  DiasTrabV = True
             Case Else
                  If IsNumeric(rshoras!Monto) Then lHorOtr = lHorOtr + rshoras!Monto
      End Select
      If rshoras!Codigo <> "14" Then
         MqueryH = MqueryH & "h" & rshoras!Codigo & "=" & IIf(IsNull(rshoras!Monto), 0, rshoras!Monto) & ""
       Else
'          If wGrupoPla = "01" Then
             MqueryH = MqueryH & "h" & rshoras!Codigo & "=" & IIf(IsNull(rshoras!Monto), 0, rshoras!Monto) & ""
'          Else
'             MqueryH = MqueryH & "h" & rshoras!codigo & "=" & IIf(IsNull(rshoras!Monto), 0, rshoras!Monto) * 8 & ""
'          End If
       End If
       rshoras.MoveNext
       If Not rshoras.EOF Then MqueryH = MqueryH & ","
   Loop
End If
'If wGrupoPla = "01" Then 'Calculos de aportacion Sin Asignacion Familiar cuando es subsidio para grupo Gallos
   If lHorSub > 0 And lHorOtr <= 0 Then lSubsidio = "S"
'End If
'Pagos Adicionales
If rspagadic.RecordCount > 0 Then
   rspagadic.MoveFirst
   MqueryP = ""
   Do While Not rspagadic.EOF
   
'      If Val(rspagadic!Monto) <> 0 Then Stop
      
      MqueryP = MqueryP & "i" & rspagadic!Codigo & "=" & Val(rspagadic!Monto) & ""
      
      
      rspagadic.MoveNext
      If Not rspagadic.EOF Then MqueryP = MqueryP & ","
   Loop
End If

'Descuentos Adicionales

If rsdesadic.RecordCount > 0 Then
   rsdesadic.MoveFirst
   MqueryD = ""
   Do While Not rsdesadic.EOF
    If rsdesadic!Codigo = "13" And wtipodoc = False And wGrupoPla <> "01" Then  'Grupo GALLOS Calcula quincena como boleta completa
      MqueryD = MqueryD & "d" & rsdesadic!Codigo & "=" & Round(Val(rsdesadic!Monto) / 2, 2) & ""
    Else
      MqueryD = MqueryD & "d" & rsdesadic!Codigo & "=" & Val(rsdesadic!Monto) & " "
    End If
      rsdesadic.MoveNext
      If Not rsdesadic.EOF Then MqueryD = MqueryD & ","
   Loop
End If

itemcosto = 1
mcad = ""
 
If rsccosto.RecordCount > 0 Then rsccosto.MoveFirst
Do While Not rsccosto.EOF
    If Not IsNull(rsccosto!Monto) Then
      If IsNumeric(rsccosto!Monto) Then
         If CCur(rsccosto!Monto) <> 0 Then
            mcad = mcad & "ccosto" & Format(itemcosto, "0") & " = '" & Trim(rsccosto!Codigo & "") & "'," & "porc" & Format(itemcosto, "0") & " = " & Str(rsccosto!Monto) & ","
         End If
      End If
    Else
        Exit Do
    End If
   rsccosto.MoveNext
   itemcosto = itemcosto + 1
Loop

If Trim(mcad & "") <> "" Then
   mcad = Mid(mcad, 1, Len(Trim(mcad)) - 1)

  Sql$ = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
  Sql$ = Sql$ & "Update platemphist set " & mcad & "," & IIf(Trim(MqueryH) = "", "", MqueryH & ",") & MqueryP & "," & MqueryD
  Sql$ = Sql$ & " where cia='" & wcia & "' and placod='" & _
     Trim(Txtcodpla.Text) & "' and codauxinterno='" & _
     Trim(Lblcodaux.Caption) & "' and proceso='" & _
     Trim(VTipobol) & "' and fechaproceso='" & _
     Format(VFProceso, FormatFecha) & _
     "' and semana='" & VSemana & "' and status='" & _
     Mstatus & "'"
     
  cn.Execute Sql
End If
mcad = ""

'Calculo de ingresos

MqueryI = ""
TotalIngresos = 0
For I = 1 To 50
   
    Select Case VTipobol
        Case Is = "01" 'Normal
             Sql$ = Me.F01(Format(I, "00"), Val(VSemana))
        Case Is = "02" 'Vacaciones
             Sql$ = Me.V01(Format(I, "00"))
        Case Is = "03" 'Gratificaciones
             Sql$ = Me.G01(Format(I, "00"))
        Case Is = "04"
            Sql$ = Me.F01(Format(I, "00"), Val(VSemana))
        Case Is = "05"
            Sql$ = Me.F01(Format(I, "00"), Val(VSemana))
        Case Is = "10"
            Sql$ = Me.F01(Format(I, "00"), Val(VSemana))
        Case Is = "11"
            Sql$ = Me.F01(Format(I, "00"), Val(VSemana))
    End Select
    
    If Trim(Sql$) <> "" Then
       'mgirao 09/11/2016 Temporal para que Mod Importe Movilidades
       Dim mtc As Double
       mtc = 0
       '---------
       cn.CursorLocation = adUseClient
       Set Rs = New ADODB.Recordset
       Set Rs = cn.Execute(Sql$, 64)
      'RS.Save "C:\ALEXTE.RS"

       If Rs.RecordCount > 0 Then
          Rs.MoveFirst
          If IsNull(Rs(0)) Or Rs(0) = 0 Then
          
          Else
         'mgirao 09/11/2016 Temporal para que Mod Importe Movilidades y Subsidios
         'If I = 3 And (UCase(wuser) = "HENRY" Or UCase(wuser) = "SA") Then
         'mtc = 0
         'mtc = Val(InputBox("Ingrese Monto Movilidad", "Movilidad Reproceso"))
         '   If Not IsNumeric(mtc) Then
         '        mtc = 0
         '   End If
         ' End If
         '---------
         
            If wtipodoc = False And (I = 10 Or I = 11 Or I = 13 Or I = 21 Or I = 22 Or I = 23 Or I = 24) Then
                MqueryI = MqueryI & "i" & Format(I, "00") & " = " & Round(Rs(0) / 2, 2) & ","
            Else
            

                   MqueryI = MqueryI & "i" & Format(I, "00") & " = " & Rs(0) & ","
                    'If I = 1 Or I = 2 Or I = 3 Or I = 4 Or I = 5 Or I = 7 Or I = 8 Or I = 9 Or I = 10 Or I = 11 Or I = 12 Or I = 13 Or I = 14 Or I = 15 Or I = 16 Or I = 17 Or I = 18 Or I = 19 Or I = 21 Or I = 22 Or I = 23 Or I = 24 Or I = 25 Or I = 26 Or I = 27 Or I = 28 Or I = 29 Or I = 30 Or I = 32 Or I = 35 Or I = 36 Or I = 37 Or I = 38 Or I = 39 Or I = 40 Or I = 44 Or I = 46 Or I = 48 Or I = 49 Or I = 34 Then
                    'If I <> 9 Then
                        TotalIngresos = TotalIngresos + Rs(0)
                    'End If

            End If
          End If
       End If
       If Rs.State = 1 Then Rs.Close
    End If
Next

If MqueryI <> "" Then
   MqueryI = Mid(MqueryI, 1, Len(Trim(MqueryI)) - 1)
   Sql$ = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
   Sql$ = Sql$ & " Update platemphist set " & MqueryI
   Sql$ = Sql$ & " where cia='" & wcia & "' and placod='" & Txtcodpla.Text & "' and codauxinterno='" & Lblcodaux.Caption & "' and proceso='" & VTipobol & "' "
   Sql$ = Sql$ & "and fechaproceso='" & Format(VFProceso, FormatFecha) & "' and semana='" & VSemana & "' and status='" & Mstatus & "'"
   cn.Execute Sql
End If

''08/07/2015
''actualizar "OTROS PAGOS" CUANDO ES GRATIFICACION
If VTipobol = "03" Then
    Otros_Pagos_Vac (True)
End If


If wGrupoPla = "01" And (VTipobol = "02" Or VTipobol = "03") Then 'Calculos de promedios es detallado en Gallos
   Call Agrega_Promedios(VTipobol, "N")
End If

'Si es subsidio Asignacion Familiar se pasa al Subsidio
'MARIO GIRAO SE SUPRIME LA ASIGNACION FAMILIAR AL SUBSIDIO
Dim Permite As Boolean
Permite = True
If (WValor03 + WValor32 + wValor43 + WValor42) > 0 Then
  Permite = False
End If

If WValor32 > 0 Or wValor43 > 0 Then
     Sql$ = "UPDATE platemphist set i32= " & WValor32 + wValor43 & ",i43=0 where cia='" & wcia & "' and "
     Sql$ = Sql$ & " placod='" & Trim(Txtcodpla.Text) & "' and "
     Sql$ = Sql$ & " codauxinterno='" & Trim(Lblcodaux.Caption) & "' and proceso='" & VTipobol & "' "
     Sql$ = Sql$ & " and fechaproceso='" & Format(VFProceso, FormatFecha) & "' and semana='" & VSemana & "' and status='" & Mstatus & "'"
     cn.Execute Sql
Else
    If wValor43 > 0 And WValor32 = 0 Then
        Sql$ = "UPDATE platemphist set i32= i32 +" & wValor43 & ",i43=0 where cia='" & wcia & "' and "
        Sql$ = Sql$ & " placod='" & Trim(Txtcodpla.Text) & "' and "
        Sql$ = Sql$ & " codauxinterno='" & Trim(Lblcodaux.Caption) & "' and proceso='" & VTipobol & "' "
        Sql$ = Sql$ & " and fechaproceso='" & Format(VFProceso, FormatFecha) & "' and semana='" & VSemana & "' and status='" & Mstatus & "'"
        cn.Execute Sql
    Else
        If wValor43 = 0 And WValor32 > 0 Then
            Sql$ = "UPDATE platemphist set i32= i32 +" & WValor32 & " where cia='" & wcia & "' and "
            Sql$ = Sql$ & " placod='" & Trim(Txtcodpla.Text) & "' and "
            Sql$ = Sql$ & " codauxinterno='" & Trim(Lblcodaux.Caption) & "' and proceso='" & VTipobol & "' "
            Sql$ = Sql$ & " and fechaproceso='" & Format(VFProceso, FormatFecha) & "' and semana='" & VSemana & "' and status='" & Mstatus & "'"
            cn.Execute Sql
        End If
    End If
End If

If Verifica_Subsidio And Permite Then
   Sql$ = "UPDATE platemphist set i42=(Case h27 when 0 then i42 else i42+i02 end),i43=(Case h27 when 0 then i43+i02 else i43 end),i02=0 where cia='" & wcia & "' and "
   'Sql$ = "UPDATE platemphist set i42=(Case h27 when 0 then i42 else i42+i02 end),i43=3526.80,i02=0 where cia='" & wcia & "' and "
   Sql$ = Sql$ & " placod='" & Trim(Txtcodpla.Text) & "' and "
   Sql$ = Sql$ & " codauxinterno='" & Trim(Lblcodaux.Caption) & "' and proceso='" & VTipobol & "' "
   Sql$ = Sql$ & " and fechaproceso='" & Format(VFProceso, FormatFecha) & "' and semana='" & VSemana & "' and status='" & Mstatus & "'"
   cn.Execute Sql
End If

'Calculo de Deducciones
MqueryCalD = ""
'mgirao cambio para solo quinta
'For I = 13 To 13
For I = 1 To 21

    'MGIRAO Registro Boleta Practicantes no ingresa a deduccion
    If vTipoTra = "05" Then Exit For
    
    If I = 1 Then
    A = A
    End If
    If I = 11 Then
    A = A
    End If
    
    Sql$ = Me.F02(Format(I, "00"), False)
    
    If mcancel = True Then
       VokDevengue = False
       If NroTrans = 1 Then
            'cn.RollbackTrans
            Sql$ = wCancelTrans & " GRABA_BOLETA"
            cn.Execute Sql$
       End If
       MsgBox "Se Cancelo la Grabacion", vbCritical, "Calculo de Boleta"
       Screen.MousePointer = vbDefault
       Exit Sub
    End If
    If Sql$ <> "" Then
       If (fAbrRst(Rs, Sql$)) Then
          Rs.MoveFirst
          If IsNull(Rs(0)) Or Rs(0) = 0 Then
            MqueryCalD = MqueryCalD & "d" & Format(I, "00") & " = " & Rs(0) & ","
          Else
             If I = 11 Then
                For J = 1 To 5
                   MqueryCalD = MqueryCalD & "d" & Format(I, "00") & Format(J, "0") & " = " & Rs(J - 1) & ","
                Next J
             'ElseIf I = 13 And wtipodoc = False Then
             Else
             'mgirao para evitar quinta en negativo
             If I = 13 And Rs(0) <= 0 Then
                MqueryCalD = MqueryCalD & "d" & Format(I, "00") & " = " & 0 & ","
             End If
             '---------
             If I = 13 And (wtipodoc = False And wGrupoPla <> "01") Then 'Grupo Gallos calcula quincena como boleta normal 240 horas
                If Rs(0) < 0 Then
                    MqueryCalD = MqueryCalD & "d" & Format(I, "00") & " = " & 0 & ","
                Else
                    MqueryCalD = MqueryCalD & "d" & Format(I, "00") & " = " & Rs(0) / 2 & ","
                End If
             ElseIf I = 6 Then
                '2011.03.09 PARA QUE NO CALCULE NUEVAMENTE EL ESSALUD VIDA SI YA LO CALCULO EN ALGUN TIPO DE BOLETA
                Dim Cadena As String
                Dim rsVal As ADODB.Recordset
                Cadena = "select d06 as Contador from plahistorico where cia = '" & wcia & "' and year(fechaproceso) = " & Year(VFProceso) & " and month(fechaproceso) = " & Month(VFProceso) & " and placod = '" & Trim(Txtcodpla.Text) & "' and d06<>0 and status <> '*'"
                Set rsVal = OpenRecordset(Cadena, cn)
                If Not rsVal.EOF Then
                    If Not rsVal.EOF Then
                        If rsVal!contador > 0 Then
                            MqueryCalD = MqueryCalD & "d" & Format(I, "00") & " = " & 0 & ","
                        Else
                            MqueryCalD = MqueryCalD & "d" & Format(I, "00") & " = " & Rs(0) & ","
                        End If
                    End If
                Else
                    'EN TEORIA NO DEBERIA SALTAR A ESTE LADO PERO EN CASO DE HACERLO SE VERA REFLEJADO EN LA BOLETA
                    'DE SER ASI MODIFICAR EL QUERY DE LA VARIABLE 'CADENA'
                    MqueryCalD = MqueryCalD & "d" & Format(I, "00") & " = " & Rs(0) & ","
                End If
             Else
                If Rs(0) < 0 And I <> 4 Then
                    'mgirao evitar que bloquee los negativos
                '    If I = 13 Then
                '       MqueryCalD = MqueryCalD & "d" & Format(I, "00") & " = " & rs(0) & ","
                '    Else
                '    MqueryCalD = MqueryCalD & "d" & Format(I, "00") & " = " & 0 & ","
                '    End If
                Else
                    MqueryCalD = MqueryCalD & "d" & Format(I, "00") & " = " & Rs(0) & ","
                End If
             End If
          End If
       End If
       End If
       If Rs.State = 1 Then Rs.Close
    End If
Next I
If MqueryCalD <> "" Then
   MqueryCalD = Mid(MqueryCalD, 1, Len(Trim(MqueryCalD)) - 1)
   Sql$ = " SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
   Sql$ = Sql$ & "Update platemphist set " & MqueryCalD
   Sql$ = Sql$ & " where cia='" & wcia & "' and placod='" & Txtcodpla.Text & "' and codauxinterno='" & Lblcodaux.Caption & "' and proceso='" & VTipobol & "' "
   Sql$ = Sql$ & "and fechaproceso='" & Format(VFProceso, FormatFecha) & "' and semana='" & VSemana & "' and status='" & Mstatus & "'"
   cn.Execute Sql
   
   Sql$ = " SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
   'Sql$ = Sql$ & "Update platemphist set d11=d111+d112+d113+d114+d115"
   'Sql$ = Sql$ & "Update platemphist set d11=d111+d112+d113+d114+d115,d112=(case when d112=0 then case when d111>0 then d111 else 0 end else d112 end),d111=(case when d112=0 then case when d111>0 then 0 else 0 end else d111 end)"
   Sql$ = Sql$ & "Update platemphist set d11=d111+d112+d113+d114+d115"
   
   Sql$ = Sql$ & " where cia='" & wcia & "' and placod='" & Txtcodpla.Text & "' and codauxinterno='" & Lblcodaux.Caption & "' and proceso='" & VTipobol & "' "
   Sql$ = Sql$ & "and fechaproceso='" & Format(VFProceso, FormatFecha) & "' and semana='" & VSemana & "' and status='" & Mstatus & "'"
   cn.Execute Sql
   
End If

'Calculo de Aportaciones
MqueryCalA = ""
'mgirao cambio para solo quinta
'For I = 20 To 20
For I = 1 To 21
    If I = 1 Then
    A = A
    End If
      'MGIRAO Registro Boleta Practicantes no ingresa a deduccion
    If vTipoTra = "05" Then Exit For
    
    Sql$ = F03(Format(I, "00"), VTipobol)
    If Sql$ <> "" Then
       cn.CursorLocation = adUseClient
       Set Rs = New ADODB.Recordset
       Set Rs = cn.Execute(Sql$, 64)
       If Rs.RecordCount > 0 Then
          Rs.MoveFirst
          If IsNull(Rs(0)) Or Rs(0) = 0 Then
          Else
             MqueryCalA = MqueryCalA & "a" & Format(I, "00") & " = " & Rs(0) & ","
          End If
       End If
       If Rs.State = 1 Then Rs.Close
    End If
Next
If MqueryCalA <> "" Then
   MqueryCalA = Mid(MqueryCalA, 1, Len(Trim(MqueryCalA)) - 1)
   Sql$ = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
   Sql$ = Sql$ & "Update platemphist set " & MqueryCalA
   Sql$ = Sql$ & " where cia='" & wcia & "' and placod='" & Txtcodpla.Text & "' and codauxinterno='" & Lblcodaux.Caption & "' and proceso='" & VTipobol & "' "
   Sql$ = Sql$ & "and fechaproceso='" & Format(VFProceso, FormatFecha) & "' and semana='" & VSemana & "' and status='" & Mstatus & "'"
   cn.Execute Sql
End If

'****************************************************
'* FECHA          : 12/06/2009
'* AUTOR          : JUAN JIMENEZ
'* Tipo de Boleta : GRATIFICACION
'* BONIFICACION EXTRAORDINARIA = ESSALUD
'* SNP                         = 0
'* AFP                         = 0
'* ESSALUD                     = 0
'****************************************************

'modificacion 12-01-11, se comenta esta parte por cambio de ley
'para que realice la bonificacion eliminar el "And 1 = 2"

'2011.06.20
'POR MODIFICACION DE LEY SE QUITA LA CONDICION "AND 1 = 2" Y SE CONSIDERA EL IMPORTE POR EPS(CALCULO YA EFECTUADO EN LA FUNCION F03)
'EL VALOR DE LA COLUMNA "a01" YA VIENE CALCULADO POR EL % CORRESPONDIENTE A LA EPS ... VER FUNCION F03
If VTipobol = "03" Then 'And 1 = 2 Then
    d_resultado = 0
    If d_resultado > 0 Then
        Sql$ = "UPDATE platemphist set i30= " & d_resultado & " where cia='" & wcia & "' and "
        Sql$ = Sql$ & " placod='" & Trim(Txtcodpla.Text) & "' and "
        Sql$ = Sql$ & " codauxinterno='" & Trim(Lblcodaux.Caption) & "' and proceso='" & VTipobol & "' "
        Sql$ = Sql$ & " and fechaproceso='" & Format(VFProceso, FormatFecha) & "' and semana='" & VSemana & "' and status='" & Mstatus & "'"
        cn.Execute Sql
    Else
        '* BONIFICACION EXTRAORDINARIA = ESSALUD
        Sql$ = "UPDATE platemphist set i30= a01 where cia='" & wcia & "' and "
        Sql$ = Sql$ & " placod='" & Trim(Txtcodpla.Text) & "' and "
        Sql$ = Sql$ & " codauxinterno='" & Trim(Lblcodaux.Caption) & "' and proceso='" & VTipobol & "' "
        Sql$ = Sql$ & " and fechaproceso='" & Format(VFProceso, FormatFecha) & "' and semana='" & VSemana & "' and status='" & Mstatus & "'"
        cn.Execute Sql
    End If
    '* SNP                         = 0
    '* AFP                         = 0
    '* ESSALUD                     = 0
    '* EPS                         = N
    Sql$ = "UPDATE platemphist set d14=0,d11=0, d04=0, a01=0, d111=0, d112=0, d113=0, d114=0, d115=0 where cia='" & wcia & "' and "
    Sql$ = Sql$ & " placod='" & Trim(Txtcodpla.Text) & "' and "
    Sql$ = Sql$ & " codauxinterno='" & Trim(Lblcodaux.Caption) & "' and proceso='" & VTipobol & "' "
    Sql$ = Sql$ & " and fechaproceso='" & Format(VFProceso, FormatFecha) & "' and semana='" & VSemana & "' and status='" & Mstatus & "'"
    cn.Execute Sql
    
    'Calculo de Qta Por Pago Extraordinario para boleta de gratificacion
    'ya no va desde el 2016
    
'    Dim mMontoQtaGrati As Currency
'    mMontoQtaGrati = 0
'    Sql$ = F02("13", True)
'    If Trim(Sql$ & "") <> "" Then
'       If (fAbrRst(rs, Sql$)) Then
'          If Trim(rs(0) & "") <> "" Then mMontoQtaGrati = rs(0)
'       End If
'       rs.Close
'       If mMontoQtaGrati <> 0 Then
'          Sql$ = "UPDATE platemphist set d13= " & mMontoQtaGrati & " where cia='" & wcia & "' and "
'          Sql$ = Sql$ & " placod='" & Trim(Txtcodpla.Text) & "' and "
'          Sql$ = Sql$ & " codauxinterno='" & Trim(Lblcodaux.Caption) & "' and proceso='" & VTipobol & "' "
'          Sql$ = Sql$ & " and fechaproceso='" & Format(VFProceso, FormatFecha) & "' and semana='" & VSemana & "' and status='" & Mstatus & "'"
'          cn.Execute Sql
'       End If
'    End If
    
    'fin ya no va
 
End If

'Si quinta es menos de 2 soles no descontar hasta el mes de setimbre y solo para obreros
If vTipoTra <> "01" And Vmes < 10 Then
   Sql$ = "UPDATE platemphist set d13=0 where cia='" & wcia & "' and "
   Sql$ = Sql$ & " placod='" & Trim(Txtcodpla.Text) & "' and "
   Sql$ = Sql$ & " codauxinterno='" & Trim(Lblcodaux.Caption) & "' and proceso='" & VTipobol & "' "
   Sql$ = Sql$ & " and fechaproceso='" & Format(VFProceso, FormatFecha) & "' and semana='" & VSemana & "' and status='" & Mstatus & "' and d13<2"
   cn.Execute Sql
End If

Dim mi As String, md As String, ma As String
mi = "i01+i02+i03+i04+i05+i06+i07+i08+i09+i10+i11+i12+i13+i14+i15+i16+i17+i18+i19+i20+i21+i22+i23+i24+i25+i26+i27+i28+i29+i30+i31+i32+i33+i34+i35+i36+i37+i38+i39+i40+i41+i42+i43+i44+i45+i46+i47+i48+i49+i50"
md = "d01+d02+d03+d04+d05+d06+d07+d08+d09+d10+d11+d12+d13+d14+d15+d16+d17+d18+d19+d20+isnull(d21,0)"
ma = "a01+a02+a03+a04+a05+a06+a07+a08+a09+a10+a11+a12+a13+a14+a15+a16+a17+a18+a19+a20+isnull(a21,0)"
Sql$ = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
If wGrupoPla = "01" Then
   Sql$ = Sql$ & "update platemphist set totaling=" & mi & ","
Else
   If DiasTrabV = True Then
       Sql$ = Sql$ & "update platemphist set totaling=" & mi & ","
   Else
       Sql$ = Sql$ & "update platemphist set h14=round(h01+h02/8,0),totaling=" & mi & ","
   End If
End If


Sql$ = Sql$ & "totalded=" & md & "," _
     & "totalapo=" & ma & "," _
     & "totneto=(" & mi & ")-" & "(" & md & ")"
   Sql$ = Sql$ & " where cia='" & wcia & "' and " & _
   " placod='" & Trim(Txtcodpla.Text) & "' and " & _
   "codauxinterno='" & Trim(Lblcodaux.Caption) & "' and proceso='" & VTipobol & "' "
   Sql$ = Sql$ & "and fechaproceso='" & Format(VFProceso, FormatFecha) & "' and semana='" & VSemana & "' and status='" & Mstatus & "'"
cn.Execute Sql

'Sql$ = "UPDATE platemphist set h14=round((COALESCE(h01,0)+COALESCE(h02,0)+COALESCE(h03,0))/8,0) where cia='" & wcia & "' and "
If wGrupoPla <> "01" And DiasTrabV = False Then 'Grupo Gallos Indica Dias Trabajados
   Sql$ = "UPDATE platemphist set h14=round((COALESCE(h01,0)+COALESCE(h02,0))/8,0) where cia='" & wcia & "' and "
   Sql$ = Sql$ & " placod='" & Trim(Txtcodpla.Text) & "' and "
   Sql$ = Sql$ & " codauxinterno='" & Trim(Lblcodaux.Caption) & "' and proceso='" & VTipobol & "' "
   Sql$ = Sql$ & " and fechaproceso='" & Format(VFProceso, FormatFecha) & "' and semana='" & VSemana & "' and status='" & Mstatus & "'"
   cn.Execute Sql
End If

'Calculo se Retencion Judicial se quita Sindicato y essakud vida  d08 y d20, d06
If VTipobol = "11" Then
   Sql$ = "select importe,status from pladeducper where cia='" & wcia & "' and placod='" & Trim(Txtcodpla.Text) & "' and concepto='05' and status<>'*' and placod in(select placod from PlaRetJudUti where cia='" & wcia & "' and status<>'*')"
Else
   'mgirao cambio para solo quinta
   'Sql$ = "select importe,status from pladeducper where cia='" & wcia & "' and placod='" & Trim(Txtcodpla.Text) & "' and concepto='05' and status<>'*'"
   'add jcms 180624
   Sql$ = "select importe,status "
   Sql$ = Sql$ & " ,isnull((select distinct IndicadorCalculoEnBaseIngresoTotalBruto from TBL_BCO_CUENTA_DJ where status<>'*' AND PLACOD=pladeducper.placod),0) as 'IndicadorCalculoEnBaseIngresoTotalBruto'"
   Sql$ = Sql$ & " from pladeducper where cia='" & wcia & "' and placod='" & Trim(Txtcodpla.Text) & "' and concepto='05' and status<>'*'"
   
End If

If (fAbrRst(Rs, Sql$)) Then
   If Trim(Rs!status & "") = "P" Then
      'Sql$ = "UPDATE platemphist set d05=round((totneto+d09+d07-i35)* " & rs!importe & " /100,2),totneto=totneto-round((totneto+d09+d07-i35)* " & rs!importe & " /100,2),totalded=totalded+round((totneto+d09+d07-i35)* " & rs!importe & " /100,2) where cia='" & wcia & "' and "
      'Sql$ = Sql$ & " placod='" & Trim(Txtcodpla.Text) & "' and "
      'Sql$ = Sql$ & " codauxinterno='" & Trim(Lblcodaux.Caption) & "' and proceso='" & VTipobol & "' "
      'Sql$ = Sql$ & " and fechaproceso='" & Format(VFProceso, FormatFecha) & "' and semana='" & VSemana & "' and status='" & Mstatus & "'"
      'cn.Execute Sql
      'modificado mgirao a solicitado jcontreras kporras 28/12/2017
      'If Trim(Txtcodpla.Text) = "O6604" Then
'        Sql$ = "UPDATE platemphist set d05=round(((totneto - I47 ) + totalded - d13 - d11 - d04 )* " & rs!importe & " /100,2),totneto=totneto-round((totneto + totalded - d13 - d11 - d04)* " & rs!importe & " /100,2),totalded=totalded+round((totneto + totalded - d13 - d11 - d04 )* " & rs!importe & " /100  + 2.99 ,2) where cia='" & wcia & "' and "

       ' Sql$ = "UPDATE platemphist set d05=1677.77,totneto=totneto-round(((totneto - 1677.77)+ totalded - d13 - d11 - d04)* " & rs!importe & " /100,2),totalded=totalded+round(((totneto - 1677.77 )+ totalded - d13 - d11 - d04 )* " & rs!importe & " /100 ,2) where cia='" & wcia & "' and "
       ' Sql$ = Sql$ & " placod='" & Trim(Txtcodpla.Text) & "' and "
       ' Sql$ = Sql$ & " codauxinterno='" & Trim(Lblcodaux.Caption) & "' and proceso='" & VTipobol & "' "
       ' Sql$ = Sql$ & " and fechaproceso='" & Format(VFProceso, FormatFecha) & "' and semana='" & VSemana & "' and status='" & Mstatus & "'"
      'Else
        Sql$ = "UPDATE platemphist set d05=round((totneto + totalded - d13 - d11 - d04 )* " & Rs!importe & " /100,2),totneto=totneto-round((totneto + totalded - d13 - d11 - d04)* " & Rs!importe & " /100,2),totalded=totalded+round((totneto + totalded - d13 - d11 - d04 )* " & Rs!importe & " /100,2) where cia='" & wcia & "' and "
        Sql$ = Sql$ & " placod='" & Trim(Txtcodpla.Text) & "' and "
        Sql$ = Sql$ & " codauxinterno='" & Trim(Lblcodaux.Caption) & "' and proceso='" & VTipobol & "' "
        Sql$ = Sql$ & " and fechaproceso='" & Format(VFProceso, FormatFecha) & "' and semana='" & VSemana & "' and status='" & Mstatus & "'"
     ' End If
     
     If Rs!IndicadorCalculoEnBaseIngresoTotalBruto Then 'si el calculo es en base al impo. total de ingreso bruto - 180624
        Sql$ = "UPDATE platemphist set d05=round((totneto + totalded)* " & Rs!importe & " /100,2),totneto=totneto-round((totneto + totalded)* " & Rs!importe & " /100,2),totalded=totalded+round((totneto + totalded )* " & Rs!importe & " /100,2) where cia='" & wcia & "' and "
        Sql$ = Sql$ & " placod='" & Trim(Txtcodpla.Text) & "' and "
        Sql$ = Sql$ & " codauxinterno='" & Trim(Lblcodaux.Caption) & "' and proceso='" & VTipobol & "' "
        Sql$ = Sql$ & " and fechaproceso='" & Format(VFProceso, FormatFecha) & "' and semana='" & VSemana & "' and status='" & Mstatus & "'"
     Else
        ' Cambio solicitado sara Roda 23/12/24 segun correo que no se afecte a este trabajador la condicion de trabajo
        If Trim(Txtcodpla.Text) = "E0656" Then
            Sql$ = "UPDATE platemphist set d05=round((totneto - i20 + totalded - d13 - d11 - d04 )* " & Rs!importe & " /100,2),totneto=totneto-round((totneto + totalded - d13 - d11 - d04)* " & Rs!importe & " /100,2),totalded=totalded+round((totneto + totalded - d13 - d11 - d04 )* " & Rs!importe & " /100,2) where cia='" & wcia & "' and "
            Sql$ = Sql$ & " placod='" & Trim(Txtcodpla.Text) & "' and "
            Sql$ = Sql$ & " codauxinterno='" & Trim(Lblcodaux.Caption) & "' and proceso='" & VTipobol & "' "
            Sql$ = Sql$ & " and fechaproceso='" & Format(VFProceso, FormatFecha) & "' and semana='" & VSemana & "' and status='" & Mstatus & "'"
        Else
            Sql$ = "UPDATE platemphist set d05=round((totneto + totalded - d13 - d11 - d04 )* " & Rs!importe & " /100,2),totneto=totneto-round((totneto + totalded - d13 - d11 - d04)* " & Rs!importe & " /100,2),totalded=totalded+round((totneto + totalded - d13 - d11 - d04 )* " & Rs!importe & " /100,2) where cia='" & wcia & "' and "
            Sql$ = Sql$ & " placod='" & Trim(Txtcodpla.Text) & "' and "
            Sql$ = Sql$ & " codauxinterno='" & Trim(Lblcodaux.Caption) & "' and proceso='" & VTipobol & "' "
            Sql$ = Sql$ & " and fechaproceso='" & Format(VFProceso, FormatFecha) & "' and semana='" & VSemana & "' and status='" & Mstatus & "'"
        End If
     End If
     
      cn.Execute Sql
   Else
      Sql$ = "UPDATE platemphist set d05= " & Rs!importe & ",totneto=totneto-" & Rs!importe & ",totalded=totalded+" & Rs!importe & " where cia='" & wcia & "' and "
      Sql$ = Sql$ & " placod='" & Trim(Txtcodpla.Text) & "' and "
      Sql$ = Sql$ & " codauxinterno='" & Trim(Lblcodaux.Caption) & "' and proceso='" & VTipobol & "' "
      Sql$ = Sql$ & " and fechaproceso='" & Format(VFProceso, FormatFecha) & "' and semana='" & VSemana & "' and status='" & Mstatus & "'"
      cn.Execute Sql
   End If
End If
Rs.Close: Set Rs = Nothing
'Fin Retencion Judicial

If wtipodoc = False And wGrupoPla = "01" Then 'En Grupo Gallos se calcula quincena sobre 240 horas, aca dividimos el neto /2
   Sql$ = "UPDATE platemphist set totneto=totneto/2 where cia='" & wcia & "' and "
   Sql$ = Sql$ & " placod='" & Trim(Txtcodpla.Text) & "' and "
   Sql$ = Sql$ & " codauxinterno='" & Trim(Lblcodaux.Caption) & "' and proceso='" & VTipobol & "' "
   Sql$ = Sql$ & " and fechaproceso='" & Format(VFProceso, FormatFecha) & "' and semana='" & VSemana & "' and status='" & Mstatus & "'"
   cn.Execute Sql
End If

'Revisar si EsNegativo
Sql$ = "Select totneto from platemphist "
Sql$ = Sql$ & "where cia='" & wcia & "' and "
Sql$ = Sql$ & " placod='" & Trim(Txtcodpla.Text) & "' and "
Sql$ = Sql$ & " codauxinterno='" & Trim(Lblcodaux.Caption) & "' and proceso='" & VTipobol & "' "
Sql$ = Sql$ & " and fechaproceso='" & Format(VFProceso, FormatFecha) & "' and semana='" & VSemana & "' and status='" & Mstatus & "'"
If (fAbrRst(Rs, Sql$)) Then
   If Rs(0) < 0 Then
      EsNegativo = True
   End If
End If
If Rs.State = 1 Then Rs.Close
'Fin Negativo

If vTipoTra = "01" Then
    Sql$ = "update platemphist set h22=h06/8 where h06<>0"
    cn.Execute Sql
End If

Sql$ = "update platemphist set h14=h01/8 where h01<>0 and h14=0"
cn.Execute Sql
Sql$ = ""

If wtipodoc = True Then
   Sql$ = "insert into plahistorico select * from platemphist"
Else
   Sql$ = "insert into plaquincena select * from platemphist"
End If
Sql$ = Sql$ & " where cia='" & wcia & "' and placod='" & Trim(Txtcodpla.Text) & "' and codauxinterno='" & Lblcodaux.Caption & "' and proceso='" & VTipobol & "' "
Sql$ = Sql$ & "and fechaproceso='" & Format(VFProceso, FormatFecha) & "' and semana='" & VSemana & "' and status='" & Mstatus & "'"

cn.Execute Sql

lItems = lItems + 1
If bImportar = False Then
   Sql$ = "SELECT ISNULL(MAX(id_boleta),1) FROM plahistorico where cia='" & wcia & "' and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " and proceso='" & VTipobol & "' and placod='" & Trim(Txtcodpla.Text) & "'"
   If (fAbrRst(Rs, Sql$)) Then LblId.Caption = Rs(0) Else LblId.Caption = ""
   Rs.Close: Set Rs = Nothing
End If

If VTipobol = "04" And lCalculaLiqui Then
   Sql$ = "SELECT MAX(id_boleta) FROM plahistorico where cia='" & wcia & "' and placod='" & Trim(Txtcodpla.Text) & "' and proceso='04' "
   If (fAbrRst(Rs, Sql$)) Then
      Sql$ = "Update platmphisliquid set id_boleta=" & Rs(0) & ""
      cn.Execute Sql
      Sql$ = "Insert into plahisliquid select * from platmphisliquid"
      cn.Execute Sql
   End If
   Rs.Close: Set Rs = Nothing
End If

Sql$ = "delete from platemphist where cia='" & wcia & "' and placod='" & Trim(Txtcodpla.Text) & "'"
Sql$ = Sql$ & " and cia='" & wcia & "' and placod='" & Txtcodpla.Text & "' and codauxinterno='" & Lblcodaux.Caption & "' and proceso='" & VTipobol & "' "
Sql$ = Sql$ & "and fechaproceso='" & Format(VFProceso, FormatFecha) & "' and semana='" & VSemana & "' and status='" & Mstatus & "'"
cn.Execute Sql

'Actualiza Fecha de Cese en el Maestro


'Cambiar Situcacion de Trabajador cuando Cesa (Situacuin EPS)
If IsDate(TxtFecCese.Text) And TxtFecCese.Enabled = True Then
   Dim lSitEPS As String
   Sql$ = "select afiliado_eps_serv from planillas where cia='" & wcia & "' and placod='" & Txtcodpla.Text & "' and status<>'*'"
    cn.CursorLocation = adUseClient
    Set Rs = New ADODB.Recordset
    Set Rs = cn.Execute(Sql$)
    If Rs.RecordCount > 0 Then
       If Rs(0) Then lSitEPS = "02" Else lSitEPS = "06"
    End If
    Rs.Close
    
    Sql$ = "Update Planillas set fcese = "
    Sql$ = Sql$ & IIf(IsDate(TxtFecCese.Text), "'" & Format(TxtFecCese, FormatFecha) & "'", "Null")
    Sql$ = Sql$ & ", mot_fin_periodo = "
    Sql$ = Sql$ & "'" & IIf(cbo_TipMotFinPer.ReturnCodigo = -1, "", Format(cbo_TipMotFinPer.ReturnCodigo, "00")) & "' "
    Sql$ = Sql$ & ", estado_eps = '" & lSitEPS & "' "
    Sql$ = Sql$ & "where cia='" & wcia & "' and placod='" & Txtcodpla.Text & "' and status<>'*'"
    
    cn.Execute Sql$

    
End If
'Fin Actualiza Fecha de Cese

'Graba Activo asignado al trabajador
If VTipobol = "01" Then
   If Trim(TxtAct4.Text) = Trim(TxtAct3.Text) Then TxtAct4.Text = "": TxtHor4.Text = 0
   If Trim(TxtAct4.Text) = Trim(TxtAct2.Text) Then TxtAct4.Text = "": TxtHor4.Text = 0
   If Trim(TxtAct4.Text) = Trim(TxtAct1.Text) Then TxtAct4.Text = "": TxtHor4.Text = 0
   If Trim(TxtAct3.Text) = Trim(TxtAct2.Text) Then TxtAct3.Text = "": TxtHor3.Text = 0
   If Trim(TxtAct3.Text) = Trim(TxtAct1.Text) Then TxtAct3.Text = "": TxtHor3.Text = 0
   If Trim(TxtAct2.Text) = Trim(TxtAct1.Text) Then TxtAct2.Text = "": TxtHor2.Text = 0
   TxtHor4.Text = Val(TxtHor4.Text)
   TxtHor3.Text = Val(TxtHor3.Text)
   TxtHor2.Text = Val(TxtHor2.Text)
   TxtHor1.Text = Val(TxtHor1.Text)
   
     
   Sql = "Usp_Pla_Trab_Activo '" & wcia & "'," & (Mid(VFProceso, 7, 4)) & ", " & (Mid(VFProceso, 4, 2)) & ",'" & VSemana & "','" & Trim(Txtcodpla.Text) & "','" & Trim(TxtAct1.Text) & "','" & Trim(TxtAct2.Text) & "','" & Trim(TxtAct3.Text) & "','" & Trim(TxtAct4.Text) & "','" & wuser & "',1," & Trim(TxtHor1.Text) & "," & Trim(TxtHor2.Text) & "," & Trim(TxtHor3.Text) & "," & Trim(TxtHor4.Text)
   cn.Execute Sql$
End If

'Actualiza Centro de Costo en Maestro del Trabajador
'If rsccosto.RecordCount > 0 Then rsccosto.MoveFirst
'Dim mCodAreaCosto As String
'Dim mPorcCosto As Double
'mPorcCosto = 0: mCodAreaCosto = ""
'
'Sql$ = "delete planilla_ccosto where cia='" & wcia & "' and placod='" & Txtcodpla.Text & "' and status<>'*'"
'cn.Execute Sql
'
'Do While Not rsccosto.EOF
'    If Not IsNull(rsccosto!Monto) Then
'      If rsccosto!Monto <> 0 Then
'         If rsccosto!Monto > mPorcCosto Then mPorcCosto = rsccosto!Monto: mCodAreaCosto = Trim(rsccosto!codigo & "")
'         Sql$ = "insert into planilla_ccosto values( '" & wcia & "','" & Txtcodpla.Text & "','" & Trim(rsccosto!codigo & "") & "'," & rsccosto!Monto & ",'')"
'         cn.Execute Sql
'      End If
'    Else
'        Exit Do
'    End If
'   rsccosto.MoveNext
'Loop
'Sql$ = "update planillas set placos='" & mCodAreaCosto & "' where cia='" & wcia & "' and placod='" & Txtcodpla.Text & "' and status<>'*'"
'cn.Execute Sql

'Fin Actualiza Centro Costo

'cn.CommitTrans
'If Not wImportarXls Then
   Sql$ = wFinTrans & " GRABAR_BOLETA"
   cn.Execute Sql$
'End If

NroTrans = 0
'On Error GoTo MyErr

Dim rsTemporal As ADODB.Recordset
'ACTIVAR PARA CONTRATOS
'Cadena = Empty
'Cadena = "SP_FECHA_CUMPLIMIENTO_CONTRATO " & _
'        "'" & wcia & "', " & _
'        "'" & Trim(Txtcodpla.Text) & "', " & _
'        "'" & mFormato_Fecha(VFProceso) & "'"
'Set rsTemporal = OpenRecordset(Cadena, cn)
'If Not rsTemporal.EOF Then
'    If Val(rsTemporal!FALTAN) = 0 Then
'        MsgBox "El Contrato del Sr(a) " & Lblnombre.Caption & ", aun no se ha renovado el contrato.", vbExclamation + vbOKOnly, "Contrato Vencido"
'    ElseIf Val(rsTemporal!FALTAN) < 0 Then
'        MsgBox "El Contrato del Sr(a) " & Lblnombre.Caption & ", aun no se ha renovado el contrato." & vbCrLf & "El contrato ha vencido hace " & Abs(rsTemporal!FALTAN) & " día(s).", vbExclamation + vbOKOnly, "Contrato Vencido"
'    Else
'        MsgBox "El Contrato del Sr(a) " & Trim(Lblnombre.Caption) & " esta por vencer en " & rsTemporal!FALTAN & ", Fecha de vencimiento :" & Format(rsTemporal!fec_fin, "dd/MM/yyyy"), vbExclamation + vbOKOnly, "Vencimiento de Contrato"
'    End If
'End If

If EsNegativo Then MsgBox "Se generó Boleta con neto en negativo para el trabajador => " & Txtcodpla.Text, vbInformation

Erase ArrDsctoCTACTE

If bImportar = False And LblId.Caption <> "" Then
   Ver_Calculo
Else
   Limpia_Boleta
End If

Screen.MousePointer = vbDefault
Exit Sub

ErrorTrans:
If NroTrans = 1 Then
    'If Not wImportarXls Then
       Sql$ = wCancelTrans & " GRABA_BOLETA"
       cn.Execute Sql$
    'End If
End If
Erase ArrDsctoCTACTE
Limpia_Boleta
Screen.MousePointer = vbDefault
MsgBox Err.Description, vbCritical, Me.Caption

End Sub
Private Sub Ver_Calculo()
  LstIngresos.ListItems.Clear
  LstDeducciones.ListItems.Clear
  LstAportaciones.ListItems.Clear
   If LblId.Caption = "" Then Exit Sub
   
   Sql$ = "usp_pla_Vista_Previa_Boleta " & LblId.Caption & ""
   If (fAbrRst(Rs, Sql$)) Then
      Rs.MoveFirst
      Do While Not Rs.EOF
         If Rs!tipo = "1" Or Rs!tipo = "2" Then
            If Rs!tipo = "2" Then
               Set Items = LstIngresos.ListItems.Add(, , "")
               Items.SubItems(1) = ""
               Items.SubItems(2) = ""
            End If
            
            Set Items = LstIngresos.ListItems.Add(, , Trim(Rs!Descripcion & ""))
            If Rs!horas <> 0 Then Items.SubItems(1) = fCadNum(Rs!horas, "##0.00") Else Items.SubItems(1) = ""
            If Rs!importe <> 0 Then Items.SubItems(2) = fCadNum(Rs!importe, "##,###,##0.00") Else Items.SubItems(2) = ""
         End If
         
         If Rs!tipo = "3" Or Rs!tipo = "4" Then
            If Rs!tipo = "4" Then
               Set Items = LstDeducciones.ListItems.Add(, , "")
               Items.SubItems(1) = ""
            End If
            
            Set Items = LstDeducciones.ListItems.Add(, , Trim(Rs!Descripcion & ""))
            If Rs!importe <> 0 Then Items.SubItems(1) = fCadNum(Rs!importe, "##,###,##0.00") Else Items.SubItems(1) = ""
         End If
         
         If Rs!tipo = "5" Or Rs!tipo = "6" Then
            If Rs!tipo = "6" Then
               Set Items = LstAportaciones.ListItems.Add(, , "")
               Items.SubItems(1) = ""
            End If
            
            Set Items = LstAportaciones.ListItems.Add(, , Trim(Rs!Descripcion & ""))
            If Rs!importe <> 0 Then Items.SubItems(1) = fCadNum(Rs!importe, "##,###,##0.00") Else Items.SubItems(1) = ""
         End If
         
         If Rs!tipo = "7" Then LstNeto.Caption = "TOTAL NETO => " & fCadNum(Rs!importe, "##,###,##0.00") & "  "
         
         Rs.MoveNext
      Loop
   End If
   Rs.Close: Set Rs = Nothing
   
   PnlPreView.Visible = True
   SSCommand1.SetFocus
   
End Sub
Public Sub Limpia_Boleta()
Frame2.Enabled = True
Frame3.Enabled = True
Frame4.Enabled = True
Frame5.Enabled = True
PnlPreView.Visible = False
BtnVerCalculo.Visible = False

VNewBoleta = True
LbLSemBolVac.Caption = ""
LblId.Caption = ""
Txtcodpla.Text = ""
Lblnombre.Caption = ""
Lblctacte.Caption = "0.00"
Lblcodaux.Caption = ""
Lblcodafp.Caption = ""
LblPensiones.Caption = ""
lblAfpTipoComision.Caption = ""
LblCodTrabSunat.Caption = ""
Lblnumafp.Caption = ""
LblBasico.Caption = ""
Lbltope.Caption = ""
Lblcargo.Caption = ""
LblPlanta.Caption = ""
VAltitud = ""
VVacacion = ""
VArea = ""
VFechaNac = ""
VFechaJub = ""
VJubilado = ""
LblFingreso.Caption = ""
TxtAct1.Text = "": TxtAct2.Text = "": TxtAct3.Text = "": TxtAct4.Text = ""
TxtHor1.Text = 0: TxtHor2.Text = 0: TxtHor3.Text = 0: TxtHor4.Text = 0
TxtFecCese.Text = "__/__/____"
cbo_TipMotFinPer.ListIndex = -1
TxtFecCese.Enabled = True
cbo_TipMotFinPer.Enabled = True

Txtcodpla.Enabled = True
'Txtcodpla.SetFocus
Frame2.Enabled = False
Frame3.Enabled = False
Frame4.Enabled = False
Frame5.Enabled = False

If rshoras.RecordCount > 0 Then rshoras.MoveFirst
Do While Not rshoras.EOF
   rshoras.Delete
   rshoras.MoveNext
Loop

If rspagadic.RecordCount > 0 Then rspagadic.MoveFirst
Do While Not rspagadic.EOF
   rspagadic!Monto = 0
   rspagadic.MoveNext
Loop

If rsdesadic.RecordCount > 0 Then rsdesadic.MoveFirst
Do While Not rsdesadic.EOF
   rsdesadic!Monto = 0
   rsdesadic.MoveNext
Loop

If rsccosto.RecordCount > 0 Then rsccosto.MoveFirst
Do While Not rsccosto.EOF
   rsccosto.Delete
   rsccosto.MoveNext
Loop

'If Not Dgrdhoras.DataSource Is Nothing Then
'    Set Dgrdhoras.DataSource = Nothing
'    If Not rshoras Is Nothing Then Set rshoras = Nothing
'End If
'
'If Not DgrdPagAdic.DataSource Is Nothing Then
'    Set DgrdPagAdic.DataSource = Nothing
'    If Not rspagadic Is Nothing Then Set rspagadic = Nothing
'End If
'
'If Not DgrdDesAdic.DataSource Is Nothing Then
'    Set DgrdDesAdic.DataSource = Nothing
'    If Not rsdesadic Is Nothing Then Set rsdesadic = Nothing
'End If
'
'If Not Dgrdccosto.DataSource Is Nothing Then
'    Set Dgrdccosto.DataSource = Nothing
'    If Not rsccosto Is Nothing Then Set rsccosto = Nothing
'End If

'Call Crea_Rs

'Call Form_Load

End Sub
Private Function Verifica_Boleta(mGraba As Boolean) As Boolean
On Error GoTo CORRIGE
If wtipodoc = True Then
   Select Case VPerPago
          Case Is = "02"
                If VTipobol = "03" Then
                    Sql$ = "select id_boleta from plahistorico where cia='" & wcia & "' and proceso='" & VTipobol & "' and placod='" & Txtcodpla.Text & "' and Year(fechaproceso)=" & Vano & " and  month(fechaproceso)=" & Vmes & " AND status<>'*' "
               Else
                    Sql$ = "select id_boleta from plahistorico where cia='" & wcia & "' and proceso='" & VTipobol & "' and placod='" & Txtcodpla.Text & "' and semana='" & VSemana & "' and Year(fechaproceso)=" & Vano & " AND month(fechaproceso)=" & Vmes & " and  status<>'*' "
                End If
          Case Is = "04"
               Sql$ = "select id_boleta from plahistorico where cia='" & wcia & "' and proceso='" & VTipobol & "' and placod='" & Txtcodpla.Text & "'  and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " and status<>'*' "
            Case Is = "03"
               Sql$ = "select id_boleta from plahistorico where cia='" & wcia & "' and proceso='" & VTipobol & "' and placod='" & Txtcodpla.Text & "'  and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " and status<>'*' "
   End Select
Else
   Sql$ = "select placod from plaquincena where cia='" & wcia & "' and placod='" & Txtcodpla.Text & "'  and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " and status<>'*' "
End If
cn.CursorLocation = adUseClient
Set Rs = New ADODB.Recordset
Set Rs = cn.Execute(Sql$, 64)
If Not Rs.EOF Then
   If wtipodoc = True Then
      If Not mGraba And LblId.Caption = "" Then
         If Rs.RecordCount > 0 Then
            LblId.Caption = ""
            MsgBox "Ya existe boleta para el periodo, no podra grabar" & Chr(13) & "Si desea modificar o eliminar debe seleccionar la boleta", vbInformation
            Verifica_Boleta = False
         End If
       Else
          Verifica_Boleta = False
       End If
   Else
      Verifica_Boleta = False
   End If
Else
   LblId.Caption = ""
   Verifica_Boleta = True
End If
If w_error = 1 Then MsgBox "No Procede Grabacion", vbInformation, Me.Caption
'If rs.RecordCount > 0 Then Verifica_Boleta = False Else Verifica_Boleta = True
If Rs.State = 1 Then Rs.Close
 Exit Function
CORRIGE:
        MsgBox "Error : " & Err.Description, vbCritical, Me.Caption
End Function
Public Sub Elimina_Boleta()

If Cierre_Planilla(Year(VFProceso), Month(VFProceso)) Then Exit Sub

If wtipodoc = True Then
    If wImportarXls = False Then
        Mgrab = MsgBox("Seguro de Eliminar Boleta", vbYesNo + vbQuestion, TitMsg)
    Else
        'IMPLEMENTACION GALLOS
        'CONFIRMANDO ELIMINAR BOLETA
        Mgrab = 6
    End If
Else

    If wImportarXls = False Then
       
       Sql$ = "select placod from plahistorico " _
       & "where cia='" & wcia & "' and proceso='01' and Year(fechaproceso) = " & Vano & " And month( fechaproceso) = " & Vmes & " " _
       & "and status<>'*' and placod='" & Txtcodpla.Text & "'"
       
       If (fAbrRst(Rs, Sql$)) Then
          MsgBox "Ya se genero la Boleta de Pago " & Chr(13) & "No se Puede Anular el Adelanto de Quincena", vbCritical, "Sistema de Planilla"
          Exit Sub
       End If
       Mgrab = MsgBox("Seguro de Eliminar Adelanto de Quincena", vbYesNo + vbQuestion, TitMsg)
    
    Else
       Mgrab = 6
    End If

End If

If Mgrab <> 6 Then Exit Sub
On Error GoTo MyErr
Dim InTrans As Boolean

'If Not wImportarXls Then
   Sql$ = wInicioTrans & " ELIMINA_BOLETA"
   cn.Execute Sql$
'End If
InTrans = True

If wtipodoc = True Then
   If wImportarXls Then
      Select Case VPerPago
          Case Is = "02"
               Sql$ = "update plahistorico set status='*',fec_modi=" & FechaSys & ",user_modi='" & wuser & "' " _
               & "where cia='" & wcia & "' and proceso='" & VTipobol & "' and semana='" & VSemana & "' and year(fechaproceso)=" & Vano & " And month( fechaproceso) = " & Vmes & " " _
               & "and status<>'*' and placod='" & Txtcodpla.Text & "'"
          Case Is = "04"
               Sql$ = "update plahistorico set status='*',fec_modi=" & FechaSys & ",user_modi='" & wuser & "' " _
               & "where cia='" & wcia & "' and proceso='" & VTipobol & "' and Year(fechaproceso) = " & Vano & " And month( fechaproceso) = " & Vmes & " " _
               & "and status<>'*' and placod='" & Txtcodpla.Text & "'"
      End Select
   Else
      Sql$ = "update plahistorico set status='*',fec_modi=" & FechaSys & ",user_modi='" & wuser & "' where Id_Boleta=" & LblId.Caption & ""
   End If
Else
   Sql$ = "update plaquincena set status='*',fec_modi=" & FechaSys & ",user_modi='" & wuser & "' " _
   & "where cia='" & wcia & "' and Year(fechaproceso) = " & Vano & " And month( fechaproceso) = " & Vmes & " " _
   & "and status<>'*' and placod='" & Txtcodpla.Text & "'"
End If

cn.Execute Sql$

'Eliminamos activo asignado en la boleta
If VTipobol = "01" Then
   If Trim(TxtHor1.Text) = "" Then TxtHor1.Text = 0
   If Trim(TxtHor2.Text) = "" Then TxtHor2.Text = 0
   If Trim(TxtHor4.Text) = "" Then TxtHor4.Text = 0
   If Trim(TxtHor3.Text) = "" Then TxtHor3.Text = 0
   
   Sql = "Usp_Pla_Trab_Activo '" & wcia & "'," & (Mid(VFProceso, 7, 4)) & ", " & (Mid(VFProceso, 4, 2)) & ",'" & VSemana & "','" & Trim(Txtcodpla.Text) & "','" & Trim(TxtAct1.Text) & "','" & Trim(TxtAct2.Text) & "','" & Trim(TxtAct3.Text) & "','" & Trim(TxtAct4.Text) & "','" & wuser & "',3," & Trim(TxtHor1.Text) & "," & Trim(TxtHor2.Text) & "," & Trim(TxtHor3.Text) & "," & Trim(TxtHor4.Text)
   cn.Execute Sql$
End If

'If Not wImportarXls Then
   Sql$ = wFinTrans & " ELIMINA_BOLETA"
   cn.Execute Sql$
'End If
InTrans = False

If wImportarXls = False Then
    MsgBox "Registros Eliminado Satisfactoriamente.", vbInformation + vbOKOnly, "Sistema"
End If

MyErr:
If InTrans Then
    'cn.RollbackTrans
    'If Not wImportarXls Then
       Sql$ = wCancelTrans & " ELIMINA_BOLETA"
       cn.Execute Sql$
    'End If
End If
If Err.Number <> 0 Then
    MsgBox Err.Description, vbExclamation, Err.Source
    Err.Clear
End If

'Call Limpia_Boleta

If wImportarXls = False Then
    Unload Me
End If

End Sub

Public Function F01(concepto As String, Optional ByVal pSemana As Integer) As String 'INGRESOS
Dim rsF01 As ADODB.Recordset
Dim mFactor As Currency
Dim nHijos As Integer
Dim RX As New ADODB.Recordset
Dim sSQL As String
Dim pSemanaPago As Integer
Dim sPeriodoPago As String
Dim nsemanasmes As Integer
Dim SemanaAnt As Integer
Dim NroAbonos As Integer
'Dim FechaiSemana As Date
Dim importeAplicado As Currency
Dim importexAplicar As Currency
mFactor = 0
nHijos = 0
F01 = ""

'mgirao añadiendo practicantes 26062019
If vTipoTra = "05" And concepto <> "01" Then Exit Function

Select Case concepto
      
       Case Is = "01" 'BASICO
            
            F01 = "select round((b.importe/factor_horas)*a.h01,2) as basico from platemphist a,plaremunbase b "
            F01 = F01 & "where a.cia='" & Trim(wcia) & "' and a.status='" & Mstatus & "' and a.placod='" & Trim(Txtcodpla.Text) & "' "
            F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
            F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & Trim(concepto) & "' and b.status<>'*'"
            
       Case Is = "02" 'ASIGNACION FAMILIAR
            sSQL = "SELECT semana_pago FROM PLACONSTANTE WHERE STATUS!='*' AND CIA='" & wcia & "' and codinterno='02' and tipomovimiento='02'"
            If (fAbrRst(RX, sSQL)) Then pSemanaPago = RX(0)
            If RX.State = 1 Then RX.Close

            If vTipoTra = "02" Then
            
                sSQL = "select Tipo from plaremunbase where cia = '" & wcia & "' and status != '*' and placod = '" & Trim(Txtcodpla.Text) & "' and concepto = '" & concepto & "'"
                If (fAbrRst(RX, sSQL)) Then sPeriodoPago = RX(0)
                If RX.State = 1 Then RX.Close
                
                If sPeriodoPago = "02" Then
                    ''F01 = "select round((b.importe/(case factor_horas when 48 then 60 else factor_horas end)),6) * 8 * (case when DateDiff(day, getdate(), fechaingreso) <= 7 then 7 else DateDiff(day, getdate(), fechaingreso) +1 end) from platemphist a,plaremunbase b "
                    F01 = "select round((b.importe/(case factor_horas when 48 then 60 else factor_horas end)),6) * 8 * " & DiasAsignacion & " from platemphist a,plaremunbase b "
                    F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Trim(Txtcodpla.Text) & "' "
                    F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
                    F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & concepto & "' and b.status<>'*'"
                  Else
                    If Semana_Calcular(pSemana, pSemanaPago, Year(FrmCabezaBol.Cmbfecha), wcia) Then
                        F01 = "select round((b.importe/factor_horas)*(factor_horas),2) as basico from platemphist a,plaremunbase b "
                        F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Trim(Txtcodpla.Text) & "' "
                        F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
                        F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & concepto & "' and b.status<>'*'"
                        F01 = "select ( " & F01 & " ) /2 "
                    End If
               End If
 
               
            Else
               Dim rsTemporal As ADODB.Recordset
               Dim t_Asignacion_Fam As Integer
               t_Asignacion_Fam = 0
'              Cadena = "select isnull(asignacion_fam_prorrateada,0) as t_asignacion from planillas where cia = '" & wcia & "' and status != '*' and placod = '" & Trim(Txtcodpla.Text) & "'"
'              Set rsTemporal = OpenRecordset(Cadena, cn)
'              If Not rsTemporal.EOF Then
'                 t_Asignacion_Fam = rsTemporal!t_asignacion
'              End If
'              If rsTemporal.State = 1 Then rsTemporal.Close: Set rsTemporal = Nothing

            'Mgirao Valido si el empleado esta en vacaciones 31/01/2019 Kprorras
            Dim TAño As Integer
            Dim tmes As Integer
            If Vmes = 1 Then
               TAño = Vano - 1
               tmes = 12
            Else
               TAño = Vano
               tmes = Vmes - 1
            End If
            Dim TmpFecha As String
            TmpFecha = "01" + "/" + Right("00" + Trim(Str(tmes)), 2) + "/" + Trim(Str(TAño))
           
            Cadena = "select fechavacai,fechavacaf from plahistorico where cia = '" & wcia & "' and status != '*' and placod = '" & Trim(Txtcodpla.Text) & "'" & " and proceso='02'"
            'Cadena = Cadena & "and year(fechaproceso)=" & TAño & " and month(fechaproceso)=" & TMes & " "
            Cadena = Cadena & "and fechaproceso>='" & TmpFecha & "'"
            Set rsTemporal = OpenRecordset(Cadena, cn)
            If Not rsTemporal.EOF Then
              If Month(rsTemporal(0)) = Vmes Or Month(rsTemporal(1)) = Vmes Then
                 t_Asignacion_Fam = 1
              Else
                  t_Asignacion_Fam = 0
              End If
            End If
            If rsTemporal.State = 1 Then rsTemporal.Close: Set rsTemporal = Nothing

            If t_Asignacion_Fam = 0 Then
               F01 = "select importe from plaremunbase where cia = '" & wcia & "' and status != '*' and placod = '" & Trim(Txtcodpla.Text) & "' and concepto = '" & concepto & "'"
            Else
               F01 = "select round((b.importe/factor_horas)*(a.h01+a.h08+a.h23+h04+h05+h25+h29),2) as basico from platemphist a,plaremunbase b "
               'F01 = "select round((b.importe/factor_horas)*(factor_horas),2) as basico from platemphist a,plaremunbase b "
               F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Trim(Txtcodpla.Text) & "' "
               F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
               F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & concepto & "' and b.status<>'*'"
            End If
            End If
            
          If VTipobol = "11" Or VTipobol = "04" Then
              F01 = "Select 0"
          End If
          
            
       Case Is = "03" 'ASIGNACION MOVILIDAD
            ' CAMBIO DE H14 POR H01  {>MA<} 10/06/2007
            'F01 = "select round((b.importe/factor_horas)*a.H01+A.H03,2) as basico from platemphist a,plaremunbase b "
            'If WValor03 = 0 Then
                F01 = "select round((b.importe/factor_horas)*a.H14*8,2) + isnull((case when year(a.fechaproceso)=2020 then (case when a.semana=17 then (select importe from reg_mov where placod='" & Txtcodpla.Text & "' ) else 0 end) else 0 end),0) as basico from platemphist a,plaremunbase b "
                F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Txtcodpla.Text & "' "
                F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
                F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & concepto & "' and b.status<>'*'"
            'Else
            '    F01 = "Select " & Round(WValor03, 2)
            'End If
            Dim x As Integer
            x = 1
            
       Case Is = "04" 'BONIFICACION T. SERVICIO
            'F01 = "select round((b.importe/factor_horas)*(a.h01+a.h02+a.h03+a.h04+a.h05+a.h12),2) as basico from platemphist a,plaremunbase b "
            F01 = "select round((b.importe/factor_horas)*(a.h01+a.h02+a.h04+a.h05+a.h12+h25+h29),2) as basico from platemphist a,plaremunbase b "
            F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Trim(Txtcodpla.Text) & "' "
            F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
            F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & concepto & "' and b.status<>'*'"
           
       Case Is = "05" 'INCREMENTO AFP 10.23%

            '*************************************************************
            F01 = "select round((b.importe/factor_horas*(a.h01+a.h02+a.h03+a.h04+a.h05+a.h12+a.h23+h25+h29)),2) as basico from platemphist a,plaremunbase b "
            F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Txtcodpla.Text & "' "
            F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
            F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & concepto & "' and b.status<>'*'"
            '*****************************************************************************************
            
       Case Is = "06" 'INCREMENTO AFP 3%
            
            F01 = "select round((b.importe/factor_horas*(a.h01+a.h02+a.h03+a.h04+a.h05+a.h12+a.h23+h25+h29)),2) as basico from platemphist a,plaremunbase b "
            F01 = F01 & "where a.cia='" & Trim(wcia) & "' and a.status='" & Mstatus & "' and a.placod='" & Trim(Txtcodpla.Text) & "' "
            F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
            F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & concepto & "' and b.status<>'*'"
            '******************************************************************************************
       
       Case Is = "07" 'BONIFICACION PATERNIDAD
            F01 = "select round((b.importe/factor_horas)*(a.h23),2) as basico from platemphist a,plaremunbase b "
            F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Txtcodpla.Text & "' "
            F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
            F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
       Case Is = "08" 'SOBRETASA
            Sql$ = "select factor from platasaanexo where modulo='02' and tipomovimiento='" & Trim(Lbltope.Caption) & "' and codinterno='" & concepto & "' and tipotrab='" & vTipoTra & "' and status<>'*'"
             
            If (fAbrRst(rsF01, Sql$)) Then
                 mFactor = rsF01!factor
                 If rsF01.State = 1 Then rsF01.Close
                 If vTipoTra = "01" Then
                    F01 = "select round(((b.importe * " & mFactor & ")/factor_horas)*(a.h13),2) as basico from platemphist a,plaremunbase b "
                    F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Txtcodpla.Text & "' "
                    F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
                    F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
                 Else
                    F01 = "select ROUND(" & mFactor & " *(a.h13),2) as basico from platemphist a "
                    F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Trim(Txtcodpla.Text) & "' "
                    F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
                 End If
            End If
       Case Is = "09" 'DOMINICAL
            If vTipoTra <> "01" Or wGrupoPla = "01" Then 'En grupo Gallos se disgrega remuneracion y se calcula dominical
               F01 = "select round(((b.importe/factor_horas)*h02),2) as basico from platemphist a,plaremunbase b "
               F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Txtcodpla.Text & "' "
               F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
               F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
            End If
       Case Is = "10" '2 PRI EXTRAS L-S
            Sql$ = "select factor from platasaanexo where  cia='" & wcia & "' and modulo='02' and tipomovimiento='" & Trim(Lbltope.Caption) & "' and codinterno='" & concepto & "' and tipotrab='" & vTipoTra & "' and status<>'*'"
            'Sql$ = "select factor from platasaanexo where modulo='02' and tipomovimiento='" & Trim(Lbltope.Caption) & "' and codinterno='" & concepto & "' and tipotrab='" & vTipoTra & "' and cargo='" & Trim(Lblcargo.Caption) & "' and status<>'*'"
            If (fAbrRst(rsF01, Sql$)) Then
               mFactor = rsF01!factor
               If rsF01.State = 1 Then rsF01.Close
               If wGrupoPla = "01" Then 'Sin Asignacion Familiar para grupo Gallos
                 F01 = "SELECT SUM(ROUND(((IMPORTE/(FACTOR_HORAS+CASE WHEN TIPO='02' AND B.tipotrabajador='02' AND a.concepto IN ('02','05','06') THEN 12 ELSE 0 END))*" & mFactor & ")*c.h10,2)) as basico"
                 F01 = F01 & " FROM PLAREMUNBASE A INNER JOIN PLANILLAS B ON (B.PLACOD=A.PLACOD AND B.STATUS!='*' AND B.CIA=A.CIA AND B.CIA='" & wcia & "') INNER JOIN platemphist C ON (C.PLACOD=A.PLACOD AND B.STATUS!='*' and YEAR(fechaproceso)=" & Vano & " and MONTH(fechaproceso)=" & Vmes & ") "
                 F01 = F01 & " WHERE A.CIA='" & wcia & "' AND A.PLACOD='" & Trim(Txtcodpla.Text) & "' AND A.STATUS!='*' "
                 'and a.concepto not in ('02')"
               Else
                 F01 = "SELECT SUM(ROUND(((IMPORTE/(FACTOR_HORAS+CASE WHEN TIPO='02' AND B.tipotrabajador='02' AND a.concepto IN ('02','05','06') THEN 12 ELSE 0 END))*" & mFactor & ")*c.h10,2)) as basico"
                 F01 = F01 & " FROM PLAREMUNBASE A INNER JOIN PLANILLAS B ON (B.PLACOD=A.PLACOD AND B.STATUS!='*' AND B.CIA=A.CIA AND B.CIA='" & wcia & "') INNER JOIN platemphist C ON (C.PLACOD=A.PLACOD AND B.STATUS!='*' and YEAR(fechaproceso)=" & Vano & " and MONTH(fechaproceso)=" & Vmes & ") "
                 F01 = F01 & " WHERE A.CIA='" & wcia & "' AND A.PLACOD='" & Trim(Txtcodpla.Text) & "' AND A.STATUS!='*' and a.concepto not in ('03','08','22')"
                 'and a.concepto not in ('02','03','08','22')"
               End If
            End If
       Case Is = "11" 'EXTRAS D-F
            Sql$ = "select factor from platasaanexo where  cia='" & wcia & "' and modulo='02' and tipomovimiento='" & Trim(Lbltope.Caption) & "' and codinterno='" & concepto & "' and tipotrab='" & vTipoTra & "' and status<>'*'"
            'Sql$ = "select factor from platasaanexo where modulo='02' and tipomovimiento='" & Trim(Lbltope.Caption) & "' and codinterno='" & concepto & "' and tipotrab='" & vTipoTra & "' and cargo='" & Trim(Lblcargo.Caption) & "' and status<>'*'"
            If (fAbrRst(rsF01, Sql$)) Then
               mFactor = rsF01!factor
               If rsF01.State = 1 Then rsF01.Close
                'bts = convierte_cant
                F01 = "SELECT SUM(ROUND(((IMPORTE/(FACTOR_HORAS+CASE WHEN TIPO='02' AND B.tipotrabajador='02' AND a.concepto IN ('02','05','06') THEN 12 ELSE 0 END))*" & mFactor & ")*c.h11,2)) as basico"
                F01 = F01 & " FROM PLAREMUNBASE A INNER JOIN PLANILLAS B ON (B.PLACOD=A.PLACOD AND B.STATUS!='*' AND B.CIA='" & wcia & "') INNER JOIN platemphist C ON (C.PLACOD=A.PLACOD AND B.STATUS!='*' and YEAR(fechaproceso)=" & Vano & " and MONTH(fechaproceso)=" & Vmes & ") "
                F01 = F01 & " WHERE A.CIA='" & wcia & "' AND B.CIA ='" & wcia & "' AND A.PLACOD='" & Trim(Txtcodpla.Text) & "' AND A.STATUS!='*' and a.concepto not in ('03','08','22')"
                'and a.concepto not in ('02','03','08','22')"
             
             End If
            
       Case Is = "12" 'FERIADOS
            F01 = "select round((b.importe/factor_horas)*a.h03,2) as basico,A.H03 from platemphist a,plaremunbase b "
            F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Trim(Txtcodpla.Text) & "' "
            F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
            F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
       Case Is = "13" 'REINTEGROS
       Case Is = "14" 'VACACIONES (CONSTRUCCION CIVIL)
            If VVacacion = "S" Then
               Sql$ = "select factor from platasaanexo where  cia='" & wcia & "' and modulo='02' and tipomovimiento='" & Trim(Lbltope.Caption) & "' and codinterno='" & concepto & "' and tipotrab='" & vTipoTra & "' and status<>'*'"
               'Sql$ = "select factor from platasaanexo where modulo='02' and tipomovimiento='" & Trim(Lbltope.Caption) & "' and codinterno='" & concepto & "' and tipotrab='" & vTipoTra & "' and cargo='" & Trim(Lblcargo.Caption) & "' and status<>'*'"
               If (fAbrRst(rsF01, Sql$)) Then
                  mFactor = rsF01!factor
                  If rsF01.State = 1 Then rsF01.Close
                  F01 = "select round(((b.importe/factor_horas)* " & mFactor & ") *(a.h01),2) as basico from platemphist a,plaremunbase b "
                  F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Txtcodpla.Text & "' "
                  F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
                  F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
               End If
            End If
       Case Is = "15" 'GRATIFICACION
       Case Is = "16" 'OTROS PAGOS
       Case Is = "17" 'ASIGNACION ESCOLAR
            If vTipoTra = "01" Then
            End If
            If vTipoTra = "02" Then
            End If
            If vTipoTra = "05" Then
               Sql$ = "select ultima from plasemanas where cia='" & wcia & "' and semana='" & Format(VSemana, "00") & "' and ano='" & Vano & "' and status<>'*'"
               If (fAbrRst(rsF01, Sql$)) Then
                  If rsF01!ultima = "S" Then
                     Sql$ = "select factor from platasaanexo where  cia='" & wcia & "' and modulo='02' and tipomovimiento='" & Trim(Lbltope.Caption) & "' and codinterno='" & concepto & "' and tipotrab='" & vTipoTra & "' and status<>'*'"
                     'Sql$ = "select factor from platasaanexo where modulo='02' and tipomovimiento='" & Trim(Lbltope.Caption) & "' and codinterno='" & concepto & "' and tipotrab='" & vTipoTra & "' and cargo='" & Trim(Lblcargo.Caption) & "' and status<>'*'"
                     If (fAbrRst(rsF01, Sql$)) Then
                        nHijos = Numero_Hijos(Trim(Txtcodpla.Text), "S", "S", VFProceso, 18)
                        mFactor = rsF01!factor
                        If rsF01.State = 1 Then rsF01.Close
                        F01 = "select round((((b.importe/factor_horas)*8)* " & mFactor & ")/12 * " & nHijos & ",2) as basico from platemphist a,plaremunbase b "
                        F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Txtcodpla.Text & "' "
                        F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
                        F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
                     End If
                  End If
               End If
            End If
       Case Is = "18" 'UTILIDADES
       
       Case Is = "19" 'PERMISOS PAGADOS
            F01 = "select round((b.importe/factor_horas)*a.h04,2) as basico,A.H04 from platemphist a,plaremunbase b "
            F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Trim(Txtcodpla.Text) & "' "
            F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
            F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
       Case Is = "20" 'CONDICION DE TRABAJO
       
'            Sql$ = "select factor from platasaanexo where modulo='02' and tipomovimiento='" & Trim(Lbltope.Caption) & "' and codinterno='" & concepto & "' and tipotrab='" & vTipoTra & "' and cargo='" & Trim(Lblcargo.Caption) & "' and status<>'*'"
'            If (fAbrRst(rsF01, Sql$)) Then
'               mFactor = rsF01!factor
'               If rsF01.State = 1 Then rsF01.Close
'               F01 = "select round(((b.importe/factor_horas)* " & mFactor & ") *(a.h01),2) as basico from platemphist a,plaremunbase b "
'               F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Txtcodpla.Text & "' "
'               F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
'               F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
'            End If
            
        Case Is = "21" '3RA EX. LS
            Sql$ = "select factor from platasaanexo where  cia='" & wcia & "' and modulo='02' and tipomovimiento='" & Trim(Lbltope.Caption) & "' and codinterno='" & concepto & "' and tipotrab='" & vTipoTra & "' and status<>'*'"
            'Sql$ = "select factor from platasaanexo where modulo='02' and tipomovimiento='" & Trim(Lbltope.Caption) & "' and codinterno='" & concepto & "' and tipotrab='" & vTipoTra & "' and cargo='" & Trim(Lblcargo.Caption) & "' and status<>'*'"
            If (fAbrRst(rsF01, Sql$)) Then
                mFactor = rsF01!factor
                If rsF01.State = 1 Then rsF01.Close
                F01 = "SELECT SUM(ROUND(((IMPORTE/(FACTOR_HORAS+CASE WHEN TIPO='02' AND B.tipotrabajador='02' AND a.concepto IN ('02','05','06') THEN 12 ELSE 0 END))*" & mFactor & ")*c.h17,2)) as basico"
                F01 = F01 & " FROM PLAREMUNBASE A INNER JOIN PLANILLAS B ON (B.PLACOD=A.PLACOD AND B.STATUS!='*' AND B.CIA='" & wcia & "') INNER JOIN platemphist C ON (C.PLACOD=A.PLACOD AND B.STATUS!='*' and YEAR(fechaproceso)=" & Vano & " and MONTH(fechaproceso)=" & Vmes & ") "
                F01 = F01 & " WHERE A.CIA='" & wcia & "' AND A.PLACOD='" & Trim(Txtcodpla.Text) & "' AND A.STATUS!='*' and a.concepto not in ('03','08','22')"
             
            End If
       Case Is = "22" 'SOBRE TASA NOCHE
'            If vTipoTra = "02" Then
'                 Dim FSTN As Double
'                 FSTN = Calcula_Factor_Sobretasa_Noche
'                 F01 = "select sum(ROUND((b.importe/factor_horas)*(a.h20),2) + " & FSTN & " ) as basico from platemphist a,plaremunbase b "
'                 F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Trim(Txtcodpla.Text) & "' "
'                 F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
'                 F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='08' and b.status<>'*'"
'            Else
'              Sql$ = "select factor from platasaanexo where  cia='" & wcia & "' and modulo='02' and tipomovimiento='" & Trim(Lbltope.Caption) & "' and codinterno='" & concepto & "' and tipotrab='" & vTipoTra & "' and cargo='" & Trim(Lblcargo.Caption) & "' and status<>'*'"
'
'              If (fAbrRst(rsF01, Sql$)) Then
'                 mFactor = rsF01!factor
'                 If rsF01.State = 1 Then rsF01.Close
'                 F01 = "select ROUND(" & mFactor & " *(a.h20),2) as basico from platemphist a,plaremunbase b "
'                 F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Trim(Txtcodpla.Text) & "' "
'                 F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
'                 F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
'              End If
'            End If

              Sql$ = "select factor from platasaanexo where  cia='" & wcia & "' and modulo='02' and tipomovimiento='" & Trim(Lbltope.Caption) & "' and codinterno='08' and tipotrab='" & vTipoTra & "' and status<>'*'"

              If (fAbrRst(rsF01, Sql$)) Then
                 mFactor = rsF01!factor
                 
                 Dim FSTN As Double
                 FSTN = Calcula_Factor_Sobretasa_Noche
                 
                 If rsF01.State = 1 Then rsF01.Close
                 F01 = "select sum(ROUND(" & mFactor & "*(a.h20),2) + " & FSTN & " ) as basico from platemphist a "
                 F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Trim(Txtcodpla.Text) & "' "
                 F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
              End If
              
       Case Is = "23" 'INGRESO POR HIDRATACION
       Case Is = "24" 'NOCHE 2 PRI EX. LS.
            F01 = "select round((b.importe/factor_horas)*a.h18,2) as basico from platemphist a,plaremunbase b "
            F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Trim(Txtcodpla.Text) & "' "
            F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
            F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
       'comenta mario para habilitar ingreso bonificacion produccion
'            Sql$ = "select factor from platasaanexo where cia='" & wcia & "' and modulo='02' and tipomovimiento='" & Trim(Lbltope.Caption) & "' and codinterno='" & concepto & "' and tipotrab='" & vTipoTra & "' and status<>'*'"
'            'Sql$ = "select factor from platasaanexo where modulo='02' and tipomovimiento='" & Trim(Lbltope.Caption) & "' and codinterno='" & concepto & "' and tipotrab='" & vTipoTra & "' and cargo='" & Trim(Lblcargo.Caption) & "' and status<>'*'"
'            If (fAbrRst(rsF01, Sql$)) Then
'               mFactor = rsF01!factor
'               If rsF01.State = 1 Then rsF01.Close
'
'                F01 = "SELECT SUM(ROUND(((IMPORTE/(FACTOR_HORAS+CASE WHEN TIPO='02' AND B.tipotrabajador='02' AND a.concepto IN ('05','06') THEN 8 ELSE 0 END))*" & mFactor & ")*c.h18,2)) as basico"
'                F01 = F01 & " FROM PLAREMUNBASE A INNER JOIN PLANILLAS B ON (B.PLACOD=A.PLACOD AND B.STATUS!='*' AND B.CIA='" & wcia & "') INNER JOIN platemphist C ON (C.PLACOD=A.PLACOD AND B.STATUS!='*' and YEAR(fechaproceso)=" & Vano & " and MONTH(fechaproceso)=" & Vmes & ") "
'                F01 = F01 & " WHERE A.CIA='" & wcia & "' AND A.PLACOD='" & Trim(Txtcodpla.Text) & "' AND A.STATUS!='*' and a.concepto not in ('02','03')"
'           End If
       Case Is = "25" 'NOCHE 3ra Ex. LS.
            Sql$ = "select factor from platasaanexo where cia='" & wcia & "' and modulo='02' and tipomovimiento='" & Trim(Lbltope.Caption) & "' and codinterno='" & concepto & "' and tipotrab='" & vTipoTra & "' and status<>'*'"
            'Sql$ = "select factor from platasaanexo where modulo='02' and tipomovimiento='" & Trim(Lbltope.Caption) & "' and codinterno='" & concepto & "' and tipotrab='" & vTipoTra & "' and cargo='" & Trim(Lblcargo.Caption) & "' and status<>'*'"
            If (fAbrRst(rsF01, Sql$)) Then
               mFactor = rsF01!factor
               If rsF01.State = 1 Then rsF01.Close
                
               'F01 = "SELECT SUM(ROUND(((IMPORTE/(FACTOR_HORAS+CASE WHEN TIPO='02' AND B.tipotrabajador='02' AND a.concepto IN ('05','06') THEN 8 ELSE 0 END))*" & mFactor & ")*c.h19,2)) as basico"
                 F01 = "SELECT SUM(ROUND(((IMPORTE/(FACTOR_HORAS+CASE WHEN TIPO='02' AND B.tipotrabajador='02' AND a.concepto IN ('02','05','06') THEN 12 ELSE 0 END))*" & mFactor & "),2)) as basico"
               F01 = F01 & " FROM PLAREMUNBASE A INNER JOIN PLANILLAS B ON (B.PLACOD=A.PLACOD AND B.STATUS!='*' AND B.CIA='" & wcia & "') INNER JOIN platemphist C ON (C.PLACOD=A.PLACOD AND B.STATUS!='*' and YEAR(fechaproceso)=" & Vano & " and MONTH(fechaproceso)=" & Vmes & ") "
               F01 = F01 & " WHERE A.CIA='" & wcia & "' AND A.PLACOD='" & Trim(Txtcodpla.Text) & "' AND A.STATUS!='*' and a.concepto not in ('03')"
               'and a.concepto not in ('02','03')"
            End If
       Case Is = "26" 'INGRESO POR LUBRICACION
       Case Is = "27" 'INGRESO POR DOBLADO DE TURNO
       Case Is = "28" 'INGRESO POR REFRIGERIO
       Case Is = "29" 'COMISION DE VENTAS
            'MGIRAO 28/11/2016
            If VTipobol <> "04" Then
                'F01 = "SELECT valorcomision FROM VTA_COMISIONVENDEDOR_MENSUAL where año =" & Vano & " and Mes='" & Vmes & "' and cia='" & wcia & "' and ficha ='" & Trim(Txtcodpla.Text) & "'"
                'add jcms 290323 nuevo plan 2023
                F01 = "SELECT isnull(ImpSoles_ComisionVta_Total,0) as 'valorcomision' FROM VTA_COMISIONVENDEDOR_MENSUAL where año =" & Vano & " and Mes='" & Vmes & "' and cia='" & wcia & "' and ficha ='" & Trim(Txtcodpla.Text) & "'"
            End If
            
       Case Is = "30" 'BONIFICACION EXTRAORDINARIA
         'Calculo Liquidacion
          If VTipobol = "04" And lCalculaLiqui Then
             F01 = "Select Round((igano+igmes+igdia)*"
             F01 = F01 & "(Select "
             F01 = F01 & "Case AFILIADO_EPS_SERV "
             F01 = F01 & "when 1 then (Select importe from maestros_2 WHERE RIGHT(CIAMAESTRO,3) = '143' AND STATUS != '*' AND RTRIM(ISNULL(CODSUNAT,'')) != '' AND COD_MAESTRO2 = planillas.codigo_eps) "
             F01 = F01 & "else (select aportacion from placonstante where cia=planillas.cia and tipomovimiento='03' and codinterno='01' and aportacion<>0 and status<>'*') "
             F01 = F01 & "End "
             F01 = F01 & "From planillas "
             F01 = F01 & "Where placod='" & Trim(Txtcodpla.Text) & "' and status<>'*')/100,2) "
             F01 = F01 & "From platmphisliquid"
          End If
       Case Is = "31" 'LICENCIA SIN GOCE DE HABER
'            F01 = "select round((b.importe/factor_horas)*a.h26,2) as basico from platemphist a,plaremunbase b "
'            F01 = F01 & "where a.cia='" & Trim(wcia) & "' and a.status='" & Mstatus & "' and a.placod='" & Trim(Txtcodpla.Text) & "' "
'            F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
'            F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
       Case Is = "32" 'ENFERMEDADES PAGAS
                F01 = "select round((b.importe/factor_horas)*a.h05,2) as basico from platemphist a,plaremunbase b "
                F01 = F01 & "where a.cia='" & Trim(wcia) & "' and a.status='" & Mstatus & "' and a.placod='" & Trim(Txtcodpla.Text) & "' "
                F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
                F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
   
       Case Is = "33" 'GASTO DE VIAJES
       Case Is = "34" 'FERIADO 1ERO DE MAYO
            'Asignacion Familiar Incremento AFP
            Sql$ = "SELECT Sum((case factor_horas when 48 then importe*4 When 240 then importe End)/30) as Factor "
            Sql$ = Sql$ & "FROM plaremunbase WHERE Cia='" & Trim(wcia) & "' AND Concepto in('02','05','06') AND status<>'*' "
            Sql$ = Sql$ & "and placod='" & Trim(Txtcodpla.Text) & "' "
            If (fAbrRst(rsF01, Sql$)) Then
               If IsNull(rsF01!factor) Then
                  mFactor = 0
               Else
                  mFactor = rsF01!factor
               End If
               If rsF01.State = 1 Then rsF01.Close
            End If
            F01 = "select Case a.h15 When 0 then 0 else round((b.importe/factor_horas)*a.h15,2) + " & mFactor & " End as basico,A.H15 from platemphist a,plaremunbase b "
            F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Trim(Txtcodpla.Text) & "' "
            F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
            F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
            mFactor = 0
       Case Is = "35" 'ASIGNACION POR FALLECIMIENTO
       Case Is = "36" 'INGRESO POR PUNTUALIDAD
       Case Is = "37" 'INGRESO POR EMBOLSADO
            F01 = "select round((b.importe/factor_horas)*a.h28,2) as basico from platemphist a,plaremunbase b "
            F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Trim(Txtcodpla.Text) & "' "
            F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
            F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
       Case Is = "38" 'BONIF.PROD. 12/06/2008
            Sql$ = "select factor from platasaanexo where  cia='" & Trim(wcia) & "' and modulo='02' and tipomovimiento='" & Trim(Lbltope.Caption) & "' and codinterno='" & concepto & "' and tipotrab='" & vTipoTra & "' and status<>'*'"
            If (fAbrRst(rsF01, Sql$)) Then
               mFactor = rsF01!factor
               If rsF01.State = 1 Then rsF01.Close
                 F01 = "SELECT SUM(ROUND(((IMPORTE/(FACTOR_HORAS+CASE WHEN TIPO='02' AND B.tipotrabajador='02' AND a.concepto IN ('05','06') THEN 8 ELSE 0 END))*" & mFactor & ")*c.h24,2)) as basico"
                 F01 = F01 & " FROM PLAREMUNBASE A INNER JOIN PLANILLAS B ON (B.PLACOD=A.PLACOD AND B.STATUS!='*' AND B.CIA=A.CIA) INNER JOIN platemphist C ON (C.PLACOD=A.PLACOD AND B.STATUS!='*' and YEAR(fechaproceso)=" & Vano & " and MONTH(fechaproceso)=" & Vmes & ") "
                 F01 = F01 & " WHERE A.CIA='" & Trim(wcia) & "' AND A.PLACOD='" & Trim(Txtcodpla.Text) & "' AND A.STATUS!='*' and a.concepto not in ('02','03')"
            End If
       
       Case Is = "39" 'VACACIONES TRUNCAS
         'Calculo Liquidacion
          If VTipobol = "04" And lCalculaLiqui Then
             F01 = "Select ivano+ivmes+ivdia From platmphisliquid"
          End If
       Case Is = "40" 'GRATIFICACIONES TRUNCAS
         'Calculo Liquidacion
          If VTipobol = "04" And lCalculaLiqui Then
             F01 = "Select igano+igmes+igdia From platmphisliquid"
          End If
       Case Is = "41" 'LIQUIDACION
       Case Is = "42" 'SUBSIDIO POR MATERNIDAD
            If WValor42 = 0 Then
                F01 = "select round((b.importe/factor_horas)*a.h27,2) as basico from platemphist a,plaremunbase b "
                F01 = F01 & "where a.cia='" & Trim(wcia) & "' and a.status='" & Mstatus & "' and a.placod='" & Trim(Txtcodpla.Text) & "' "
                F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
                F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
            Else
                F01 = "Select " & WValor42
            End If
               ' F01 = "SELECT 253.66"
       Case Is = "43" 'SUBSIDIO POR ENFERMEDAD CALCULADO IGUAL AL BASICO
            If vTipoTra = "02" Then
            
                    F01 = "select round((b.importe/factor_horas)*(a.h06+a.h30),2) as basico from platemphist a,plaremunbase b "
                    F01 = F01 & "where a.cia='" & Trim(wcia) & "' and a.status='" & Mstatus & "' and a.placod='" & Trim(Txtcodpla.Text) & "' "
                    F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
                    F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"

            Else
                    'mgirao agregar asig fam
                    F01 = "select (isnull((select round((b.importe/factor_horas)*(a.h06+a.h30),2) as basico from platemphist a,plaremunbase b "
                    F01 = F01 & "where a.cia='" & Trim(wcia) & "' and a.status='" & Mstatus & "' and a.placod='" & Trim(Txtcodpla.Text) & "' "
                    F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
                    F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*' ),0) + "
                    F01 = F01 & "isnull((select round((b.importe/factor_horas)*(a.h06+a.h30),2) as basico from platemphist a,plaremunbase b "
                    F01 = F01 & "where a.cia='" & Trim(wcia) & "' and a.status='" & Mstatus & "' and a.placod='" & Trim(Txtcodpla.Text) & "' "
                    F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
                    F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='02' and b.status<>'*' ),0) + "
                    F01 = F01 & "isnull((select round((b.importe/factor_horas)*(a.h06+a.h30),2) as basico from platemphist a,plaremunbase b "
                    F01 = F01 & "where a.cia='" & Trim(wcia) & "' and a.status='" & Mstatus & "' and a.placod='" & Trim(Txtcodpla.Text) & "' "
                    F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
                    F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='05' and b.status<>'*' ),0) + "
                    F01 = F01 & "isnull((select round((b.importe/factor_horas)*(a.h06+a.h30),2) as basico from platemphist a,plaremunbase b "
                    F01 = F01 & "where a.cia='" & Trim(wcia) & "' and a.status='" & Mstatus & "' and a.placod='" & Trim(Txtcodpla.Text) & "' "
                    F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
                    F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='06' and b.status<>'*' ),0) )"
           
            End If
       
       Case Is = "44" 'INGRESO POR REEMPLAZO DE CARGO
       Case Is = "45" 'DEPOSITO CTS
          If VTipobol = "04" And wGrupoPla = "01" Then
             F01 = "Select icano+icmes+icdia From platmphisliquid"
          End If
       Case Is = "46" 'CANASTA DE NAVIDAD O SIMILARES
       Case Is = "47" 'INDEMNIZACION DESPIDO ARBITR
       Case Is = "48" 'Licencia Sindical
            F01 = "select round((b.importe/factor_horas)*a.h29,2) as basico,A.H29 from platemphist a,plaremunbase b "
            F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Trim(Txtcodpla.Text) & "' "
            F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
            F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
       Case Is = "49" 'Licencia con goce de haber
            F01 = "select round((b.importe/factor_horas)*a.h25,2) as basico,A.H25 from platemphist a,plaremunbase b "
            F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Trim(Txtcodpla.Text) & "' "
            F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
            F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
            
End Select

End Function
Public Function F02(concepto As String, QtaGra As Boolean) As String 'DEDUCCIONES
Dim rsF02 As ADODB.Recordset
Dim rsImporte As ADODB.Recordset
Dim rsF02afp As ADODB.Recordset
Dim F02str As String
Dim rsTope As ADODB.Recordset
Dim mFactor As Currency
Dim mperiodoafp As String
Dim vNombField As String
Dim mtope As Currency
Dim mincremento As Currency
Dim difmincremento As Currency
Dim ComisionMes As Currency
Dim ComisionProy As Currency
Dim mproy As Currency
Dim MUIT As Currency
Dim mgra As Integer
Dim msemano As Integer
Dim mpertope As Integer
Dim J As Integer
Dim conceptosremu As String
Dim snmanual As Byte, snfijo As Byte
Dim cptoincrementos As String
Dim IMPORTEESVMES As Currency
Dim Porc_dEsSalud As Double
Dim Porc_EPS As Double
Dim Afecto_EPS As Boolean
Dim mExtraOrd As Currency
Dim mbol As String
Dim UltimaSemana As String
Dim IsAfp As Boolean
Dim pNroSemana As Integer
Dim IsExcluirPagoSindicato As Boolean
Dim vSemVaca As String
Dim ValincrementaMesutilidad As Integer
Dim IngresoOtrasEmpresas As String
Dim IngresoComisionesVtas As String
ValincrementaMesutilidad = 0


IsExcluirPagoSindicato = False
IsAfp = False
mbol = VTipobol

'Fondo CJMMS solo Compañia con TipoTrabSunat="S" y Codigo de tipo de trabajdor SUNAT="37"
'MGL CORRIGE PARA EL NUEVO DESCUENTO DE COMEDOR 23012025
'If concepto = "14" Or concepto = "15" Or concepto = "17" Then
If concepto = "14" Or concepto = "15" Then
   If Trim(LblCodTrabSunat.Caption & "") <> "37" Then
      F02 = "select 0"
      Exit Function
    'MGL CORRIGE PARA EL NUEVO DESCUENTO DE COMEDOR 23012025
   'Or concepto = "17"
   ElseIf (concepto = "15") And Trim(Lblcodafp.Caption) = "01" Then
      F02 = "select 0"
      Exit Function
   Else
      Dim TipoTrabSunat As String
      Dim RqTS As ADODB.Recordset
      Sql$ = "select TipoTrabSunat from cia where cod_cia='" & wcia & "' and status<>'*'"
      If fAbrRst(RqTS, Sql$) Then TipoTrabSunat = Trim(RqTS(0) & "")
      RqTS.Close: Set RqTS = Nothing
      If TipoTrabSunat <> "S" Then
         F02 = "select 0"
         Exit Function
      End If
   End If
End If

If VTipobol = "04" Then mbol = "01"
mFactor = 0
F02 = ""
mtope = 0

Porc_dEsSalud = 0
Porc_EPS = 0
Afecto_EPS = False

'If lSubsidio = "S" And wGrupoPla = "01" Then 'Se calcula solo para AFP cuando es subsidio para grupo Gallos
If lSubsidio = "S" Then 'Se calcula solo para AFP cuando es subsidio
   If concepto <> "11" And concepto <> "04" And concepto <> "09" And concepto <> "06" And concepto <> "07" Then F02 = "Select 0": Exit Function
End If
'Utilidades y Cesado no se calcula EssaludVida
If (concepto = "06" And VTipobol = "11") Or (concepto = "06" And Es_Cesado) Then
   F02 = " SELECT " & 0
   Exit Function
End If
'CONCEPTO DE DEDUCCION  TARDANZAS CARGADO EL 25/07/2017 MGIRAO
If (concepto = "10") Then
   'F02 = "select round((b.importe/factor_horas)*(a.h29 *(1.000/60.000)),2) as basico,A.H29 from platemphist a,plaremunbase b "
   'F02 = F02 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Trim(Txtcodpla.Text) & "' "
   'F02 = F02 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
   'F02 = F02 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
  
   'F02 = "Update platemphist set tottardanza = (select SUM(round((b.importe/factor_horas)*(a.h19 *(1/60.000)),2)) from platemphist a,plaremunbase b "
   F02 = "Update platemphist set tottardanza = h19 * (select sum(round((b.importe/b.factor_horas),2)) from plaremunbase b where b.placod ='" & Trim(Txtcodpla.Text) & "' and b.concepto in ('01','02') and b.status<>'*' and b.cia='01') "
   ' F02 = "Update platemphist set tottardanza = (select SUM(round((b.importe/factor_horas)*(a.h19 *(1/60.000))),2)) from platemphist a,plaremunbase b "
   'F02 = F02 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Trim(Txtcodpla.Text) & "' "
   'F02 = F02 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
   'F02 = F02 & "and b.cia=a.cia and b.placod=a.placod and b.concepto IN ('01','02') and b.status<>'*')"
   'F02 = F02 & " where cia='" & wcia & "' and Status='" & Mstatus & "' and placod='" & Trim(Txtcodpla.Text) & "' "
   F02 = F02 & " Where cia='" & wcia & "' and status<>'*' and month(fechaproceso)=" & Vmes & " and year(fechaproceso)=" & Vano & " and placod='" & Trim(Txtcodpla.Text) & "' and tipotrab='01'"
   'F02 = F02 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
   cn.Execute F02
   F02 = " SELECT " & 0

   Exit Function
End If

If concepto = "04" And Trim(Lblcodafp.Caption) <> "01" And Trim(Lblcodafp.Caption) <> "02" And Trim(Lblcodafp.Caption) <> "" Then
   'Con AFP revisar si se calculo SNP en boletas enteriores del mismo mes para extornar
   F02 = "select 0"
   Sql$ = "select sum(d04*-1) from plahistorico where cia='" & wcia & "' and year(fechaproceso)=" & Vano & "  and month(fechaproceso)=" & Vmes & "  and status<>'*' and placod='" & Txtcodpla.Text & "' "
   If (fAbrRst(rsF02, Sql$)) Then
       If Trim(rsF02(0) & "") <> "" Then F02 = "select " & Trim(Str(rsF02(0)))
   End If
   rsF02.Close: Set rsF02 = Nothing
ElseIf (concepto <> "04" Or Trim(Lblcodafp.Caption) = "01" Or Trim(Lblcodafp.Caption) = "" Or Trim(Lblcodafp.Caption) = "02") And concepto <> "11" And concepto <> "13" And concepto <> "20" Then  'SIN AFP
    If Not IsDate(VFechaJub) Then
        Sql$ = "select deduccion,adicional,status from placonstante where cia='" & wcia & "' and tipomovimiento='03' and codinterno='" & concepto & "' and deduccion<>0 and status<>'*'"
    If (fAbrRst(rsF02, Sql$)) Then
        If Not IsNull(rsF02!deduccion) Then
            If rsF02!deduccion <> 0 Then mFactor = rsF02!deduccion: snmanual = IIf(rsF02!adicional = "S", 0, 1): snfijo = IIf(rsF02!status = "F", 1, 0)
        End If
    End If
    
    If rsF02.State = 1 Then rsF02.Close
    
    If (snfijo = 1 And snmanual = 0 And ((sn_essaludvida = 1 And concepto = "06")) Or (sn_sindicato = 1 And concepto = "08")) Then
        Call Acumula_Mes(concepto, "D")
        If (wtipodoc = True Or wGrupoPla = "01") And Not (sn_sindicato = 1 And concepto = "08") Then 'Quincena se calculoa como boleta normal para grupo GALLOS 240 horas
            If macui = mFactor Then
                F02 = " SELECT " & 0
            Else
                F02 = " SELECT " & mFactor
            End If
        Else
            If sn_sindicato = 1 And concepto = "08" And VTipobol = "01" Then
                Sql$ = "usp_Pla_ConsultarDatosSindicato '" & wcia & "','02','" & VTipobol & "','" & vTipoTra
                Sql$ = Sql$ & "','" & Format(FrmCabezaBol.Cmbfecha.Value, "mm/dd/yyyy") & "','" & VSemana
                Sql$ = Sql$ & "','" & Trim(Txtcodpla.Text) & "'"
                If (fAbrRst(rsF02, Sql$)) Then
                    IsExcluirPagoSindicato = True
                End If
                    
                If IsExcluirPagoSindicato = False Then
                    If vTipoTra = "02" Then
                        pNroSemana = 0
                        'Sql$ = "select  count(*)"
                        Sql$ = "select  MAX(semana)"
                        Sql$ = Sql$ & " From plasemanas"
                        Sql$ = Sql$ & " where   cia='" & wcia & "' and year(fechaf)=" & Year(FrmCabezaBol.Cmbfecha) & " and"
                        Sql$ = Sql$ & " month(fechaf)=" & Month(FrmCabezaBol.Cmbfecha) & " and status !='*'"
                        If (fAbrRst(rsF02, Sql$)) Then
                            pNroSemana = rsF02(0)
                            rsF02.Close
                        End If
                        If pNroSemana > 0 Then
                           If Val(VSemana) = pNroSemana Then
                              Dim lImpTotSind As Double
                              lImpTotSind = 0
                              Sql$ = "select SUM(d08) from plahistorico where cia='" & wcia & "' and YEAR(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " and placod='" & Trim(Txtcodpla.Text) & "' and proceso='01' and status<>'*'"
                              If (fAbrRst(rsF02, Sql$)) Then If Not IsNull(rsF02(0)) Then lImpTotSind = rsF02(0)
                              rsF02.Close
                              'F02 = "select " & (sueldominimo * mFactor / 100) - lImpTotSind
                              F02 = "select " & Round((sueldominimo * mFactor / 100)) - lImpTotSind
                              
                           Else
                              'F02 = "select " & (sueldominimo * mFactor / 100) / pNroSemana
                              'F02 = "select " & (sueldominimo * mFactor / 100) / 4
                              'ADD MP 290323
                              F02 = "select " & Round((sueldominimo * mFactor / 100)) / 4
                           End If
                        Else
                            F02 = " SELECT " & 0
                        End If
                    Else
                        F02 = "select " & sueldominimo * mFactor / 100
                    End If
                Else
                    F02 = " SELECT " & 0
                End If
            Else
                F02 = " SELECT " & 0
            End If
        End If
        Exit Function
    End If
    
    If mFactor <> 0 Then
       Sql$ = "select cod_remu from plaafectos where cia='" & wcia & "' and tipo='D' and tboleta='" & mbol & "'  and  codigo='" & concepto & "' and status<>'*'"
       If (fAbrRst(rsF02, Sql$)) Then
          rsF02.MoveFirst
          F02str = ""
          Do While Not rsF02.EOF
             F02str = F02str & "i" & Trim(rsF02!cod_remu) & "+"
             rsF02.MoveNext
          Loop
          F02str = "(" & Mid(F02str, 1, Len(Trim(F02str)) - 1) & ")"
          If rsF02.State = 1 Then rsF02.Close
          Call Acumula_Mes(concepto, "D")
          F02 = "select round(((" & F02str & " + " & macui & ") * " & mFactor & " /100)-" & macus & ",2) as deduccion from platemphist "
          F02 = F02 & "where cia='" & wcia & "' and status='" & Mstatus & "' and placod='" & Txtcodpla.Text & "' "
          F02 = F02 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & ""
       End If
    End If
   End If
ElseIf Trim(concepto) = "20" Then 'SOLICITUD DE SINDICATO
   If VTipobol = "02" Or VTipobol = "03" Or VTipobol = "11" Then
     F02 = "select 0"
   Else
      Dim VSemanaSindicato As String
      VSemanaSindicato = VSemana
      If Trim(vTipoTra) = "01" Then
          VSemanaSindicato = Vmes
      End If
      Sql$ = "usp_Pla_ConsultarImportePlaSolicitudSindicato '" & wcia & "','" & Trim(VTipobol)
      Sql$ = Sql$ & "','" & Trim(vTipoTra) & "','" & Trim(Txtcodpla.Text) & "',"
      Sql$ = Sql$ & Val(Mid(VFProceso, 7, 4)) & ",'" & VSemanaSindicato & "'"
      If (fAbrRst(rsF02, Sql$)) Then
          F02 = "select " & CStr(rsF02!importe)
        
'        Sql$ = "Update PlaSolicitudSindicatoDetalle"
'        Sql$ = Sql$ & " set PagoAcuen=" & CStr(rsF02!importe)
'        Sql$ = Sql$ & " where id=" & CStr(rsF02!Id)
'        Sql$ = Sql$ & " and placod= '" & Trim(Txtcodpla.Text) & "' and status !='*'"
'        cn.Execute Sql$
      Else
        F02 = "select 0"
      End If
   End If
   Exit Function
ElseIf Trim(concepto) = "11" And Trim(Lblcodafp.Caption) <> "" Then 'AFP
   Sql$ = "select * from maestros_2 where ciamaestro='01069' and cod_maestro2='" & Trim(Lblcodafp.Caption) & "' and flag8='S' and status<>'*'"
   If (fAbrRst(rsF02, Sql$)) Then IsAfp = True
   
   If Not ((IsDate(VFechaJub) Or VJubilado = "S") And IsAfp = True) Then
   If Trim(Lblcodafp.Caption) = "01" Or Trim(Lblcodafp.Caption) = "02" Then GoTo AFP
    Sql$ = "select cod_remu from plaafectos where cia='" & wcia & "' and tipo='D' and tboleta='" & mbol & "'  and  codigo='" & concepto & "' and status<>'*'"

    If (fAbrRst(rsF02, Sql$)) Then
       rsF02.MoveFirst
       F02str = ""
       Do While Not rsF02.EOF
          F02str = F02str & "i" & Trim(rsF02!cod_remu) & "+"
          
          rsF02.MoveNext
       Loop
       
       F02str = "(" & Mid(F02str, 1, Len(Trim(F02str)) - 1) & ")"
       
       If rsF02.State = 1 Then rsF02.Close
       mperiodoafp = Format(Vano, "0000") & Format(Vmes, "00")
       Sql$ = "select afp01,afp02,afp03,afp04,afp05,tope from  plaafp where periodo='" & mperiodoafp & "' and codafp='" & Lblcodafp.Caption & "' and status<>'*' and cia='" & wcia & "'"
    
       If Not (fAbrRst(rsF02, Sql$)) Then
          MsgBox "No se Encuentran Factores de Calculo para AFP", vbCritical, "Calculo de Boleta"
          mcancel = True
          Exit Function
       End If
       Sql$ = Acumula_Mes_Afp(concepto, "D")
       If (fAbrRst(rsF02afp, Sql$)) Then
          For J = 1 To 5
              vNombField = " as D11" & Format(J, "0")
              
              mFactor = rsF02(J - 1)
              
              If J = 4 Then
                    If lblAfpTipoComision.Caption = "M" Then mFactor = 0#
              ElseIf J = 5 Then
                    If lblAfpTipoComision.Caption = "F" Then mFactor = 0#
              End If
                            
              If J = 2 Then
                 If manos > 64 And IsAfp = True Then mFactor = 0
                 Call Acumula_Mes_Afp112(concepto, "D")
                 mtope = macui
                 
                 Sql$ = "select " & F02str & " as tope from platemphist where cia='" & wcia & "' and status='" & Mstatus & "' and placod='" & Txtcodpla.Text & "' " _
                      & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & ""
                      
                 If (fAbrRst(rsTope, Sql$)) Then mtope = mtope + rsTope!tope
                 
                 If wtipodoc = False Then
                    If wGrupoPla <> "01" Then 'En quincena Grupo GALLOS trabaja con 240 hora como una boleta normal
                       mtope = mtope + rsTope!tope
                    End If
                 End If
                 
                 If mtope > rsF02!tope Then mtope = rsF02!tope
                 If rsTope.State = 1 Then rsTope.Close
                 
                 'If wtipodoc = False Then
                 If wtipodoc = False And wGrupoPla <> "01" Then 'En quincena Grupo GALLOS trabaja con 240 hora como una boleta normal
                    F02 = F02 & "round((((" & mtope & ") * " & mFactor & " /100)-" & macus & ")/2,2) "
                 Else
                    F02 = F02 & "round(((" & mtope & ") * " & mFactor & " /100)-" & macus & ",2) "
                 End If
                 F02 = F02 & vNombField & ","
              Else
                 If Not IsNull(rsF02afp(0)) Then
                    'If wtipodoc  Then
                    If wtipodoc Or wGrupoPla = "01" Then 'En quincena Grupo GALLOS trabaja con 240 hora como una boleta normal
                       F02 = F02 & "round(((" & F02str & " + " & rsF02afp(0) & ") * " & mFactor & " /100)-" & rsF02afp(J) & ",2) "
                       F02 = F02 & vNombField & ","
                    Else
                       F02 = F02 & "round((((" & F02str & " + " & rsF02afp(0) & ") * " & mFactor & " /100)-" & rsF02afp(J) & "),2) "
                       F02 = F02 & vNombField & ","
                    End If
                 Else
                    'If wtipodoc Then
                    If wtipodoc Or wGrupoPla = "01" Then 'En quincena Grupo GALLOS trabaja con 240 hora como una boleta normal
                        F02 = F02 & "round(((" & F02str & " ) * " & mFactor & " /100),2) "
                        F02 = F02 & vNombField & ","
                    Else
                        F02 = F02 & "round(((" & F02str & " ) * " & mFactor & " /100),2) "
                        F02 = F02 & vNombField & ","
                    End If
                 End If
              End If
          Next J
          If rsF02afp.State = 1 Then rsF02afp.Close
          If rsF02.State = 1 Then rsF02.Close
       End If
       F02 = Mid(F02, 1, Len(Trim(F02)) - 1)
       F02 = "select " & F02
       F02 = F02 & " from platemphist "
       F02 = F02 & "where cia='" & wcia & "' and status='" & Mstatus & "' and placod='" & Txtcodpla.Text & "' "
       F02 = F02 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & ""
AFP:
    End If
   End If
'de 13 lo cambie a 131 para que no entre
'mario
'ElseIf concepto = "13" And (VTipobol <> "03" Or QtaGra) Then 'Quinta Categoria
ElseIf concepto = "13" And (VTipobol <> "03" Or QtaGra) Then 'Quinta Categoria

   IngresoOtrasEmpresas = "N"
   IngresoComisionesVtas = "N"
   Sql$ = "select Placod from Pla_Trab_Otras_Empresas where cia='" & wcia & "' and placod='" & Txtcodpla.Text & "' and status<>'*'"
   If (fAbrRst(rsF02, Sql$)) Then IngresoOtrasEmpresas = "S"
   rsF02.Close: Set rsF02 = Nothing
   Sql$ = ""
   


   ValincrementaMesutilidad = 0
   If VTipobol = "11" Then
      ValincrementaMesutilidad = 1
      'Buscamos boleta de vacaciones o normal en el mes que se genera la utilidad para determinar el proyectado
      If vTipoTra = "01" Then
         Sql$ = "select Top 1  Placod from plahistorico where cia='" & wcia & "' and year(fechaproceso)=" & Vano & "  and month(fechaproceso)=" & Vmes & "  and status<>'*' and placod='" & Txtcodpla.Text & "' and proceso in('01','02')"
         If (fAbrRst(rsF02, Sql$)) Then ValincrementaMesutilidad = 0
         rsF02.Close: Set rsF02 = Nothing
       End If
   End If
   
    'PREGUNTAMOS SI ESTA AFECTO O NO
    
    If VTipobol = "04" And Not lCalculaLiqui Then GoTo quinta
    
    If VTipobol = "05" Then GoTo quinta
    If Not sn_quinta Then GoTo quinta
    If VTipobol = "10" Then GoTo quinta
    If Verifica_Subsidio Then GoTo quinta
    
    Dim Dif_Meses As Integer
    Dim mProyGrati As Double
    Dim mProyProm As Double
    Dim xTempProy As Double
    Dim xMTipBolNor As String
    xMTipBolNor = VTipobol
    If VTipobol = "11" Then xMTipBolNor = "01"
    If VTipobol = "04" Then xMTipBolNor = "01"
    
    vSemVaca = VSemana
    Dim rsAuxiliar As ADODB.Recordset
    'Buscar Hasta que mes se debe proyectar los promedios
    Dim lMesTopeProyProm As Integer
    lMesTopeProyProm = 0
    Sql$ = "Select Mes_Tope From Pla_Qta_Mes_Tope_Proyecc where cia='01' and Status<>'*'"
    Set rsAuxiliar = OpenRecordset(Sql$, cn)
    If Not rsAuxiliar.EOF Then lMesTopeProyProm = rsAuxiliar(0) Else lMesTopeProyProm = 12
    rsAuxiliar.Close: Set rsAuxiliar = Nothing
    
    Sql$ = "select fingreso from planillas where cia = '" & wcia & "' and placod = '" & Txtcodpla.Text & "' and status != '*'"
    Set rsAuxiliar = OpenRecordset(Sql$, cn)
    If Not rsAuxiliar.EOF Then
        Dif_Meses = DateDiff("m", rsAuxiliar!fIngreso, CDate("31/12/" & Year(FrmCabezaBol.Cmbfecha.Value)))
        If Day(rsAuxiliar!fIngreso) = 1 Then Dif_Meses = Dif_Meses + 1
    End If
    If rsAuxiliar.State = 1 Then rsAuxiliar.Close
    
    Sql$ = "select cod_remu from plaafectos where cia='" & wcia & "' and tipo='D' and tboleta='" & xMTipBolNor & "'  and  codigo='" & concepto & "' and status<>'*'"
   
   If (fAbrRst(rsF02, Sql$)) Then
      rsF02.MoveFirst
      F02str = ""
      conceptosremu = ""
      cptoincrementos = ""
      Do While Not rsF02.EOF
         F02str = F02str & "i" & Trim(rsF02!cod_remu) & "+"
         rsF02.MoveNext
      Loop

      F02str = "(" & Mid(F02str, 1, Len(Trim(F02str)) - 1) & ")"
      If rsF02.State = 1 Then rsF02.Close
      MUIT = 0
      mtope = 0
      mpertope = 0
      mgra = 0
      msemano = 0
      mincremento = 0
      If vTipoTra <> "01" Then
         Sql$ = "select isnull(max(semana),0) as ultima from plasemanas where cia='" & wcia & "' and ano='" & Format(Vano, "0000") & "' and status<>'*'"
         If (fAbrRst(rsF02, Sql$)) Then mpertope = rsF02(0): msemano = rsF02(0)
      Else
         If Vmes > 6 Then mpertope = 12 Else mpertope = 13
      End If
      If rsF02.State = 1 Then rsF02.Close
      
      'ACUMULADO DE TODOS INGRESOS
      Call Acumula_Ano(concepto, "D")
      
      
      'Remuneraciones del mes de otras empresas
      If IngresoOtrasEmpresas = "S" Then
         Sql$ = "select isnull(SUM(ingreso),0) as Basico from Pla_Trab_Otras_Empresas_Meses where Cia='" & wcia & "' AND Ayo=" & Vano & " and Mes<=" & Vmes & " AND placod='" & Txtcodpla.Text & "'  and Status<>'*'"
         If (fAbrRst(rsF02, Sql$)) Then macui = macui + rsF02(0)
         If rsF02.State = 1 Then rsF02.Close
      End If
      
      mtope = macui
      
'      cptoincrementos = F02str
      
      'OBTENER EL INCREMENTO
      If Len(Trim(cptoincrementos)) > 0 Then
        Sql$ = "select " & cptoincrementos & " as tope from platemphist where cia='" & wcia & "' and status='" & Mstatus & "' and placod='" & Txtcodpla.Text & "' " _
             & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & ""
        If (fAbrRst(rsF02, Sql$)) Then mincremento = rsF02!tope
        If rsF02.State = 1 Then rsF02.Close
      End If
      
      Sql$ = "select " & F02str & " as tope from platemphist where cia='" & wcia & "' and status='" & Mstatus & "' and placod='" & Txtcodpla.Text & "' " _
           & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & ""
      If (fAbrRst(rsF02, Sql$)) Then mtope = mtope + rsF02!tope
      
      'Si trabajador es no domicliado se calcula el 30% de quinta categoria en cada boleta
      If Not sn_Domiciliado Then
         F02 = "Select Round(" & rsF02!tope & " * 0.3, 2)"
         If rsF02.State = 1 Then rsF02.Close
         Exit Function
      End If
      
      If rsF02.State = 1 Then rsF02.Close
      
      'Ingreso Comisiones Vendedor del mes Mgirao 30/10/2016
      Dim F02Comision As String
      ComisionMes = 0: F02Comision = ""
      ComisionProy = 0
     'select *,codinterno as cod_remu from placonstante where cia='01' and tipomovimiento='02' and promqta='S' and status<>'*' and calculo='N' AND basico='S'
      Sql$ = "select codinterno as cod_remu from placonstante where cia='" & wcia & "' and tipomovimiento='02' and promqta='S' and calculo='N' and status<>'*'"
      If (fAbrRst(rsF02, Sql$)) Then rsF02.MoveFirst
      Do While Not rsF02.EOF
            F02Comision = F02Comision & "i" & Trim(rsF02!cod_remu) & "+"
            rsF02.MoveNext
      Loop
      F02Comision = "(" & Mid(F02Comision, 1, Len(Trim(F02Comision)) - 1) & ")"
      If rsF02.State = 1 Then rsF02.Close
      
      Sql$ = "select " & F02Comision & " as Comision from platemphist where cia='" & wcia & "' and status='" & Mstatus & "' and placod='" & Txtcodpla.Text & "' " _
           & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & ""
           
      If (fAbrRst(rsF02, Sql$)) Then ComisionMes = rsF02!Comision
      ComisionProy = ((ComisionMes + macomi) / 2) * (12 - (Vmes))
      
      
      If rsF02.State = 1 Then rsF02.Close
      
      'Ingresos Extraordinarios del mes
      Dim F02ExtOrd As String
      mExtraOrd = 0: F02ExtOrd = ""
      Sql$ = "select codinterno as cod_remu from placonstante where cia='" & wcia & "' and tipomovimiento='02' and extraord='S' and status<>'*'"
      If (fAbrRst(rsF02, Sql$)) Then rsF02.MoveFirst
      Do While Not rsF02.EOF
            F02ExtOrd = F02ExtOrd & "i" & Trim(rsF02!cod_remu) & "+"
            rsF02.MoveNext
      Loop
      F02ExtOrd = "(" & Mid(F02ExtOrd, 1, Len(Trim(F02ExtOrd)) - 1) & ")"
      If rsF02.State = 1 Then rsF02.Close
      
      Sql$ = "select " & F02ExtOrd & " as ExtraOrd from platemphist where cia='" & wcia & "' and status='" & Mstatus & "' and placod='" & Txtcodpla.Text & "' " _
           & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & ""
           
      If (fAbrRst(rsF02, Sql$)) Then mExtraOrd = rsF02!extraord
      
      'Monto Extraordinario en plahistorico del mes
      If vTipoTra = "01" Then
         Sql$ = "select " & F02ExtOrd & " as ExtraOrd from plahistorico where cia='" & wcia & "' and placod='" & Txtcodpla.Text & "' " _
             & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " and proceso in('01','02','11') and status<>'*' "
           
         If (fAbrRst(rsF02, Sql$)) Then mExtraOrd = mExtraOrd + rsF02!extraord
         If rsF02.State = 1 Then rsF02.Close
      End If
      'Fin Ingresos Extraordinarios
      
        Sql$ = "select concepto,moneda,sum((importe/factor_horas)) as base, SUM(importe) as importe from plaremunbase a,plaafectos b where a.cia='" & wcia & "' " _
           & "and placod='" & Txtcodpla.Text & "' and a.status<>'*' and b.tipo='D' AND B.CIA='" & wcia & "' and a.concepto<>'08' and b.codigo='" & concepto & "' and b.tboleta='" & xMTipBolNor & "' and b.status<>'*' and a.concepto=b.cod_remu " _
           & "Group By Placod,a.concepto,a.moneda"
      
      If (fAbrRst(rsF02, Sql$)) Then
      importe = 0
        Do While Not rsF02.EOF
            mproy = mproy + rsF02!base
            importe = importe + rsF02!importe
            rsF02.MoveNext
        Loop
        'mproy = Round(mproy, 2)
        mProyGrati = Round(importe, 2)
      End If
      
      'Agregamos Basico de otras Empresas
      Dim BasicoOtras As Double
      BasicoOtras = 0
      If IngresoOtrasEmpresas = "S" Then
         Sql$ = "select isnull(SUM(Basico),0) as Basico from Pla_Trab_Otras_Empresas_Meses where Cia='" & wcia & "' AND Ayo=" & Vano & " and Mes=" & Vmes & " AND placod='" & Txtcodpla.Text & "' and Status<>'*'"
         If (fAbrRst(rsF02, Sql$)) Then
            mproy = mproy + rsF02!Basico / 240
            'importe = importe + rsF02!Basico
            mProyGrati = mProyGrati + rsF02!Basico
            BasicoOtras = rsF02!Basico
         End If
         If rsF02.State = 1 Then rsF02.Close
         Sql$ = ""
      End If
      'Fin Otras Empresas
      
      
      'Promedios para proyectar
      
      If Vmes < 12 And Vmes <= lMesTopeProyProm Then mProyProm = Calcula_Promedios_Qta
      If mProyProm > 0 Then mProyProm = Round(mProyProm / 240, 2)
      'Saca Importe total de remuneraciones
      
      If rsF02.State = 1 Then rsF02.Close
      
      Sql$ = "select uit from plauit where ano='" & Format(Vano, "0000") & "' and moneda='S/.' and status<>'*'"
      If (fAbrRst(rsF02, Sql$)) Then MUIT = rsF02!uit
      If rsF02.State = 1 Then rsF02.Close
      
      If vTipoTra = "05" Then
         Sql$ = "select factor from platasaanexo where modulo='02' and tipomovimiento='" & Trim(Lbltope.Caption) & "' and codinterno='20' and tipotrab='" & vTipoTra & "' and cargo='" & Trim(Lblcargo.Caption) & "' and status<>'*'"
         If (fAbrRst(rsF02, Sql$)) Then mFactor = rsF02!factor
         If rsF02.State = 1 Then rsF02.Close
      
         Sql$ = "select concepto,moneda,importe/factor_horas as base from plaremunbase where cia='" & wcia & "' and placod='" & Txtcodpla.Text & "' and concepto='01' and status<>'*'"
         If (fAbrRst(rsF02, Sql$)) Then mproy = mproy + Round(rsF02!base * mFactor, 2)
         If rsF02.State = 1 Then rsF02.Close
      End If
          
          
    'PARA EL CALCULO DE EPS
    Dim rsTemporal As ADODB.Recordset
    Cadena = "SELECT APORTACION FROM PLACONSTANTE WHERE CIA = '" & wcia & "' AND TIPOMOVIMIENTO = '03' AND STATUS != '*' AND CODINTERNO = '01'"
    Set rsTemporal = OpenRecordset(Cadena, cn)
    If Not rsTemporal.EOF Then
        Porc_dEsSalud = Val(rsTemporal!aportacion)
    Else
        Porc_dEsSalud = 0
    End If
    If rsTemporal.State = adStateOpen Then rsTemporal.Close
    
    'Cadena = "SELECT ISNULL(APORTACION_EPS,0) AS APORTACION_EPS FROM CIA WHERE COD_CIA = '" & wcia & "' AND STATUS != '*'"
    Cadena = "select codigo_eps from planillas where Cia='" & wcia & "' and placod='" & Trim(Txtcodpla.Text) & "' and status<>'*'"
    Set rsTemporal = OpenRecordset(Cadena, cn)
    If Not rsTemporal.EOF Then
        'Afecto_EPS = CBool(rsTemporal!Aportacion_EPS)
        If Trim(rsTemporal!codigo_eps & "") <> "" Then Afecto_EPS = True
    Else
        Afecto_EPS = False
    End If
    If rsTemporal.State = adStateOpen Then rsTemporal.Close
    If Afecto_EPS Then
        Cadena = "SELECT ISNULL(CODIGO_EPS,'') AS CODIGO_EPS FROM PLANILLAS WHERE CIA = '" & wcia & "' AND PLACOD = '" & Trim(Txtcodpla.Text) & "' AND STATUS != '*' AND TIPOTRABAJADOR = '" & vTipoTra & "'"
        Set rsTemporal = OpenRecordset(Cadena, cn)
        If Not rsTemporal.EOF Then
            If rsTemporal!codigo_eps <> "" Then
                Cadena = "SELECT ISNULL(plaIMPORTE,0) AS IMPORTE FROM MAESTROS_2 WHERE RIGHT(CIAMAESTRO,3) = '143' AND STATUS != '*' AND RTRIM(ISNULL(CODSUNAT,'')) != '' AND COD_MAESTRO2 = '" & Trim(rsTemporal!codigo_eps) & "'"
                If rsTemporal.State = adStateOpen Then rsTemporal.Close
                Set rsTemporal = OpenRecordset(Cadena, cn)
                If Not rsTemporal.EOF Then
                    Porc_EPS = Val(rsTemporal!importe)
                Else
                    Porc_EPS = 0
                End If
            Else
                Porc_EPS = 0
            End If
        Else
            Porc_EPS = 0
        End If
    End If
      'Porc_EPS = Porc_dEsSalud - Porc_EPS
      'Porc_EPS = 0: Porc_dEsSalud = 0 ' ya no se Proyecta Ingreso Extraordinario
      If vTipoTra = "01" Then
         'If wtipodoc = True Then
         If wtipodoc = True Or wGrupoPla = "01" Then 'En quincena Grupo GALLOS trabaja con 240 hora como una boleta normal
            'If Vmes < 12 Then mproy = ((mproy * VHoras) + mincremento) * (mpertope - Vmes + 1) Else mproy = 0
            'Cambio autorizado por JCB 09/07/09
            
            If Afecto_EPS Then
                If Vmes < 12 Then mproy = (((mproy + mProyProm) * VHoras) + mincremento) * (mpertope - Vmes + IIf(Dif_Meses > 5, 1, Dif_Meses / 6)) + (IIf(Vmes > 6, (IIf(Dif_Meses > 5, (importe * Porc_EPS) / 100, (((importe * Porc_EPS) / 100) * Dif_Meses / 6))), ((importe * Porc_EPS) / 100) * 2)) Else mproy = 0
            Else
                If Vmes < 12 Then
                   'mproy = (((mproy + mProyProm) * VHoras) + mincremento) * (mpertope - Vmes + IIf(Dif_Meses > 5, 1, Dif_Meses / 6)) + (IIf(Vmes > 6, (IIf(Dif_Meses > 5, (importe * Porc_dEsSalud) / 100, (((importe * Porc_dEsSalud) / 100) * Dif_Meses / 6))), ((importe * Porc_dEsSalud) / 100) * 2))
                   mproy = (((mproy) * VHoras) + mincremento) * (mpertope - (Vmes - ValincrementaMesutilidad) + IIf(Dif_Meses > 5, 1, Dif_Meses / 6)) + (IIf(Vmes > 6, (IIf(Dif_Meses > 5, (importe * Porc_dEsSalud) / 100, (((importe * Porc_dEsSalud) / 100) * Dif_Meses / 6))), ((importe * Porc_dEsSalud) / 100) * 2))
                   mproy = mproy + (mProyProm * VHoras) * (12 - (Vmes))
                Else
                   mproy = 0
                End If
            End If
            
            If IngresoOtrasEmpresas = "S" Then
               If Vmes < 12 Then
                    If Vmes < 7 Then
                       mproy = mproy + ((BasicoOtras * Porc_dEsSalud / 100) * 2)
                    Else
                       mproy = mproy + (BasicoOtras * Porc_dEsSalud / 100)
                    End If
               End If
            End If

         Else
            If Vmes = 12 Then
               mproy = (((mproy * VHoras) + difmincremento) * (mpertope - Vmes)) + Round((((mproy * VHoras) + difmincremento) / 2), 2)
            Else
              'Cambio autorizado por JCB 09/07/09
              If Afecto_EPS Then
                mproy = (((mproy * VHoras) + difmincremento) * (mpertope - Vmes + IIf(Dif_Meses > 5, 1, Dif_Meses / 6))) + Round((((mproy * VHoras) + difmincremento) / 2), 2) + (IIf(Vmes > 6, (IIf(Dif_Meses > 5, (importe * Porc_EPS) / 100, (((importe * Porc_EPS) / 100) * Dif_Meses / 6))), ((importe * Porc_EPS) / 100) * 2))
              Else
                mproy = (((mproy * VHoras) + difmincremento) * (mpertope - Vmes + IIf(Dif_Meses > 5, 1, Dif_Meses / 6))) + Round((((mproy * VHoras) + difmincremento) / 2), 2) + (IIf(Vmes > 6, (IIf(Dif_Meses > 5, (importe * Porc_dEsSalud) / 100, (((importe * Porc_dEsSalud) / 100) * Dif_Meses / 6))), ((importe * Porc_dEsSalud) / 100) * 2))
              End If
              
            End If
         End If
      Else
         mgra = Busca_Grati()
         If VTipobol = "03" Then mgra = mgra - 1
         
         If vTipoTra = "05" Then
            Sql$ = "select importe/factor_horas as base,b.factor  from plaremunbase a,platasaanexo b where a.cia='" & wcia & "' and a.placod='" & Txtcodpla.Text & "'  and a.concepto='01' " _
                 & "and a.status<>'*' and b.cia='" & wcia & "' and b.tipomovimiento='01' and b.codinterno='15' and b.status<>'*' and b.tipotrab='" & vTipoTra & "' and b.cargo='" & Trim(Lblcargo.Caption) & "'"
            If (fAbrRst(rsF02, Sql$)) Then mproy = ((mproy * VHoras) + mincremento) * (mpertope - Val(VSemana)) + ((rsF02!base * 8) * (rsF02!factor * mgra)) + ((rsF02!base * 8) * (mpertope - Val(VSemana)))
            If rsF02.State = 1 Then rsF02.Close
         Else
         
         'MODIFICADO POR Ricardo Hinostroza
         'Fecha:18/10/2008
         'Motivo:calculo de boleta
            Dim EncuentraVaca As Boolean
            EncuentraVaca = False
            If VTipobol = "03" And vTipoTra = "02" Then
               UltimaSemana = Devuelve_Ultima_Semana
            End If
            If VTipobol = "02" And Vmes = 12 Then
               'If Vmes = 12 Then mproy = 0
               mproy = 0
            Else
               If VTipobol = "02" And VSemana = "" Then 'Obtenemos la Semana de las vacaciones
                  Dim SqlFec As String
                  Dim RqVSem As ADODB.Recordset
                  vSemVaca = ""
                  SqlFec = " select semana from plasemanas where cia='" & wcia & "' and ano=" & Vano & " and status<>'*' and '" & Format(VFProceso, FormatFecha) & "' between fechai and fechaf"
                  If (fAbrRst(RqVSem, SqlFec)) Then
                     VSemana = Trim(RqVSem!semana & "")
                     VSemana = Format(CCur(VSemana + 4), "00")
                  End If
                  RqVSem.Close: Set RqVSem = Nothing
                End If
                If VTipobol = "01" And LbLSemBolVac.Caption <> "" Then
                   VSemana = Format(CCur(VSemana + 5), "00")
                End If
                
                If Vmes = 12 Then
                   'Busca Vacaciones si encuentra vacaciones no se proyectan las semanas restantes
                   Dim RqVac As ADODB.Recordset
                   Sql$ = "select placod from plahistorico where cia='" & wcia & "' and placod='" & Txtcodpla.Text & "' and proceso='02' and year(fechaproceso)=" & Mid(VFProceso, 7, 4) & " and month(fechaproceso)=12 and status<>'*'"
                   If fAbrRst(RqVac, Sql) Then mproy = 0: EncuentraVaca = True
                   RqVac.Close: Set RqVac = Nothing
                End If
                If mpertope = Val(VSemana) And vTipoTra = "02" Then
                   mproy = 0
                ElseIf EncuentraVaca = False Then
                   Sql$ = "select importe/factor_horas as base from plaremunbase where Cia='" & wcia & "' and placod='" & Txtcodpla.Text & "'  and concepto='01' and status<>'*'"
                   If (fAbrRst(rsF02, Sql$)) Then
                       If Afecto_EPS Then
                           If VTipobol = "03" And vTipoTra = "02" Then
                              mproy = ((mproy * VHoras) + mincremento) * (mpertope - Val(UltimaSemana)) + (((mproy * 240) + mincremento) * mgra) + (rsF02!base * 8) * (mpertope - Val(UltimaSemana)) + IIf(Vmes > 6 And mgra <> 0, ((mproy * 240) * Porc_Porc_EPS) / 100, (((mproy * 240) * Porc_Porc_EPS) / 100) * 2)
                           Else
                              mproy = ((mproy * VHoras) + mincremento) * (mpertope - Val(VSemana)) + (((mproy * 240) + mincremento) * mgra) + (rsF02!base * 8) * (mpertope - Val(VSemana)) + IIf(Vmes > 6 And mgra <> 0, ((mproy * 240) * Porc_Porc_EPS) / 100, (((mproy * 240) * Porc_Porc_EPS) / 100) * 2)
                           End If
                       Else
                           'mproy = ((mproy * VHoras) + mincremento) * (mpertope - Val(VSemana)) + (((mproy * 240) + mincremento) * mgra) + (rsF02!base * 8) * (mpertope - Val(VSemana)) + IIf(Vmes > 6 And mgra <> 0, ((mproy * 240) * Porc_dEsSalud) / 100, (((mproy * 240) * Porc_dEsSalud) / 100) * 2)
                           xTempProy = mproy
                           If VTipobol = "03" And vTipoTra = "02" Then
                              mproy = (((xTempProy + mProyProm) * VHoras) + mincremento) * (mpertope - Val(UltimaSemana))
                              mproy = mproy + (rsF02!base * 8) * (mpertope - Val(UltimaSemana))
                           Else
                              mproy = (((xTempProy + mProyProm) * VHoras) + mincremento) * (mpertope - Val(VSemana))
                              mproy = mproy + (rsF02!base * 8) * (mpertope - Val(VSemana))
                           End If
                           mproy = mproy + (((xTempProy * 240) + mincremento) * mgra)
                           mproy = mproy + IIf(Vmes > 6 And mgra <> 0, ((xTempProy * 240) * Porc_dEsSalud) / 100, (((xTempProy * 240) * Porc_dEsSalud) / 100) * 2)
                       End If
                   End If
                   If rsF02.State = 1 Then rsF02.Close
                End If
            End If
            'mproy = (mproy * VHoras) * (mpertope - Val(VSemana)) + ((mproy * 240) * mgra) + (mproy * 8) * (mpertope - Val(VSemana))
         End If
      End If
      If VTipobol = "11" And Es_Cesado Then mproy = 0 'BOLETA DE UTILIDADES Y CESADO, NO SE PROYECTA
      If VTipobol = "04" Then mproy = 0 'BOLETA DE LIQUIDACION, NO SE PROYECTA
      
      'anula el calculo de quinta hasta que supere el mes 3 para obreros
     ' If vTipoTra = "02" And VTipobol = "01" And Vmes <= 3 Then mproy = 0

      mtope = mtope + mproy + ComisionProy
      Dim mCalExtOrd As Currency
      mCalExtOrd = 0
      Dim mtope2 As Currency
      If QtaGra And VTipobol = "03" And vTipoTra = "01" Then
         mtope = mtope + mProyGrati
      End If
      Dim mFacUti As Currency
      mFacUti = 0
      
      'mgirao
      ' Determino Importe de quinta diferencia vs proyectada
         Dim QtaDif As Double
         Dim QuintaMes As Double
         QtaDif = 0
         QuintaMes = 0
          If VTipobol = "01" And vTipoTra = "02" Then
            'Sql$ = "exec PLANSS_RETEQTA_ACUMULADA '" & wcia & "'," & Vano & ",'" & vTipoTra & "','" & Txtcodpla.Text & "'"
            'mgirao para no proyectar - comentar cuando no se use
            '-------
            Sql$ = "exec PLANSS_RETEQTA_ACUMULADA_V2 '" & wcia & "'," & Vano & ",'" & vTipoTra & "','" & Txtcodpla.Text & "'," & mproy & ",'" & VSemana & "'"
            'Sql$ = "exec PLANSS_RETEQTA_ACUMULADA_V2 '" & wcia & "'," & Vano & ",'" & vTipoTra & "','" & Txtcodpla.Text & "'," & 0
            If (fAbrRst(rsF02, Sql$)) Then QtaDif = rsF02(13) + (TotalIngresos * 0.08)
            'If QtaDif > 0 Then
            '   QtaDif = QtaDif * -1
            'End If
            If rsF02.State = 1 Then rsF02.Close
            'Sql$ = "select sum(d13) from plahistorico  where cia= '" & wcia & "' and status<>'*' and placod = '" & Txtcodpla.Text & "' and YEAR(fechaproceso)=" & Vano & " and MONTH(fechaproceso)='08'"
            Sql$ = "select sum(d13) from plahistorico  where cia= '" & wcia & "' and status<>'*' and placod = '" & Txtcodpla.Text & "' and YEAR(fechaproceso)=" & Vano & " and MONTH(fechaproceso)='" & Vmes & "'"
            If (fAbrRst(rsF02, Sql$)) Then If IsNull(rsF02(0)) Then QuintaMes = 0 Else QuintaMes = rsF02(0)
            If rsF02.State = 1 Then rsF02.Close
          End If
      
      If mtope > Round(MUIT * 7, 2) Then
         mtope2 = mtope
         mtope = mtope - Round(MUIT * 7, 2)
         If Vano > 2014 Then
            Select Case mtope
              'mgirao
                Case Is < (Round(MUIT * 5, 2) + 1)
                     If (mtope2 - mExtraOrd) < Round(MUIT * 7, 2) Then
                        mFactor = 0
                        mCalExtOrd = Round((mtope2 - Round(MUIT * 7, 2)) * 8 / 100, 2)
                        If QtaDif < 0 And mCalExtOrd = 0 Then mCalExtOrd = QtaDif
                        mFacUti = 0.08
                     Else
                        mFactor = Round((mtope - mExtraOrd) * 0.08, 2)
                        mCalExtOrd = Round(mtope * 0.08, 2) - mFactor
                        If QtaDif < 0 And mCalExtOrd = 0 Then mCalExtOrd = QtaDif
                        mFacUti = 0.08
                     End If
                Case Is < (Round(MUIT * 20, 2) + 1)
                     mFactor = Round((((mtope - mExtraOrd) - (MUIT * 5)) * 0.14) + (MUIT * 5) * 0.08, 2)
                     mCalExtOrd = Round(((mtope - (MUIT * 5)) * 0.14) + (MUIT * 5) * 0.08, 2) - mFactor
                     mFacUti = 0.14
                Case Is < (Round(MUIT * 35, 2) + 1)
                     mFactor = Round((((mtope - mExtraOrd) - (MUIT * 20)) * 0.17) + ((MUIT * 20) - (MUIT * 5)) * 0.14, 2) + ((MUIT * 5) * 0.08)
                     mCalExtOrd = Round(((mtope - (MUIT * 20)) * 0.17) + ((MUIT * 20) - (MUIT * 5)) * 0.14, 2) + ((MUIT * 5) * 0.08) - mFactor
                     mFacUti = 0.17
                Case Is < (Round(MUIT * 45, 2) + 1)
                     mFactor = Round((((mtope - mExtraOrd) - (MUIT * 35)) * 0.2) + ((MUIT * 35) - (MUIT * 20)) * 0.17 + ((MUIT * 20) - (MUIT * 5)) * 0.14, 2) + ((MUIT * 5) * 0.08)
                     mCalExtOrd = Round(((mtope - (MUIT * 35)) * 0.2) + ((MUIT * 35) - (MUIT * 20)) * 0.17 + ((MUIT * 20) - (MUIT * 5)) * 0.14, 2) + ((MUIT * 5) * 0.08) - mFactor
                     mFacUti = 0.2
                Case Else
                     mFactor = Round((((mtope - mExtraOrd) - (MUIT * 45)) * 0.3) + ((MUIT * 45) - (MUIT * 35)) * 0.2 + ((MUIT * 35) - (MUIT * 20)) * 0.17 + ((MUIT * 20) - (MUIT * 5)) * 0.14, 2) + ((MUIT * 5) * 0.08)
                     mCalExtOrd = Round(((mtope - (MUIT * 45)) * 0.3) + ((MUIT * 45) - (MUIT * 35)) * 0.2 + ((MUIT * 35) - (MUIT * 20)) * 0.17 + ((MUIT * 20) - (MUIT * 5)) * 0.14, 2) + ((MUIT * 5) * 0.08) - mFactor
                     mFacUti = 0.3
            End Select
'                Case Is < (Round(MUIT * 12, 2) + 1)
'                     If (mtope2 - mExtraOrd) < Round(MUIT * 7, 2) Then
'                        mFactor = 0
'                        mCalExtOrd = Round((mtope2 - Round(MUIT * 7, 2)) * 8 / 100, 2)
'                        mFacUti = 0.08
'                     Else
'                        mFactor = Round((mtope - mExtraOrd) * 0.08, 2)
'                        mCalExtOrd = Round(mtope * 0.08, 2) - mFactor
'                        mFacUti = 0.08
'                     End If
'                Case Is < (Round(MUIT * 27, 2) + 1)
'                     mFactor = Round((((mtope - mExtraOrd) - (MUIT * 12)) * 0.14) + (MUIT * 12) * 0.08, 2)
'                     mCalExtOrd = Round(((mtope - (MUIT * 12)) * 0.14) + (MUIT * 12) * 0.08, 2) - mFactor
'                     mFacUti = 0.14
'                Case Is < (Round(MUIT * 42, 2) + 1)
'                     mFactor = Round((((mtope - mExtraOrd) - (MUIT * 27)) * 0.17) + ((MUIT * 27) - (MUIT * 12)) * 0.14, 2) + ((MUIT * 12) * 0.08)
'                     mCalExtOrd = Round(((mtope - (MUIT * 27)) * 0.17) + ((MUIT * 27) - (MUIT * 12)) * 0.14, 2) + ((MUIT * 12) * 0.08) - mFactor
'                     mFacUti = 0.17
'                Case Is < (Round(MUIT * 52, 2) + 1)
'                     mFactor = Round((((mtope - mExtraOrd) - (MUIT * 42)) * 0.2) + ((MUIT * 42) - (MUIT * 27)) * 0.17 + ((MUIT * 27) - (MUIT * 12)) * 0.14, 2) + ((MUIT * 12) * 0.08)
'                     mCalExtOrd = Round(((mtope - (MUIT * 42)) * 0.2) + ((MUIT * 42) - (MUIT * 27)) * 0.17 + ((MUIT * 27) - (MUIT * 12)) * 0.14, 2) + ((MUIT * 12) * 0.08) - mFactor
'                     mFacUti = 0.2
'                Case Else
'                     mFactor = Round((((mtope - mExtraOrd) - (MUIT * 52)) * 0.3) + ((MUIT * 52) - (MUIT * 42)) * 0.2 + ((MUIT * 42) - (MUIT * 27)) * 0.17 + ((MUIT * 27) - (MUIT * 12)) * 0.14, 2) + ((MUIT * 12) * 0.08)
'                     mCalExtOrd = Round(((mtope - (MUIT * 52)) * 0.3) + ((MUIT * 52) - (MUIT * 42)) * 0.2 + ((MUIT * 42) - (MUIT * 27)) * 0.17 + ((MUIT * 27) - (MUIT * 12)) * 0.14, 2) + ((MUIT * 12) * 0.08) - mFactor
'                     mFacUti = 0.3
'            End Select
         Else
            Select Case mtope
                Case Is < (Round(MUIT * 27, 2) + 1)
                     'mFactor = Round(mtope * 0.15, 2)
                     If (mtope2 - mExtraOrd) < Round(MUIT * 7, 2) Then
                        mFactor = 0
                        mCalExtOrd = Round((mtope2 - Round(MUIT * 7, 2)) * 15 / 100, 2)
                        mFacUti = 0.15
                     Else
                        mFactor = Round((mtope - mExtraOrd) * 0.15, 2)
                        mCalExtOrd = Round(mtope * 0.15, 2) - mFactor
                        mFacUti = 0.15
                     End If
                Case Is < (Round(MUIT * 54, 2) + 1)
                     'mFactor = Round(((mtope - (MUIT * 27)) * 0.21) + (MUIT * 27) * 0.15, 2)
                     mFactor = Round((((mtope - mExtraOrd) - (MUIT * 27)) * 0.21) + (MUIT * 27) * 0.15, 2)
                     mCalExtOrd = Round(((mtope - (MUIT * 27)) * 0.21) + (MUIT * 27) * 0.15, 2) - mFactor
                     mFacUti = 0.21
                Case Else
                     mFactor = Round((((mtope - mExtraOrd) - (MUIT * 54)) * 0.3) + ((MUIT * 54) - (MUIT * 27)) * 0.21, 2) + ((MUIT * 27) * 0.15)
                     mCalExtOrd = Round(((mtope - (MUIT * 54)) * 0.3) + ((MUIT * 54) - (MUIT * 27)) * 0.21, 2) + ((MUIT * 27) * 0.15) - mFactor
                     mFacUti = 0.3
            End Select
         End If
         'quinta descontada en otras empresas
         Dim QtaOtras As Currency
         QtaOtras = 0
         If IngresoOtrasEmpresas = "S" Then
            Sql$ = "select isnull(SUM(Quinta),0) as Quinta from Pla_Trab_Otras_Empresas_Meses where Cia='" & wcia & "' AND Ayo=" & Vano & " and Mes<=" & Vmes & " AND placod='" & Txtcodpla.Text & "'  and Status<>'*'"
            If (fAbrRst(rsF02, Sql$)) Then QtaOtras = rsF02(0)
            If rsF02.State = 1 Then rsF02.Close
         End If

         
         If QtaGra And VTipobol = "03" Then
            F02 = "Select Round(" & mCalExtOrd & ", 2)"
         ElseIf VTipobol = "11" Then
             If mFactor - (macus + QtaOtras) >= Round(mExtraOrd * mFacUti, 2) Then
                F02 = "Select Round(" & mExtraOrd & " * " & mFacUti & ", 2)"
             Else
                If mtope > mExtraOrd Then
                   F02 = "Select Round(" & mExtraOrd & " * " & mFacUti & ", 2)"
                Else
                   If mtope > 0 And mExtraOrd > mtope And (mtope * mFacUti) > macus Then
                      F02 = "Select Round((" & (mtope * mFacUti) & " - " & macus & "), 2)"
                   Else
                      F02 = "Select Round((" & mFactor & " - " & macus & "), 2)"
                   End If
                End If
             End If
         Else
            If vTipoTra = "01" Then
               If wtipodoc = True Or wGrupoPla = "01" Then 'En quincena Grupo GALLOS trabaja con 240 hora como una boleta normal
                  F02 = "Select Round((((" & mFactor & " - " & (macus + QtaOtras) & ") / (12 - " & Vmes & " + 1))+(" & mCalExtOrd & "))-(" & macuqtames & "), 2)"
               Else
                  F02 = "Select Round((((" & mFactor & " - " & (macus + QtaOtras) & ") / (12 - " & Vmes & " + 1))+(" & mCalExtOrd & "))-(" & macuqtames & "), 2)"
               End If
            Else
               If VTipobol = "02" Then
                  F02 = "Select Round(((" & mFactor & " - " & (macus + QtaOtras) & ") / (12 - " & Vmes & " + 1))+(" & mCalExtOrd & "), 2)"
               Else
                  If Vmes = 12 And EncuentraVaca Then
                     F02 = "Select Round((" & mFactor & " - " & (macus - QtaOtras) & ") + (" & mCalExtOrd & "), 2)"
                  Else
                     'mgirao
                     If Val(VSemana) > Val(msemano) Then VSemana = msemano
                     If mCalExtOrd < 0 Then
                        If QuintaMes + mCalExtOrd > 0 Then
                           F02 = "Select Round(" & mCalExtOrd & ", 2)"
                        Else
                           'aca verifico la quinta en negativo ---Revisar ojo
                           'F02 = "Select Round(" & QuintaMes * -1 & ", 2)"
                            F02 = "Select Round(" & QuintaMes * 0 & ", 2)"
                        End If
                     Else
                        F02 = "Select Round(((" & mFactor & " - " & (macus - QtaOtras) & ")  / (" & msemano & " - " & Val(VSemana) & " + 1))+(" & mCalExtOrd & "), 2)"
                     End If
                     
                  End If
               End If
            End If
         End If
      End If
      If VTipobol = "02" Then VSemana = vSemVaca
      If VTipobol = "01" And LbLSemBolVac.Caption <> "" Then VSemana = vSemVaca
   End If
quinta:
ElseIf concepto = "15" Then
       MsgBox "666"
End If
macui = 0: macus = 0: macomi = 0: Qtames = 0: QuintaMes = 0:  mCalExtOrd = 0
End Function
Private Function Devuelve_Ultima_Semana() As String
Devuelve_Ultima_Semana = ""
Dim Rq As ADODB.Recordset
Sql = "select top 1 semana from plahistorico where cia='" & wcia & "' and year(fechaproceso)=" & Vano & " and tipotrab='02' and proceso='01' and status<>'*' order by convert(integer,semana) desc"
If fAbrRst(Rq, Sql) Then Devuelve_Ultima_Semana = Trim(Rq(0) & "")
Rq.Close
End Function
Private Sub Carga_Horas()
Dim rs2 As ADODB.Recordset
Dim mbol As String
Dim mconceptos As String
Dim mhor As Currency
Dim mdiasfalta As String
Dim mdiasferiado As String
Dim wBeginMonth As String
Dim mHourTra As Currency
Dim con As String
Dim I As Integer
Dim NroTrans As Integer
On Error GoTo CORRIGE
NroTrans = 0
If Trim(Txtcodpla.Text) = "" Then Exit Sub
'If VTipobol = "04" Then Exit Sub
If VTipobol = "11" Then Exit Sub


'If VTipobol = "05" And wGrupoPla <> "01" Then Exit Sub
 t_horas_trabajadas = 0
Sql$ = "select iniciomes from cia where cod_cia='" & _
wcia & "' and status<>'*'"

If (fAbrRst(Rs, Sql$)) Then
   If IsNull(Rs!iniciomes) Then wBeginMonth = "1" Else wBeginMonth = Rs!iniciomes
End If
Rs.Close

If Trim(wBeginMonth) <> "1" Then
   If Vmes = 1 Then
      wBeginMonth = Format(wBeginMonth, "00") & "/12/" & Format(Vano - 1, "0000")
   Else
     wBeginMonth = Format(wBeginMonth, "00") & "/" & Format(Vmes - 1, "00") & "/" & Format(Vano, "0000")
   End If
Else
   wBeginMonth = Format(wBeginMonth, "00") & "/" & Format(Vmes, "00") & "/" & Format(Vano, "0000")
End If

If VPerPago = "02" Then wBeginMonth = VfDel

If rshoras.RecordCount > 0 Then rshoras.MoveFirst
Do While Not rshoras.EOF
   rshoras.Delete
   rshoras.MoveNext
Loop
'cn.BeginTrans
Sql$ = wInicioTrans & " CARGA_HORAS"
cn.Execute Sql$

NroTrans = 1

Sql$ = "delete from platemphist where cia='" & wcia & "' and placod='" & Trim(Txtcodpla.Text) & "'"
cn.Execute Sql
'cn.CommitTrans
Sql$ = wFinTrans & " CARGAR_HORAS"
cn.Execute Sql$

NroTrans = 0

mdiasfalta = 0
mdiasferiado = 0
VHorasnormal = VHoras
'If VTipobol = "01" Or VTipobol = "10" Or (VTipobol = "05" And wGrupoPla = "01") Then mbol = "N"
If VTipobol = "01" Or VTipobol = "10" Or VTipobol = "05" Then mbol = "N"
If VTipobol = "02" Then mbol = "V"
If VTipobol = "03" Then mbol = "G"
If VTipobol = "04" Then mbol = "N"

wciamae = Funciones.Determina_Maestro("01077")

If (VTipobol = "01" Or VTipobol = "10" Or VTipobol = "04") Or (VTipobol = "05" And wGrupoPla = "01") Or wtipodoc <> True Then

   Sql$ = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
    
   If wtipodoc = True Then
     
      Select Case VPerPago
             Case Is = "04" 'Mensual
                  Sql$ = Sql$ & "select distinct(concepto) from platareo where fecha " _
                       & "BETWEEN '" & Format(wBeginMonth, FormatFecha) + Coneccion.FormatTimei & _
                       "' AND '" & Format(VFProceso, FormatFecha) + FormatTimef & "' " _
                       & "and codigotrab='" & Trim(Txtcodpla) & "' and status<>'*'"
                       
             Case Is = "02" 'Semanal
                  Sql$ = Sql$ & "select distinct(concepto) from platareo where fecha " _
                       & " BETWEEN '" & Format(VfDel, FormatFecha) + FormatTimei & "' AND '" & Format(VfAl, FormatFecha) + FormatTimef & "' " _
                       & "and codigotrab='" & Trim(Txtcodpla) & "' and status<>'*'"
      End Select
   Else
      Sql$ = Sql$ & "select distinct(concepto) from platareo where fecha " _
           & "BETWEEN '" & Format(wBeginMonth, FormatFecha) + FormatTimei & _
           "' AND '" & Format(VFProceso, FormatFecha) + FormatTimef & "' " _
           & "and codigotrab='" & Trim(Txtcodpla) & "' and status<>'*'"
      mbol = "N"
   End If
   
   If (fAbrRst(Rs, Sql$)) Then
      Rs.MoveFirst
      mconceptos = ""
      Do While Not Rs.EOF
         mconceptos = mconceptos & "'" & Trim(Rs!concepto) & "',"
         Rs.MoveNext
      Loop
      mconceptos = "(cod_maestro2 in (" & Mid(mconceptos, 1, Len(mconceptos) - 1) & ")"
   End If
   
   If Rs.State = 1 Then Rs.Close
   
   If Trim(mconceptos) <> "" Then
      Sql$ = "Select * from maestros_2 where status<>'*' and CHARINDEX('" & mbol & "',flag1 )>0 and " & mconceptos
   Else
      Sql$ = "Select * from maestros_2 where status<>'*' and CHARINDEX('" & mbol & "',flag1 )>0 "
   End If
   
   Dim mFlagBol As String
   Select Case VTipobol
          Case "02": mFlagBol = "V"
          Case "03": mFlagBol = "G"
          Case Else: mFlagBol = "N"
   End Select
   
   If vTipoTra = "01" Then
      If Trim(mconceptos) <> "" Then
         Sql$ = Sql$ & " or cod_maestro2 in('01'))"
      Else
         If VTipobol = "10" Then
            Sql$ = Sql$ & " and cod_maestro2 in('01')"
         Else
            Sql$ = Sql$ & " and cod_maestro2 in  (Select cod_maestro2 from maestros_2 where status<>'*' and right(ciamaestro,3)= '077'  and flag7<>'' and not flag7 is null and CHARINDEX('" & mFlagBol & "',flag1 )>0 )"
         End If
      End If
   Else
     If Trim(mconceptos) <> "" Then
         Sql$ = Sql$ & " or cod_maestro2 in (Select cod_maestro2 from maestros_2 where status<>'*' and right(ciamaestro,3)= '077'  and flag7<>'' and not flag7 is null) and CHARINDEX('" & mFlagBol & "',flag1 )>0)"
     Else
         Sql$ = Sql$ & " and cod_maestro2 in (Select cod_maestro2 from maestros_2 where status<>'*' and right(ciamaestro,3)= '077'  and flag7<>'' and not flag7 is null and CHARINDEX('" & mFlagBol & "',flag1 )>0 ) "
     End If
   End If
   
   Sql$ = Sql$ & wciamae
   Dim rsOrdenHoras As ADODB.Recordset
   If (fAbrRst(Rs, Sql$)) Then
      Rs.MoveFirst
      con = ""
      Sql = "Select cod_maestro2 from maestros_2 where status<>'*' and right(ciamaestro,3)= '077'  and flag7<>'' and not flag7 is null order by convert(integer,flag7)"
      If (fAbrRst(rsOrdenHoras, Sql$)) Then
         rsOrdenHoras.MoveFirst
         Do While Not rsOrdenHoras.EOF
            con = con & Trim(rsOrdenHoras!cod_maestro2 & "")
            rsOrdenHoras.MoveNext
         Loop
      End If
      rsOrdenHoras.Close: Set rsOrdenHoras = Nothing
      'con = "1401020304060507080910171112131518192324252627"
      
       
      For I = 1 To (Len(con) / 2)
      'Do While Not RS.EOF
         Rs.Filter = "COD_MAESTRO2='" & Mid(con, I + (I - 1), 2) & "'"
         Debug.Print Mid(con, I + (I - 1), 2)
        If Not Rs.EOF Then  ' {<MA>} 01/02/2007
            rshoras.AddNew
            rshoras!Codigo = Trim(Rs!cod_maestro2)
            rshoras!Descripcion = UCase(Rs!DESCRIP)
            If Trim(Rs("cod_maestro2")) = "14" And vTipoTra = "02" Then
              rshoras("MONTO") = 6
            End If

            If vTipoTra = "02" Then
                If rshoras!Codigo = "01" Then
                    rshoras!Monto = 6 * 8
                    t_horas_trabajadas = rshoras!Monto
                End If
            
                If rshoras!Codigo = "02" Then
                    rshoras!Monto = (6 / 6) * 8
                End If
            End If
            
         mhor = 0
         
         Sql$ = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
                  
         Select Case VPerPago
               Case Is = "04"
                    Sql$ = Sql$ & "select concepto,motivo,sum(tiempo) as tiempo from platareo where fecha " _
                         & "BETWEEN '" & Format(wBeginMonth, FormatFecha) + FormatTimei & "' AND '" & Format(VFProceso, FormatFecha) + FormatTimef & "' " _
                         & "and codigotrab='" & Trim(Txtcodpla) & "' and concepto='" & _
                         Trim(Rs!cod_maestro2) & "' and status<>'*'  group by concepto,motivo order by concepto,motivo"
               Case Is = "02"
                    Sql$ = Sql$ & "select concepto,motivo,sum(tiempo) as tiempo from platareo where fecha " _
                         & " BETWEEN '" & Format(VfDel, FormatFecha) + FormatTimei & "' AND '" & _
                         Format(VfAl, FormatFecha) + FormatTimef & "' " _
                         & "and codigotrab='" & Trim(Txtcodpla) & "' and concepto='" & Trim(Rs!cod_maestro2) & "' and status<>'*'  group by concepto,motivo order by concepto,motivo"
         End Select
 
         If (fAbrRst(rs2, Sql$)) Then
            rs2.MoveFirst
            Do While Not rs2.EOF
               Select Case rs2!motivo
                      Case Is = "DI"
                           'If S Then
                              mhor = mhor + (rs2!tiempo * 8 * 60)
                              mdiasfalta = mdiasfalta + rs2!tiempo
                              If Rs!cod_maestro2 = "03" Then mdiasferiado = mdiasferiado + rs2!tiempo
                           'Else
                            '  MsgBox "HOLA"
                           'End If
                      Case Is = "HO"
                           mhor = mhor + Int(rs2!tiempo) * 60 + ((rs2!tiempo - Int(rs2!tiempo)) * 100)
                      Case Is = "MI"
                           mhor = mhor + rs2!tiempo
               End Select
               rs2.MoveNext
            Loop
            If rs2.State = 1 Then rs2.Close
            mhor = Int(mhor / 60) + ((mhor Mod 60) / 100)
                        
            rshoras!Monto = IIf(mhor = 0, Null, mhor)
            
            If Trim(Rs!flag2) = "-" Then
               VHorasnormal = VHorasnormal - mhor
            End If
         End If
         
         If rshoras("CODIGO") = "14" And IsNull(rshoras("MONTO")) And vTipoTra = "02" Then
            rshoras("MONTO") = 6
         End If
         
         Rs.MoveNext
      'Loop
      End If
      Next
      
      'TIPICO
       
      rshoras.MoveFirst
      mHourTra = 0
      If rshoras!Codigo = "01" Then
         If wtipodoc = True Then
            mHourTra = Calc_Horas_FecIng(wBeginMonth)
            If VTipobol <> "10" Then
                rshoras!Monto = VHorasnormal
                t_horas_trabajadas = VHorasnormal
            Else
                rshoras!Monto = 0
            End If
         Else
            VHorasnormal = 0
            If VNewBoleta = True Then VHorasnormal = Calc_Horas_Quincena(wBeginMonth)
             t_horas_trabajadas = VHorasnormal
            rshoras!Monto = VHorasnormal
         End If
         If rshoras.RecordCount > 1 Then rshoras.MoveNext
      End If
      If vTipoTra <> "01" Then
        If rshoras!Codigo = "02" Then rshoras!Monto = ((6 - mdiasfalta + mdiasferiado) * 8) / 6
      End If
      
      If mHourTra <> 0 Then
         rshoras.AddNew
         rshoras!Codigo = "03"
         rshoras!Descripcion = "FERIADOS"
         rshoras!Monto = mHourTra
      End If
      
   End If
Else
    Sql$ = "Select * from maestros_2 where status<>'*' and CHARINDEX('" & mbol & "',flag1 )>0"
    Sql$ = Sql$ & wciamae
    cn.CursorLocation = adUseClient
    Set Rs = New ADODB.Recordset
    Set Rs = cn.Execute(Sql$, 64)
    If rshoras.RecordCount > 0 Then
       rshoras.MoveFirst
       Do While Not rshoras.EOF
          rshoras.Delete
          rshoras.MoveNext
       Loop
    End If
    If Not Rs.RecordCount > 0 Then
        If (VTipobol <> "04" And VTipobol <> "05" And VTipobol <> "11" And VTipobol <> "10") Then
        MsgBox "No Existen Horas Registradas", vbCritical, TitMsg: Exit Sub
        End If
    End If
    If Rs.RecordCount > 0 Then Rs.MoveFirst
    Do While Not Rs.EOF
       rshoras.AddNew
       rshoras!Codigo = Trim(Rs!cod_maestro2)
       rshoras!Descripcion = Trim(Rs!DESCRIP)
       If VTipobol = "04" Then rshoras!Monto = VHoras Else rshoras!Monto = "0.00"
       Rs.MoveNext
    Loop
    If Rs.State = 1 Then Rs.Close
End If

'solo para datos guardados
Call llena_horas

Exit Sub
CORRIGE:
If NroTrans = 1 Then
    'cn.RollbackTrans
    Sql$ = wCancelTrans & " CARGA_HORAS"
    cn.Execute Sql$
End If
MsgBox "Error :" & Err.Description, vbCritical, "Sistema de Planillas"
End Sub
Public Function F03(concepto As String, ByVal VTipobol As String) As String  'APORTACIONES
Dim rsF03 As ADODB.Recordset
Dim rscalculo As ADODB.Recordset
Dim F03str As String
Dim mFactor As Currency
Dim mbol As String
mFactor = 0
F03 = ""

If (concepto = "15" Or concepto = "17") Then
   If Trim(Lblcodafp.Caption) = "01" Or Trim(LblCodTrabSunat.Caption & "") <> "37" Then
       F03 = "select 0"
       Exit Function
    End If
End If

mbol = VTipobol
If VTipobol = "04" Then mbol = "01"

If lSubsidio = "S" And wGrupoPla = "01" Then 'No hay Aportaciones cuando es subsidio para grupo Gallos
   F03 = "Select 0": Exit Function
End If

If concepto = "03" Then
   If Not (VTipobol = "01" Or VTipobol = "02") Then F03 = "Select 0": Exit Function
   Sql$ = "select senati from cia where cod_cia='" & wcia & "' and status<>'*'"
   If Not (fAbrRst(rsF03, Sql$)) Then F03 = "Select 0": Exit Function
   If Trim(rsF03!senati) <> "S" Then F03 = "Select 0": Exit Function
   If rsF03.State = 1 Then rsF03.Close
      
'   Sql$ = "select * from planilla_ccosto a"
'   Sql$ = Sql$ & " inner join maestros_32 b on a.ccosto=b.cod_maestro3"
'   Sql$ = Sql$ & " where a.cia='" & wcia & "' and a.placod='" & Trim(Txtcodpla.Text) & "' AND a.status !='*' and"
'   Sql$ = Sql$ & " b.ciamaestro='" & wcia & "055' and b.cod_maestro2='" & Trim(vTipoTra) & "' and flag3='1' and b.status !='*'"
      
'   Sql$ = "Select * from maestros_2 where ciamaestro='" & wciamae & "' and cod_maestro2='" & Trim(VArea) & "' and status<>'*'"
'   'wciamae

   Sql$ = "select * from planilla_ccosto a"
   Sql$ = Sql$ & " inner join PLA_CCOSTOS b on a.ccosto=b.codigo"
   Sql$ = Sql$ & " where a.cia='" & wcia & "' and a.placod='" & Trim(Txtcodpla.Text) & "' AND a.status !='*' and"
   Sql$ = Sql$ & " b.cia=a.cia and b.codigo=a.ccosto and senati='1' and b.status !='*'"
   
   If Not (fAbrRst(rsF03, Sql$)) Then
      F03 = "Select 0": Exit Function
   End If
   If rsF03.State = 1 Then rsF03.Close
End If


If concepto = "18" Or concepto = "19" Then

    Dim mperiodosctr, xcodsctr As String
    Dim xporsalud, xporpension As Currency
    Dim xtopemax As Double
    Dim mtope As Double

    mperiodosctr = ""
 
    Sql$ = "select cod_remu from plaafectos where cia='" & wcia & "' and tipo='A' and tboleta='" & mbol & "'  and  codigo='" & concepto & "' and status<>'*'"

    If (fAbrRst(rsF03, Sql$)) Then
         rsF03.MoveFirst
         F03str = ""
         Do While Not rsF03.EOF
             F03str = F03str & "i" & Trim(rsF03!cod_remu) & "+"
               
             rsF03.MoveNext
         Loop
            
         F03str = "(" & Mid(F03str, 1, Len(Trim(F03str)) - 1) & ")"
        
         mperiodosctr = Format(Vano, "0000") & Format(Vmes, "00")
         xcodsctr = Bcodsctr(Txtcodpla.Text)
    
         Sql$ = "SELECT porsalud,porpension,tope FROM PLASCTR where periodo='" & mperiodosctr & "' and codSCTR='" & xcodsctr & "' and status<>'*' and cia='" & wcia & "'"
    
        If Not (fAbrRst(rsF03, Sql$)) Then
            If bImportar = False Then
              MsgBox "No se Encuentran Factores de Calculo para SCTR", vbCritical, "Calculo de Boleta"
            Else
               LstObs.AddItem Trim(Txtcodpla.Text) & ": No se Encuentran Factores de Calculo para SCTR"
            End If
            mcancel = True
            Exit Function
        Else
            xporsalud = rsF03!porsalud
            xporpension = rsF03!porpension
            xtopemax = rsF03!tope
        End If
    
    If concepto = "18" Then
        mFactor = xporsalud
    Else
     mFactor = xporpension
    End If
    
    mtope = xtopemax
    If VTipobol <> "04" Then Call Acumula_Mes_SCTR(concepto, "A")
    F03 = "select (" & F03str & " + " & macui & ") , (" & mFactor & " /100)," & macus & " from platemphist "
    F03 = F03 & "where cia='" & wcia & "' and status='" & Mstatus & "' and placod='" & Txtcodpla.Text & "' "
    F03 = F03 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & ""
      
    If (fAbrRst(rscalculo, F03)) Then
       If rscalculo(0) > xtopemax And concepto = "19" Then
          F03 = " SELECT " & Round((mtope * (mFactor / 100)) - macus, 2)
       Else
          F03 = " SELECT " & Round((rscalculo(0) * (mFactor / 100)) - macus, 2)
       End If
    End If
 End If
 macui = 0: macus = 0: macomi = 0
 Exit Function
End If

    'MODIFICACION PARA EL CALCULO DE LA EPS
    Dim rsTemporal  As ADODB.Recordset
    Dim Afecto_EPS  As Boolean
    Dim Por_EPS     As Double
    
'    Cadena = "SELECT ISNULL(APORTACION_EPS,0) AS APORTACION_EPS FROM CIA WHERE COD_CIA = '" & wcia & "' AND STATUS != '*'"
'    Set rsTemporal = OpenRecordset(Cadena, cn)
'    If Not rsTemporal.EOF Then
'        Afecto_EPS = CBool(rsTemporal!Aportacion_EPS)
'    Else
'        Afecto_EPS = False
'    End If
    If concepto = "01" Then
        If Verifica_Subsidio Then
           F03 = "Select 0": Exit Function
        End If
        Cadena = "SELECT AFILIADO_EPS_SERV FROM PLANILLAS WHERE CIA = '" & wcia & "' AND STATUS != '*' AND PLACOD = '" & Txtcodpla.Text & "'"
        Set rsTemporal = OpenRecordset(Cadena, cn)
        If Trim(rsTemporal!afiliado_eps_serv & "") = True Then
           Afecto_EPS = True
        Else
            Afecto_EPS = False
        End If
    
        If rsTemporal.State = adStateOpen Then rsTemporal.Close
        If Afecto_EPS Then
            Cadena = "SELECT ISNULL(CODIGO_EPS,'') AS CODIGO_EPS FROM PLANILLAS WHERE CIA = '" & wcia & "' AND PLACOD = '" & Trim(Txtcodpla.Text) & "' AND STATUS != '*' AND TIPOTRABAJADOR = '" & vTipoTra & "'"
            Set rsTemporal = OpenRecordset(Cadena, cn)
            If Not rsTemporal.EOF Then
                If rsTemporal!codigo_eps <> "" Then
                    Cadena = "SELECT ISNULL(PLAIMPORTE,0) AS IMPORTE FROM MAESTROS_2 WHERE RIGHT(CIAMAESTRO,3) = '143' AND STATUS != '*' AND RTRIM(ISNULL(CODSUNAT,'')) != '' AND COD_MAESTRO2 = '" & Trim(rsTemporal!codigo_eps) & "'"
                    If rsTemporal.State = adStateOpen Then rsTemporal.Close
                    Set rsTemporal = OpenRecordset(Cadena, cn)
                    If Not rsTemporal.EOF Then
                        Por_EPS = Val(rsTemporal!importe)
                    Else
                        Por_EPS = 0
                    End If
                Else
                    Por_EPS = 0
                End If
            Else
                Por_EPS = 0
            End If
        End If
        
        If rsTemporal.State = adStateOpen Then rsTemporal.Close
        
    End If
    
    Sql$ = "select aportacion from placonstante where cia='" & wcia & "' and tipomovimiento='03' and codinterno='" & concepto & "' and aportacion<>0 and status<>'*'"
    If (fAbrRst(rsF03, Sql$)) Then
       If Not IsNull(rsF03!aportacion) Then
          If concepto = "01" And VTipobol = "03" Then
             If rsF03!aportacion <> 0 Then mFactor = IIf(Afecto_EPS = True, Por_EPS, rsF03!aportacion)
          Else
             If rsF03!aportacion <> 0 Then mFactor = rsF03!aportacion
          End If
       End If
    End If
    If rsF03.State = 1 Then rsF03.Close
    
If mFactor <> 0 Then
   Sql$ = "select cod_remu from plaafectos where cia='" & wcia & "' and tipo='A' and tboleta='" & mbol & "'  and  codigo='" & concepto & "' and status<>'*'"
   If (fAbrRst(rsF03, Sql$)) Then
      rsF03.MoveFirst
      F03str = ""
      Do While Not rsF03.EOF
         F03str = F03str & "i" & Trim(rsF03!cod_remu) & "+"
         rsF03.MoveNext
      Loop
      
      F03str = "(" & Mid(F03str, 1, Len(Trim(F03str)) - 1) & ")"
      If rsF03.State = 1 Then rsF03.Close
      
      If VTipobol <> "04" Then Call Acumula_Mes(concepto, "A")
      
      F03 = "select (" & F03str & " + " & macui & ") , (" & mFactor & " /100)," & macus & " from platemphist "
      F03 = F03 & "where cia='" & wcia & "' and status='" & Mstatus & "' and placod='" & Txtcodpla.Text & "' "
      F03 = F03 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & ""
      
      
      Dim lCalcSobSMin As Boolean
      lCalcSobSMin = True
      If wGrupoPla = "01" Then lCalcSobSMin = False 'No Calculo de aportaciones sobre sueldo minimo Grupo Gallos

      If (fAbrRst(rscalculo, F03)) Then
        If VTipobol = "03" Then
           F03 = " SELECT " & Round((rscalculo(0) * (mFactor / 100)) - macus, 2)
        Else
          If rscalculo(0) > sueldominimo Or (concepto <> "01") Or Not lCalcSobSMin Then
              'F03 = " SELECT " & Round((rscalculo(0) * (mFactor / 100)) - macus, 2)
              F03 = " SELECT " & Format((rscalculo(0) * (mFactor / 100)) - macus, "#########.00")
          Else
              'F03 = "SELECT " & Round((sueldominimo * (mFactor / 100)) - macus, 2)
              F03 = "SELECT " & Format((sueldominimo * (mFactor / 100)) - macus, "#########.00")
          End If
        End If
      End If
   End If
End If

macui = 0: macus = 0: macomi = 0
End Function
Public Function V01(concepto As String) As String
'INGRESOS
Dim rsV01 As ADODB.Recordset
Dim mFactor As Currency
Dim fACTORfAMILIA As String
Dim sPeriodoPago As String


'FECHA DE MODIFICACION 08/01/2008
'If wcia = "05" Then fACTORfAMILIA = 2 Else fACTORfAMILIA = 1

If wcia = "05" Then fACTORfAMILIA = 1 Else fACTORfAMILIA = 1

mFactor = 0
V01 = ""
Select Case concepto
       Case Is = "02" 'ASIGNACION FAMILIAR
            Dim RX As New ADODB.Recordset
            sSQL = "select Tipo from plaremunbase where cia = '" & wcia & "' and status != '*' and placod = '" & Trim(Txtcodpla.Text) & "' and concepto = '" & concepto & "'"
            If (fAbrRst(RX, sSQL)) Then sPeriodoPago = RX(0)
            If RX.State = 1 Then RX.Close
            
            If sPeriodoPago = "02" Then
               V01 = "select round(((b.importe*4)/240)*(a.h12*" & fACTORfAMILIA & "),2) as basico from platemphist a,plaremunbase b "
               V01 = V01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Txtcodpla.Text & "' "
               V01 = V01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
               V01 = V01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & concepto & "' and b.status<>'*'"
            Else
               V01 = "select round((b.importe/factor_horas)*(a.h12*" & fACTORfAMILIA & "),2) as basico from platemphist a,plaremunbase b "
               V01 = V01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Txtcodpla.Text & "' "
               V01 = V01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
               V01 = V01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & concepto & "' and b.status<>'*'"
            End If
       Case Is = "04" 'BONIFICACION T. SERVICIO
            V01 = "select round((b.importe/factor_horas)*(a.h12),2) as basico from platemphist a,plaremunbase b "
            V01 = V01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Txtcodpla.Text & "' "
            V01 = V01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
            V01 = V01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & concepto & "' and b.status<>'*'"
       Case Is = "05" 'INCREMENTO AFP 10.23%
            V01 = "select round((b.importe/CASE WHEN TIPO=02 THEN factor_horas+8 ELSE FACTOR_HORAS END)*(a.h12),2) as basico from platemphist a,plaremunbase b "
            V01 = V01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Txtcodpla.Text & "' "
            V01 = V01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
            V01 = V01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & concepto & "' and b.status<>'*'"
       Case Is = "06" 'INCREMENTO AFP 3%
            V01 = "select round((b.importe/CASE WHEN TIPO=02 THEN factor_horas+8 ELSE FACTOR_HORAS END)*(a.h12),2) as basico from platemphist a,plaremunbase b "
            V01 = V01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Txtcodpla.Text & "' "
            V01 = V01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
            V01 = V01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & concepto & "' and b.status<>'*'"
       Case Is = "07" 'BONIFICACION COSTO DE VIDA
            V01 = "select round((b.importe/factor_horas)*(a.h12),2) as basico from platemphist a,plaremunbase b "
            V01 = V01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Txtcodpla.Text & "' "
            V01 = V01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
            V01 = V01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & concepto & "' and b.status<>'*'"
       Case Is = "14" 'VACACIONES
            V01 = "select round((b.importe/factor_horas)*(a.h12),2) as basico from platemphist a,plaremunbase b "
            V01 = V01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Txtcodpla.Text & "' "
            V01 = V01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
            V01 = V01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
       Case Is = "20" 'BUC
            Sql$ = "select factor from platasaanexo where modulo='02' and tipomovimiento='" & Trim(Lbltope.Caption) & "' and codinterno='" & concepto & "' and tipotrab='" & vTipoTra & "' and status<>'*'"
            'Sql$ = "select factor from platasaanexo where modulo='02' and tipomovimiento='" & Trim(Lbltope.Caption) & "' and codinterno='" & concepto & "' and tipotrab='" & vTipoTra & "' and cargo='" & Trim(Lblcargo.Caption) & "' and status<>'*'"
            If (fAbrRst(rsV01, Sql$)) Then
               mFactor = rsV01!factor
               If rsV01.State = 1 Then rsV01.Close
               V01 = "select round(((b.importe/factor_horas)* " & mFactor & ") *(a.h12),2) as basico from platemphist a,plaremunbase b "
               V01 = V01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Txtcodpla.Text & "' "
               V01 = V01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
               V01 = V01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
            End If
       Case Is = "34" 'FERIADO 1ERO DE MAYO
           Sql$ = "SELECT Sum((case factor_horas when 48 then importe*4 When 240 then importe End)/30) as Factor "
            Sql$ = Sql$ & "FROM plaremunbase WHERE Cia='" & Trim(wcia) & "' AND Concepto in('02','05','06') AND status<>'*' "
            Sql$ = Sql$ & "and placod='" & Trim(Txtcodpla.Text) & "' "
            If (fAbrRst(RX, Sql$)) Then
               If IsNull(RX!factor) Then
                  mFactor = 0
               Else
                  mFactor = RX!factor
               End If
               If RX.State = 1 Then RX.Close
            End If
            
            V01 = "select Case a.H15 when 0 then 0 else round((b.importe/factor_horas)*a.h15,2) + " & mFactor & " End as basico,A.H15 from platemphist a,plaremunbase b "
            V01 = V01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Trim(Txtcodpla.Text) & "' "
            V01 = V01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
            V01 = V01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
            mFactor = 0
       Case Is = "129"
       
        'MODIFICADO POR RICARDO HINOSTROZA
        'AGREGAR CALCULO DE HORARIO VACACIONAL
        'FEC MODIFICACION 10/04/2008
        
             V01 = "SELECT ROUND(SUM(I29)/6,2)  AS BASICO FROM PLAHISTORICO WHERE CIA ='" & wcia & "' AND STATUS<>'*' AND PLACOD='" & Txtcodpla.Text & "' AND YEAR(FECHAPROCESO) =" & Vano & " AND (MONTH(FECHAPROCESO) BETWEEN 1 AND 3)"
       
End Select
'Debug.Print concepto
'Debug.Print SQL$
End Function
Public Function G01(concepto As String) As String 'INGRESOS
Dim rsG01 As ADODB.Recordset
Dim mFactor As Currency
Dim sPeriodoPago As String
Dim mh As Integer
mFactor = 0
If Val(Mid(VFProceso, 4, 2)) = 7 Then mh = 7 Else mh = 5
G01 = ""
If vTipoTra <> "05" Then
    Select Case concepto
           Case Is = "02" 'ASIGNACION FAMILIAR
           
                Dim RX As New ADODB.Recordset
                sSQL = "select Tipo from plaremunbase where cia = '" & wcia & "' and status != '*' and placod = '" & Trim(Txtcodpla.Text) & "' and concepto = '" & concepto & "'"
                If (fAbrRst(RX, sSQL)) Then sPeriodoPago = RX(0)
                If RX.State = 1 Then RX.Close
                
                If sPeriodoPago = "02" Then
                   'G01 = "select round(((((b.importe*4)/240)*240)/6)*(a.h21),2) as basico from platemphist a,plaremunbase b " 'solo para calculo con meses
                   G01 = "select round(((((b.importe*4)/240)*240)/6)*(a.h21) + ((((b.importe*4)/240))/6)*(8*a.h14) ,2)  as basico from platemphist a,plaremunbase b "
                   G01 = G01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Txtcodpla.Text & "' "
                   G01 = G01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
                   G01 = G01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & concepto & "' and b.status<>'*'"
                Else
                   'G01 = "select round((((b.importe/factor_horas)*240)/6)*(a.h21),2) as basico from platemphist a,plaremunbase b " 'solo para calculo con meses
                   G01 = "select round((((b.importe/factor_horas)*240)/6)*(a.h21) + (((b.importe/factor_horas))/6)*(8*a.h14) ,2)  as basico from platemphist a,plaremunbase b "
                   G01 = G01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Txtcodpla.Text & "' "
                   G01 = G01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
                   G01 = G01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & concepto & "' and b.status<>'*'"
                End If
           Case Is = "04" 'BONIFICACION T. SERVICIO
                'G01 = "select round((((b.importe/factor_horas)*240)/6)*(a.h21),2) as basico from platemphist a,plaremunbase b "
                G01 = "select round((((b.importe/factor_horas)*240)/6)*(a.h21) + (((b.importe/factor_horas))/6)*(8*a.h14) ,2) as basico from platemphist a,plaremunbase b "
                G01 = G01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Txtcodpla.Text & "' "
                G01 = G01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
                G01 = G01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & concepto & "' and b.status<>'*'"
           Case Is = "05" 'INCREMENTO AFP 10.23%
                'G01 = "select round((((b.importe/CASE WHEN TIPO='02' THEN factor_horas+8 ELSE FACTOR_HORAS END)*240)/6)*(a.h21),2) as basico from platemphist a,plaremunbase b "
                G01 = "select round((((b.importe/CASE WHEN TIPO='02' THEN factor_horas+8 ELSE FACTOR_HORAS END)*240)/6)*(a.h21) + (((b.importe/CASE WHEN TIPO='02' THEN factor_horas+8 ELSE FACTOR_HORAS END))/6)*(8*a.h14),2) as basico from platemphist a,plaremunbase b "
                G01 = G01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Txtcodpla.Text & "' "
                G01 = G01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
                G01 = G01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & concepto & "' and b.status<>'*'"
                
           Case Is = "06" 'INCREMENTO AFP 3%
                'G01 = "select round((((b.importe/CASE WHEN TIPO='02' THEN factor_horas+8 ELSE FACTOR_HORAS END)*240)/6)*(a.h21),2) as basico from platemphist a,plaremunbase b "
                G01 = "select round((((b.importe/CASE WHEN TIPO='02' THEN factor_horas+8 ELSE FACTOR_HORAS END)*240)/6)*(a.h21) + (((b.importe/CASE WHEN TIPO='02' THEN factor_horas+8 ELSE FACTOR_HORAS END))/6)*(8*a.h14),2) as basico from platemphist a,plaremunbase b "
                G01 = G01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Txtcodpla.Text & "' "
                G01 = G01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
                G01 = G01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & concepto & "' and b.status<>'*'"
                                
           Case Is = "07" 'BONIFICACION COSTO DE VIDA
                G01 = "select round((((b.importe/factor_horas)*240)/6)*(a.h21) + (((b.importe/factor_horas))/6)*(8*a.h14) ,2)  as basico from platemphist a,plaremunbase b "
                G01 = G01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Txtcodpla.Text & "' "
                G01 = G01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
                G01 = G01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & concepto & "' and b.status<>'*'"
                
           Case Is = "15" 'GRATIFICACION
                'G01 = "select round((((b.importe/factor_horas)*240)/6)*(a.h21),2) as basico from platemphist a,plaremunbase b "
                G01 = "select round((((b.importe/factor_horas)*240)/6)*(a.h21) + (((b.importe/factor_horas))/6)*(8*a.h14) ,2) as basico from platemphist a,plaremunbase b "
                G01 = G01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Txtcodpla.Text & "' "
                G01 = G01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
                G01 = G01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
                                
                
    End Select
Else
   Select Case concepto
          Case Is = "15" 'GRATIFICACION
               Sql$ = "select factor from platasaanexo where modulo='02' and tipomovimiento='" & Trim(Lbltope.Caption) & "' and codinterno='" & concepto & "' and tipotrab='" & vTipoTra & "' and status<>'*'"
               'Sql$ = "select factor from platasaanexo where modulo='02' and tipomovimiento='" & Trim(Lbltope.Caption) & "' and codinterno='" & concepto & "' and tipotrab='" & vTipoTra & "' and cargo='" & Trim(Lblcargo.Caption) & "' and status<>'*'"
               If (fAbrRst(rsG01, Sql$)) Then
                  mFactor = rsG01!factor
                  If rsG01.State = 1 Then rsG01.Close
                  G01 = "select round(((((b.importe/factor_horas)*8)* " & mFactor & ") / " & mh & ") * a.h14 ,2) as basico from platemphist a,plaremunbase b "
                  G01 = G01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Txtcodpla.Text & "' "
                  G01 = G01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
                  G01 = G01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
               End If
   End Select
End If
End Function
Private Sub Acumula_Mes(concepto As String, tipo As String)
Dim mcad As String
Dim SqlAcu As String
Dim RsAcumula As New Recordset
Dim rs2 As ADODB.Recordset
Dim mtb As String
Dim tope As Integer
macui = 0: macus = 0: macomi = 0
If VTipobol = "01" Or VTipobol = "02" Or VTipobol = "04" Then
    tope = 3
ElseIf VTipobol = "03" Then
    tope = 1
End If
For I = 1 To tope
    If VTipobol = "01" Then
        If I = 1 Then mtb = "01"
    ElseIf VTipobol = "03" Then
        If I = 1 Then mtb = "03"
    ElseIf VTipobol = "04" Then
        If I = 1 Then mtb = "01"
    End If
    'Modificacion 15/02/2010
     If VTipobol = "02" And concepto = "01" And tipo = "A" And I = 1 Then
       mtb = "01"
     End If
    'Fin Modificacion 15/02/2010
    
    If I = 2 Then mtb = "02"
    If I = 3 Then mtb = "04"
    'If i = 3 Then mtb = "03"
    If concepto = "06" And tipo = "D" Then
        Sql$ = "select 'D06' AS COD_REMU"
    Else
        Sql$ = "select cod_remu from plaafectos where cia='" & wcia & "' and tipo='" & tipo & "' and tboleta='" & mtb & "'  and  codigo='" & concepto & "' and status<>'*'"
    End If
    If (fAbrRst(RsAcumula, Sql$)) Then
       RsAcumula.MoveFirst
       mcad = ""
       Do While Not RsAcumula.EOF
        If concepto = "06" And tipo = "D" Then
            mcad = mcad & Trim(RsAcumula!cod_remu) & "+"
        Else
          mcad = mcad & "i" & Trim(RsAcumula!cod_remu) & "+"
        End If
          RsAcumula.MoveNext
       Loop
       
       mcad = "(" & Mid(mcad, 1, Len(Trim(mcad)) - 1) & ")"
       SqlAcu = "select sum(" & mcad & ") as ing,sum(" & Trim(tipo) & Trim(concepto) & ") as ded "
       SqlAcu = SqlAcu & "from plahistorico "
       SqlAcu = SqlAcu & "where cia='" & wcia & "' and status<>'*' and placod='" & Txtcodpla.Text & "' "
'       If concepto = "06" And tipo = "D" Then
'        SqlAcu = SqlAcu & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
'       Else
        SqlAcu = SqlAcu & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " and proceso='" & mtb & "'"
'       End If
'        macui = 0
'        macus = 0
       If (fAbrRst(rs2, SqlAcu)) Then
          If Not IsNull(rs2!ing) Then macui = macui + rs2!ing
          If Not IsNull(rs2!ded) Then macus = macus + rs2!ded
       End If
    End If
Next
If RsAcumula.State = 1 Then RsAcumula.Close
End Sub
Private Function Acumula_Mes_Afp(concepto As String, tipo As String) As String
Dim mcad As String
Dim SqlAcu As String
Dim RsAcumula As New Recordset

macui = 0: macus = 0
    Sql$ = "select cod_remu from plaafectos where cia='" & wcia & "' and tipo='" & tipo & "' and tboleta IN('01')  and  codigo='" & concepto & "' and status<>'*'"
    If (fAbrRst(RsAcumula, Sql$)) Then
       RsAcumula.MoveFirst
       mcad = ""
       Do While Not RsAcumula.EOF
          mcad = mcad & "i" & Trim(RsAcumula!cod_remu) & "+"
          RsAcumula.MoveNext
       Loop
       mcad = "(" & Mid(mcad, 1, Len(Trim(mcad)) - 1) & ")"
       SqlAcu = "select sum(" & mcad & ") as ing,sum(d111) as ded1, " _
              & "sum(d112) as ded2, " _
              & "sum(d113) as ded3, " _
              & "sum(d114) as ded4, " _
              & "sum(d115) as ded5 "
       SqlAcu = SqlAcu & "from plahistorico "
       SqlAcu = SqlAcu & "where cia='" & wcia & "' and status<>'*' and placod='" & Txtcodpla.Text & "' "
       SqlAcu = SqlAcu & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " and "
       SqlAcu = SqlAcu & "proceso in('01','02','04','05')"
       'SqlAcu = SqlAcu & "proceso in('01','02','03','04','05')"
    End If
If RsAcumula.State = 1 Then RsAcumula.Close
Acumula_Mes_Afp = SqlAcu
End Function
Private Sub Acumula_Mes_Afp112(concepto As String, tipo As String)
Dim mcad As String
Dim SqlAcu As String
Dim RsAcumula As New Recordset
Dim rs2 As ADODB.Recordset
Dim mtb As String
Dim tope As Integer
macui = 0: macus = 0: macomi = 0

If VTipobol = "01" Or VTipobol = "02" Then
    tope = 2
ElseIf VTipobol = "04" Then
   tope = 3
ElseIf VTipobol = "03" Then
    tope = 1
End If
For I = 1 To tope
'    If VTipobol = "02" And i <> 2 Then i = i + 1
'    If VTipobol <> "02" And i = 2 Then i = i + 1
'    If i > 3 Then Exit For
'    If i = 1 Then mtb = "01"
'    If i = 2 Then mtb = "02"
'    If i = 3 Then mtb = "03"
    Rem 2011.03.10
'    If VTipobol = "01" Then
'        If I = 1 Then mtb = "01"
'    ElseIf VTipobol = "03" Then
'        If I = 1 Then mtb = "03"
'    End If
    'If I = 2 Then mtb = "02"
    
    If VTipobol = "01" Then
        If I = 1 Then mtb = "01"
        If I = 2 Then mtb = "02"
    ElseIf VTipobol = "02" Then
        If I = 1 Then mtb = "01"
        If I = 2 Then mtb = "02"
    ElseIf VTipobol = "04" Then
        If I = 1 Then mtb = "01"
        If I = 2 Then mtb = "02"
        If I = 3 Then mtb = "04"
    ElseIf VTipobol = "03" Then
        If I = 1 Then mtb = "03"
    End If

    Sql$ = "select cod_remu from plaafectos where cia='" & wcia & "' and tipo='" & tipo & "' and tboleta='" & mtb & "'  and  codigo='" & concepto & "' and status<>'*'"
    If (fAbrRst(RsAcumula, Sql$)) Then
       RsAcumula.MoveFirst
       mcad = ""
       Do While Not RsAcumula.EOF
          mcad = mcad & "i" & Trim(RsAcumula!cod_remu) & "+"
          RsAcumula.MoveNext
       Loop
       mcad = "(" & Mid(mcad, 1, Len(Trim(mcad)) - 1) & ")"
       SqlAcu = "select sum(" & mcad & ") as ing,sum(d112) as ded "
       SqlAcu = SqlAcu & "from plahistorico "
       SqlAcu = SqlAcu & "where cia='" & wcia & "' and status<>'*' and placod='" & Txtcodpla.Text & "' "
       SqlAcu = SqlAcu & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " and proceso='" & mtb & "'"
       If (fAbrRst(rs2, SqlAcu)) Then
          If Not IsNull(rs2!ing) Then macui = macui + rs2!ing
          If Not IsNull(rs2!ded) Then macus = macus + rs2!ded
       End If
    End If
Next
If RsAcumula.State = 1 Then RsAcumula.Close
End Sub
Private Sub Acumula_Ano(concepto As String, tipo As String)
Dim mcad As String
Dim SqlAcu As String
Dim RsAcumula As New Recordset
Dim rs2 As ADODB.Recordset
Dim mtb As String
macui = 0: macus = 0: macuqtames = 0: macomi = 0
For I = 1 To 6
    If I = 1 Then mtb = "01"
    If I = 2 Then mtb = "02"
    If I = 3 Then mtb = "03"
    If I = 4 Then mtb = "04"
    If I = 5 Then mtb = "10"
    If I = 6 Then mtb = "11"
    
    'Sql$ = "select cod_remu from plaafectos where cia='" & wcia & "' and tipo='" & tipo & "' and tboleta='" & mtb & "'  and  codigo='" & concepto & "' and status<>'*'"
    Sql$ = "select cod_remu from plaafectos where cia='" & wcia & "' and tipo='" & tipo & "' and tboleta='01'  and  codigo='" & concepto & "' and status<>'*'"
    If (fAbrRst(RsAcumula, Sql$)) Then
       RsAcumula.MoveFirst
       mcad = ""
       Do While Not RsAcumula.EOF
          mcad = mcad & "i" & Trim(RsAcumula!cod_remu) & "+"
          RsAcumula.MoveNext
       Loop

       mcad = "(" & Mid(mcad, 1, Len(Trim(mcad)) - 1) & ")"
       'SqlAcu = "select sum(" & mcad & ") as ing,sum(" & Trim(tipo) & Trim(concepto) & ") as ded "
      If mtb = "03" And Vmes = 7 And concepto = "13" Then
         Dim rsF02 As ADODB.Recordset
         Dim F02ExtOrd As String
         mExtraOrd = 0: F02ExtOrd = ""
         Sql$ = "select codinterno as cod_remu from placonstante where cia='" & wcia & "' and tipomovimiento='02' and extraord='S' and status<>'*' and codinterno<>'30'"
         If (fAbrRst(rsF02, Sql$)) Then rsF02.MoveFirst
         Do While Not rsF02.EOF
               F02ExtOrd = F02ExtOrd & "i" & Trim(rsF02!cod_remu) & "+"
               rsF02.MoveNext
         Loop
         F02ExtOrd = "(" & Mid(F02ExtOrd, 1, Len(Trim(F02ExtOrd)) - 1) & ")"
         If rsF02.State = 1 Then rsF02.Close:
         If Len(Trim(F02ExtOrd)) > 2 Then
            mcad = "(" & mcad & " - " & F02ExtOrd & ")"
            F02ExtOrd = ""
         End If
      End If
       'agregado mgirao 30/10/2016
       If Vano = 2016 Then
        SqlAcu = "select (sum(i29) /count(*)) as PromedioComi from  plahistorico "
        SqlAcu = SqlAcu & "where cia='" & wcia & "' and status<>'*' and placod='" & Txtcodpla.Text & "' "
        SqlAcu = SqlAcu & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)<" & Vmes & " and proceso='" & mtb & "'and month(fechaproceso)>=9 group by placod"
       Else
        SqlAcu = "select (sum(i29) /count(*)) as PromedioComi from  plahistorico "
        SqlAcu = SqlAcu & "where cia='" & wcia & "' and status<>'*' and placod='" & Txtcodpla.Text & "' "
        SqlAcu = SqlAcu & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)<" & Vmes & " and proceso='" & mtb & "' group by placod"
       End If
       If (fAbrRst(rs2, SqlAcu)) Then
          If Not IsNull(rs2!PromedioComi) Then macomi = macomi + rs2!PromedioComi
       End If
       
        If RsAcumula.State = 1 Then RsAcumula.Close
       SqlAcu = "select sum(" & mcad & ") as ing,"
       SqlAcu = SqlAcu & "sum(case month(fechaproceso) when " & Vmes & " then 0 else " & Trim(tipo) & Trim(concepto) & " end) as ded,"
       If mtb = "03" And Vmes = 7 And concepto = "13" Then
          SqlAcu = SqlAcu & "0 as dedmes "
       Else
          SqlAcu = SqlAcu & "sum(case month(fechaproceso) when " & Vmes & " then " & Trim(tipo) & Trim(concepto) & " else 0 end) as dedmes "
       End If
       SqlAcu = SqlAcu & "from plahistorico "
       SqlAcu = SqlAcu & "where cia='" & wcia & "' and status<>'*' and placod='" & Txtcodpla.Text & "' "
       SqlAcu = SqlAcu & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)<=" & Vmes & " and proceso='" & mtb & "'"

       If (fAbrRst(rs2, SqlAcu)) Then
          If Not IsNull(rs2!ing) Then macui = macui + rs2!ing
          If Not IsNull(rs2!ded) Then macus = macus + rs2!ded
          If vTipoTra = "01" Then
             If Not IsNull(rs2!dedmes) Then macuqtames = macuqtames + rs2!dedmes
          Else
             If Not IsNull(rs2!dedmes) Then macus = macus + rs2!dedmes
          End If
          
       End If
    End If
Next
If RsAcumula.State = 1 Then RsAcumula.Close
End Sub
Private Function Busca_Grati() As Integer
Dim rsGrati As New Recordset
Select Case Vmes
       Case Is = 1, 2, 3, 4, 5, 6
            Busca_Grati = 2
       Case Is = 7
            Sql$ = "select placod from plahistorico where cia='" & wcia & "' and placod='" & Txtcodpla.Text & "' and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " " _
                 & "and proceso='03' and status<>'*'"
            If (fAbrRst(rsGrati, Sql$)) Then Busca_Grati = 1 Else Busca_Grati = 2
            rsGrati.Close
       Case Is = 8, 9, 10, 11
            Busca_Grati = 1
       Case Is = 12
            Sql$ = "select placod from plahistorico where cia='" & wcia & "' and placod='" & Txtcodpla.Text & "' and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " " _
                 & "and proceso='03' and status<>'*'"
            If (fAbrRst(rsGrati, Sql$)) Then Busca_Grati = 0 Else Busca_Grati = 1
            rsGrati.Close
End Select

End Function
Private Function Calc_Horas_FecIng(Inicio As String) As Currency
Dim mFIng As String
Dim mWorkNew As Boolean
Dim VHNew As String
Dim mDateBegin As String

Calc_Horas_FecIng = 0
mFIng = Mid(LblFingreso, 4, 2) & "/" & Left(LblFingreso, 2) & "/" & Right(LblFingreso, 4)
mWorkNew = False

If VNewBoleta = True Then
   If Compara_Fechas(mFIng, Inicio) = True Then
      mWorkNew = True
'      VHNew = InputBox("Primera Boleta del Trabajador" & Chr(13) & "Se Requiere que ingrese Horas Trabajadas", "Trabajador Nuevo")
'      Do While Not IsNumeric(VHNew)
'         VHNew = InputBox("Primera Boleta del Trabajador" & Chr(13) & "Se Requiere que ingrese Horas Trabajadas", "Trabajador Nuevo")
'      Loop
'      Do While VHNew > VHoras Or Val(VHNew) = 0
'         If VHNew >= VHoras Then MsgBox "Las Horas no deben ser mayores a " & Trim(Str(VHoras)), vbInformation, "Horas Trabajadas"
'         VHNew = "0"
'         VHNew = InputBox("Primera Boleta del Trabajador" & Chr(13) & "Se Requiere que ingrese Horas Trabajadas", "Trabajador Nuevo", VHNew)
'         If Not IsNumeric(VHNew) Then VHNew = "0"
'      Loop
'      VHorasnormal = VHNew
   End If
End If
If mWorkNew = True Then mDateBegin = mFIng Else mDateBegin = Inicio


'FERIADOS
Sql$ = ""
If VTipobol = "01" Then
     
   Select Case VPerPago
          Case Is = "04" 'Mensual
               Sql$ = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
               Sql$ = Sql$ & "select count(fecha) from plaferiados where cia='" & wcia & "' and fecha " _
                     & "BETWEEN '" & Format(mDateBegin, FormatFecha) + FormatTimei & "' AND '" & Format(VFProceso, FormatFecha) + FormatTimef & "' " _
                    & "and status<>'*'"
          Case Is = "02" 'Semanal
               Sql$ = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
               Sql$ = Sql$ & "select count(fecha) from plaferiados where cia='" & wcia & "' and fecha " _
                     & "BETWEEN '" & Format(mDateBegin, FormatFecha) + FormatTimei & "' AND '" & Format(VfAl, FormatFecha) + FormatTimef & "' " _
                    & "and status<>'*'"
   End Select
   If Trim(Sql$) <> "" Then If (fAbrRst(Rs, Sql$)) Then Rs.MoveFirst
   If Not IsNull(Rs(0)) Then
      If Rs(0) <> 0 Then
         VHorasnormal = VHorasnormal - (Rs(0) * mHourDay)
         Calc_Horas_FecIng = (Rs(0) * mHourDay)
      End If
   Else
   
   End If
End If

End Function
Private Sub Otros_Pagos_Vac(EsGrati As Boolean) 'comacsa

Dim mCadOtPag As String
Dim mPer As Integer
Dim mDateBeginVac As String
Dim mDateEndVac As String
Dim mdia As Integer
Dim meses As Integer
Dim RX As ADODB.Recordset
Dim FecIni As String, FecFin As String
'SE ELIMINO POR ERROR 16/02/2009
If Month(VFProceso) = 7 Then
    FecIni = "01/01/" & Year(VFProceso)
    FecFin = "07/01/" & Year(VFProceso)
Else
    FecIni = "06/01/" & Year(VFProceso)
    FecFin = "12/01/" & Year(VFProceso)
End If

mDateBeginVac = Fecha_Promedios(mPer, VFProceso)

If Val(Mid(VFProceso, 4, 2)) = 1 Then
   mdia = Ultimo_Dia(12, Val(Mid(VFProceso, 7, 4)) - 1)
   mDateEndVac = Format(mdia, "00") & "/12/" & Format(Val(Mid(VFProceso, 7, 4)) - 1, "0000")
Else
   mdia = Ultimo_Dia(Val(Mid(VFProceso, 4, 2) - 1), Val(Mid(VFProceso, 7, 4)))
   mDateEndVac = Format(mdia, "00") & "/" & Format(Val(Mid(VFProceso, 4, 2) - 1), "00") & "/" & Mid(VFProceso, 7, 4)
End If

mDateBeginVac = "01/" & Month(DateAdd("m", (6 - 1) * -1, mDateEndVac)) & "/" & Year(DateAdd("m", -5, mDateEndVac))

If Not EsGrati Then
   FecIni = Format(mDateBeginVac, "mm/dd/yyyy")
   FecFin = Format(mDateEndVac, "mm/dd/yyyy")
End If

mPer = 0

Dim lHorasExt As Boolean
If rspagadic.RecordCount > o Then rspagadic.MoveFirst
Do While Not rspagadic.EOF
   If rspagadic!Codigo = "16" Then
      Sql = "select distinct(codinterno),factor,factor_divisionario from platasaanexo where cia='" & wcia & "' and tipomovimiento='" & vTipoTra & "' and modulo='01'  and basecalculo='16' and status<>'*'"
      If (fAbrRst(Rs, Sql)) Then Rs.MoveFirst: mPer = Val(Rs(2)): meses = Val(Rs(1))
      mCadOtPag = ""
      Rs.MoveFirst
      Do While Not Rs.EOF
        If VTipobol = "03" Or VTipobol = "02" Then
              Sql = ""
              If Rs(0) = "10" Or Rs(0) = "11" Or Rs(0) = "21" Or Rs(0) = "25" Then 'Horas Extras
              'If rs(0) = "10" Or rs(0) = "11" Or rs(0) = "21" Or rs(0) = "24" Or rs(0) = "25" Then 'Horas Extras
                  Sql = "SET DATEFORMAT MDY SELECT COUNT(DISTINCT(MONTH(FECHAPROCESO))) FROM plahistorico WHERE cia='" & wcia & "' and placod='" & Txtcodpla.Text & "' and status!='*' and (fechaproceso>='" & FecIni & "' and fechaproceso<'" & FecFin & "') and I10+I11+I21+I25 >0 AND proceso='01'"
                 'Sql = "SET DATEFORMAT MDY SELECT COUNT(DISTINCT(MONTH(FECHAPROCESO))) FROM plahistorico WHERE cia='" & wcia & "' and placod='" & Txtcodpla.Text & "' and status!='*' and (fechaproceso>'" & FecIni & "' and fechaproceso<'" & FecFin & "') and I10+I11+I21+I24+I25 >0 AND proceso='01'"
              ElseIf Rs(0) = "08" Or Rs(0) = "22" Then 'Sobretasa
                 'Sql = "SET DATEFORMAT MDY SELECT COUNT(" & "I" & rs(0) & ") FROM plahistorico WHERE cia='" & wcia & "' and placod='" & Txtcodpla.Text & "' and status!='*' and (fechaproceso>'" & FecIni & "' and fechaproceso<'" & FecFin & "') and I08+I22 >0 AND proceso='01'"
                 Sql = "SET DATEFORMAT MDY SELECT COUNT(DISTINCT(MONTH(FECHAPROCESO))) FROM plahistorico WHERE cia='" & wcia & "' and placod='" & Txtcodpla.Text & "' and status!='*' and (fechaproceso>='" & FecIni & "' and fechaproceso<'" & FecFin & "') and I08+I22 >0 AND proceso='01'"
              Else
                 'Sql = "SET DATEFORMAT MDY SELECT COUNT(" & "I" & rs(0) & ") FROM plahistorico WHERE cia='" & wcia & "' and placod='" & Txtcodpla.Text & "' and status!='*' and (fechaproceso>'" & FecIni & "' and fechaproceso<'" & FecFin & "') and " & "I" & rs(0) & ">0 AND proceso='01'"
                 'Sql = "SET DATEFORMAT MDY SELECT COUNT(DISTINCT(MONTH(FECHAPROCESO))) FROM plahistorico WHERE cia='" & wcia & "' and placod='" & Txtcodpla.Text & "' and status!='*' and (fechaproceso>'" & FecIni & "' and fechaproceso<'" & FecFin & "') and " & "I" & rs(0) & ">0 AND proceso='01'"
                 'add corrige error no toma el primer dia del periodo
                 Sql = "SET DATEFORMAT MDY SELECT COUNT(DISTINCT(MONTH(FECHAPROCESO))) FROM plahistorico WHERE cia='" & wcia & "' and placod='" & Txtcodpla.Text & "' and status!='*' and (fechaproceso>='" & FecIni & "' and fechaproceso<'" & FecFin & "') and " & "I" & Rs(0) & ">0 AND proceso='01'"
              End If
              If (fAbrRst(RX, Sql)) Then
                  If RX(0) >= 3 Then
                       mCadOtPag = mCadOtPag & "I" & Rs(0) & "+"
                  End If
              End If
        Else
           mCadOtPag = mCadOtPag & "I" & Rs(0) & "+"
        End If
         Rs.MoveNext
      Loop
      Rs.Close
      Exit Do
   End If
   rspagadic.MoveNext
Loop

'AGREGAMOS LAS HORAS EXTRAS QUE NO FUESEN CARGADAS
If VTipobol <> "03" Then
   If InStr(1, mCadOtPag, "10") > 0 Or InStr(1, mCadOtPag, "11") > 0 Or InStr(1, mCadOtPag, "21") > 0 Or InStr(1, mCadOtPag, "24") > 0 Or InStr(1, mCadOtPag, "25") > 0 Then
       If InStr(1, mCadOtPag, "10") = 0 Then mCadOtPag = mCadOtPag & "I10+"
       If InStr(1, mCadOtPag, "11") = 0 Then mCadOtPag = mCadOtPag & "I11+"
       If InStr(1, mCadOtPag, "21") = 0 Then mCadOtPag = mCadOtPag & "I21+"
      ' If InStr(1, mCadOtPag, "24") = 0 Then mCadOtPag = mCadOtPag & "I24+"
       If InStr(1, mCadOtPag, "25") = 0 Then mCadOtPag = mCadOtPag & "I25+"
   End If
End If

Sql = ""

'mDateBeginVac = Fecha_Promedios(mPer, VFProceso)
'
'If Val(Mid(VFProceso, 4, 2)) = 1 Then
'   mdia = Ultimo_Dia(12, Val(Mid(VFProceso, 7, 4)) - 1)
'   mDateEndVac = Format(mdia, "00") & "/12/" & Format(Val(Mid(VFProceso, 7, 4)) - 1, "0000")
'Else
'   mdia = Ultimo_Dia(Val(Mid(VFProceso, 4, 2) - 1), Val(Mid(VFProceso, 7, 4)))
'   mDateEndVac = Format(mdia, "00") & "/" & Format(Val(Mid(VFProceso, 4, 2) - 1), "00") & "/" & Mid(VFProceso, 7, 4)
'End If
'
'mDateBeginVac = "01/" & Month(DateAdd("m", (meses - 1) * -1, mDateEndVac)) & "/" & Year(DateAdd("m", -5, mDateEndVac))

If Trim(mCadOtPag) <> "" Then
   mCadOtPag = Mid(mCadOtPag, 1, Len(Trim(mCadOtPag)) - 1)
   mCadOtPag = "sum(" & mCadOtPag & ")"
   
   Dim FECINICIO As String
   Dim FECFINAL As String
                
   FECINICIO = Format(mDateBeginVac, "mm/dd/yyyy")
   FECFINAL = Format(mDateEndVac, "mm/dd/yyyy")
                    
    Sql = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " select "
    Sql = Sql & "CASE WHEN dbo.FC_VALIDAPAGOS('" & wcia & "','" & Txtcodpla.Text & "','" & FECINICIO & "','" & FECFINAL & "')>=3 THEN "
    Sql = Sql & mCadOtPag & "else 0 end from plahistorico"

End If

If Trim(Sql) <> "" Then
   Sql = Sql & " where cia='" & wcia & "' and fechaproceso Between '" & Format(mDateBeginVac, FormatFecha) & Space(1) & FormatTimei & "' and '" & Format(mDateEndVac, FormatFecha) & Space(1) & FormatTimef & "'"
   Sql = Sql & " and proceso='01' and placod='" & Txtcodpla.Text & "' and status<>'*'"

   If (fAbrRst(Rs, Sql)) Then Rs.MoveFirst
   
   
      If EsGrati Then
      Dim gMeses As Integer
      Dim gDias As Integer
      Dim G01 As String
      Dim OtrGrati As Double
      Dim d_resultado As Double
      
      G01 = "select  h21,h14 from platemphist "
      G01 = G01 & "where cia='" & wcia & "' and status='" & Mstatus & "' and placod='" & Txtcodpla.Text & "' "
      G01 = G01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
      
      gDias = 0: gMeses = 0: OtrGrati = 0
      
      Dim Rsg As ADODB.Recordset
      If fAbrRst(Rsg, G01) Then
         gMeses = Rsg!h21: gDias = Rsg!h14
      End If
      If Rsg.State = 1 Then Rsg.Close: Set Rsg = Nothing
      If Not IsNull(Rs(0)) Then OtrGrati = Rs(0) / 6
      
      d_resultado = Round(((OtrGrati / 6) * (gMeses)) + ((((OtrGrati / 240) / 6)) * (8 * gDias)), 2)
      
      
            
      ''actualizar en el temporal
      G01 = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
      G01 = G01 & " Update platemphist set I16 =" & d_resultado
      G01 = G01 & " where cia='" & wcia & "' and placod='" & Txtcodpla.Text & "' and codauxinterno='" & Lblcodaux.Caption & "' and proceso='" & VTipobol & "' "
      G01 = G01 & "and fechaproceso='" & Format(VFProceso, FormatFecha) & "' and semana='" & VSemana & "' and status='" & Mstatus & "'"
      cn.Execute G01
      
      If Not IsNull(Rs(0)) Then rspagadic!Monto = d_resultado
           
   Else
      If Not IsNull(Rs(0)) Then rspagadic!Monto = Rs(0) / mPer
   End If
         

End If
End Sub
Private Function Calc_Horas_Quincena(Inicio As String) As Currency
Dim mFIng As String
Dim mWorkNew As Boolean
Dim VHNew As String

Calc_Horas_Quincena = 0
mFIng = Mid(LblFingreso, 4, 2) & "/" & Left(LblFingreso, 2) & "/" & Right(LblFingreso, 4)
mWorkNew = False
If wruc = "20100037689" Then 'Agregados no se calcula quincena se ingresa
Else
   VHNew = InputBox("Adelanto de Quincena" & Chr(13) & "Se Requiere que ingrese Horas Trabajadas", "Adelanto de Quincena")
   Do While Not IsNumeric(VHNew)
      VHNew = InputBox("Adelanto de Quincena" & Chr(13) & "Se Requiere que ingrese Horas Trabajadas", "Adelanto de Quincena")
   Loop
   Do While VHNew > VHoras Or Val(VHNew) = 0
      If VHNew >= VHoras Then MsgBox "Las Horas no deben ser mayores a " & Trim(Str(VHoras)), vbInformation, "Horas Trabajadas"
      VHNew = "0"
      VHNew = InputBox("Adelanto de Quincena" & Chr(13) & "Se Requiere que ingrese Horas Trabajadas", "Adelanto de Quincena", VHNew)
     If Not IsNumeric(VHNew) Then VHNew = "0"
   Loop
   Calc_Horas_Quincena = VHNew
End If
End Function
Private Sub Descarga_ctaCte(Codigo As String, tipobol As String, fecha As String, sem As String, tiptrab As String, importe As Currency)
Dim Sql As String
Dim RX As ADODB.Recordset
      Quincena = 0
      For I = 0 To MAXROW - 1
        If ArrDsctoCTACTE(0, I) = Trim(Txtcodpla.Text) Then
            If ArrDsctoCTACTE(2, I) > 0 Then
            
                Sql = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
                Sql = Sql & "insert into plabolcte values('" & wcia & "','" & UCase(Trim(Codigo)) & "','" & tipobol & "','', " _
                & "'" & fecha & "','" & sem & "','" & tiptrab & "','" & Lblcodaux.Caption & "','" & ArrDsctoCTACTE(3, I) & "','" & fecha & "', " _
                & "'" & wmoncont & "'," & ArrDsctoCTACTE(2, I) & ",'','" & wuser & "'," & FechaSys & "," & ArrDsctoCTACTE(1, I) & "," & IIf(wtipodoc = True, 0, 1) & ")"
                
                cn.Execute Sql
            End If
        
            Sql = "UPDATE plactacte set pago_acuenta=pago_acuenta + " & CStr(ArrDsctoCTACTE(2, I)) + " WHERE cia='" & wcia & "' and placod='" & Trim(Txtcodpla.Text) & "' and id_doc='" & CStr(ArrDsctoCTACTE(1, I)) & "' and status<>'*'"
            cn.Execute Sql
            
            Sql = "UPDATE plactacte set fecha_cancela='" & fecha & "' WHERE cia='" & wcia & "' and placod='" & Trim(Txtcodpla.Text) & "' and id_doc='" & ArrDsctoCTACTE(1, I) & "' and status<>'*' AND IMPORTE=PAGO_ACUENTA"
            cn.Execute Sql
                
        End If
      Next
End Sub

Private Sub Calcula_Devengue_Vaca()
Dim mNumBol As Integer
Dim I As Integer
Dim rsdevengue As ADODB.Recordset
Dim NroTrans As Integer
On Error GoTo Salir

VokDevengue = True
Sql = "select placod,recordacu from plaprovvaca where cia='" & wcia & "' and year(fechaproceso)=" & Vano - 1 & " and month(fechaproceso)=12 and status<>'*'"
If (fAbrRst(rsdevengue, Sql)) Then rsdevengue.MoveFirst
Panelprogress.Visible = True
Panelprogress.ZOrder 0
Me.Refresh
Barra.Value = 0
Barra.Max = rsdevengue.RecordCount
Do While Not rsdevengue.EOF
   Barra.Value = rsdevengue.AbsolutePosition
   mNumBol = Val(Mid(rsdevengue!recordacu, 1, 4))
   If mNumBol > 0 Then
      NumDev = 0
      For I = 1 To mNumBol
          NumDev = NumDev + 1
          Txtcodpla.Text = rsdevengue!PlaCod
          Txtcodpla_LostFocus
          Grabar_Boleta
      Next
   End If
   rsdevengue.MoveNext
Loop
rsdevengue.Close
Panelprogress.Visible = False
If VokDevengue = False Then
    'cn.BeginTrans
    Sql$ = wInicioTrans & " DEVENGUE_VACA"
    cn.Execute Sql$

   NroTrans = 1
   
   Sql = "update plahistorico set status='*',fec_modi=" & FechaSys & ",user_modi='" & wuser & "' " _
        & "where cia='" & wcia & "' and proceso='02' and Year(fechaproceso) = " & Vano & " And month( fechaproceso) = " & Vmes & " " _
        & "and status='D'"
   cn.Execute Sql
   
   'cn.CommitTrans
    Sql$ = wFinTrans & " DEVENGUE_VACA"
    cn.Execute Sql$

   NroTrans = 0
   MsgBox "Se detectaron Irregularidades " & Chr(13) & "Se cancelara el calculo", vbCritical, "Devengue de Vacaciones"
End If
Unload Me
Frmprovision.Provisiones ("D")
Frmprovision.Txtano.Text = Vano
Frmprovision.Cmbmes.ListIndex = Vmes - 1
Frmprovision.Show
Frmprovision.ZOrder 0

Exit Sub
Salir:
If NroTrans = 1 Then
    'cn.RollbackTrans
    Sql$ = wCancelTrans & " DEVENGUE_VACA"
    cn.Execute Sql$
End If

MsgBox Err.Description, vbCritical, Me.Caption
End Sub
Private Sub Grabar_Devengada()
Dim NroTrans As Integer
On Error GoTo ErrorTrans
NroTrans = 0

Mgrab = MsgBox("Seguro de Grabar Boleta", vbYesNo + vbQuestion, TitMsg)
If Mgrab <> 6 Then Exit Sub
Screen.MousePointer = vbArrowHourglass

'cn.BeginTrans
Sql$ = wInicioTrans & " GRABAR_DEVENGUE"
cn.Execute Sql$

NroTrans = 1
Sql = "select * from plahistorico where cia='" & wcia & "' and proceso='02' and year(fechaproceso)<=" & Val(Mid(VFProceso, 7, 4)) & " and month(fechaproceso)=1 and status='D' and placod='" & Txtcodpla.Text & "' order by turno"
If (fAbrRst(Rs, Sql)) Then Rs.MoveFirst
Sql = "update plahistorico set status='T' where cia='" & Rs!cia & "' and placod='" & Rs!PlaCod & "' and status='" & Rs!status & "' and turno='" & Rs!turno & "'"
cn.Execute Sql
Rs.Close

'cn.CommitTrans
Sql$ = wFinTrans & " GRABAR_DEVENGUE"
cn.Execute Sql$

Limpia_Boleta
Screen.MousePointer = vbDefault
Exit Sub
ErrorTrans:
If NroTrans = 1 Then
    'cn.RollbackTrans
    Sql$ = wCancelTrans & " GRABAR_DEVENGUE"
    cn.Execute Sql$
End If
MsgBox Err.Description, vbCritical, Me.Caption

Limpia_Boleta
Screen.MousePointer = vbDefault

End Sub

Sub llena_horas()
    Dim RZ As New ADODB.Recordset
    Dim Rt As New ADODB.Recordset
    Dim con As String, cod As String
    Dim horas As Variant, VALOR As Variant
    Dim I As Integer, ccosto As String, PORC As String
    Dim INGRESOS As String
            
    horas = ""
    For I = 1 To 30
         horas = horas & "H" & Format(I, "00") & ","
    Next

    horas = Left(horas, Len(horas) - 1)

    'PAGOS ADICIONALES
    If rspagadic.RecordCount > 0 Then
       rspagadic.MoveFirst
       Do While Not rspagadic.EOF
                INGRESOS = INGRESOS & "I" & rspagadic(0) & ","
                rspagadic.MoveNext
       Loop

       INGRESOS = Left(INGRESOS, Len(INGRESOS) - 1)
    End If
    
    'DESCUENTOS
    If rsdesadic.RecordCount > 0 Then
       rsdesadic.MoveFirst
       Do While Not rsdesadic.EOF
                DESCUENTOS = DESCUENTOS & "d" & rsdesadic(0) & ","
                rsdesadic.MoveNext
       Loop

       DESCUENTOS = Left(DESCUENTOS, Len(DESCUENTOS) - 1)
    End If
        
    ccosto = "ccosto1,ccosto2,ccosto3,ccosto4,ccosto5"
    PORC = "PORC1,PORC2,PORC3,PORC4,PORC5"
    
    con = "select " & horas & "," & ccosto & "," & PORC & "," & INGRESOS & _
    "," & DESCUENTOS & " from plahistorico where cia='" & wcia & "' and " & "proceso='" & _
    VTipobol & "' and placod='" & Trim(Txtcodpla.Text) & "' and semana='" & _
    VSemana & "' and Year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " and day(fechaproceso)=" & Vdia & " and status<>'*'"

    RZ.Open con, cn, adOpenStatic, adLockReadOnly
    
    If RZ.RecordCount = 0 Then Exit Sub
    
    BoletaCargada = True
    
    rshoras.MoveFirst
    Do While Not rshoras.EOF
      With rshoras
           cod = "h" & Format(.Fields.Item(0), "00")
'           If .Fields.Item(0) <> 14 Then
            VALOR = RZ.Fields.Item(cod).Value
'           Else
'            VALOR = RZ.Fields.Item(cod).Value / 8
'           End If
           .Fields.Item(2) = IIf(VALOR = 0, Null, VALOR)
              
           .MoveNext
      End With
    Loop
  
  'LLENA CENTRO DE COSTO
  con = "Select cod_maestro2,descrip from maestros_2 where status<>'*' " & _
  "and ciamaestro= '01044' ORDER BY cod_maestro2"
  Rt.Open con, cn, adOpenStatic, adLockReadOnly
  
    For I = 0 To 4
        con = "ccosto" & (I + 1)
        'rsccosto.AddNew
        VALOR = RZ.Fields.Item("PORC" & (I + 1))
        If VALOR <> 0 Then
'           rsccosto.Fields("MONTO") = VALOR
'
'           VALOR = RZ.Fields.Item(con)
'           Rt.Filter = "cod_maestro2='" & VALOR & "'"
'           If Not Rt.EOF Then
'            rsccosto.Fields("CODIGO") = Trim(Rt("COD_MAESTRO2"))
'            rsccosto.Fields("DESCRIPCION") = Rt("DESCRIP")
'           End If
'           rsccosto.MoveNext
        End If
    Next
  
    Rt.Close
    'RZ.Close
    'Set RZ = Nothing
    
    
    'LLENA PAGOSADICIONALES
    rspagadic.MoveFirst
    Do While Not rspagadic.EOF
          VALOR = "I" & rspagadic.Fields("CODIGO")
          rspagadic.Fields("MONTO") = RZ.Fields(VALOR)
          rspagadic.MoveNext
    Loop
    
    'llena descuentos
    rsdesadic.MoveFirst
    Do While Not rsdesadic.EOF
       VALOR = "d" & rsdesadic.Fields("CODIGO")
       rsdesadic.Fields("MONTO") = RZ.Fields(VALOR)
       rsdesadic.MoveNext
    Loop
    
End Sub

Sub MUESTRA_CUENTACORRIENTE()
Dim RX As New ADODB.Recordset
Dim I As Integer
Dim sumporc As Currency
On Error GoTo CORRIGE
MAXROW = 0
'VSemana
'CON = "SELECT DES.MONTO FROM PLACTACTE CTA,PLADESCTA DES WHERE " & _
'"CTA.PLACOD='" & Trim(Txtcodpla.Text) & "' AND (CTA.IMPORTE-CTA.PAGO_ACUENTA)>0 AND " & _
'"DES.CODAUXINTERNO=CTA.CODAUXINTERNO AND RIGHT(DES.FECHA,2)='" & _
'VSemana & "' AND CTA.STATUS<>'*'"

If BoletaCargada Then Exit Sub

If VTipobol = "01" Then
    con = "SELECT a.id_doc,a.tipo,a.importe,COALESCE(SUM(b.importe),0) as pago_acuenta,a.partes,"
    con = con & " a.partes-(SELECT COALESCE(SUM(importe),0) FROM plabolcte WHERE cia=a.cia and placod=a.placod and id_doc=a.id_doc and status!='*' AND sn_quincena=0"
    con = con & " AND MONTH(fechaproceso)=" & Month(CDate(VFProceso)) & " AND YEAR(fechaproceso)=" & Year(CDate(VFProceso)) & ") as saldomes,"
    con = con & "(SELECT COALESCE(SUM(importe),0) FROM plabolcte WHERE cia=a.cia and placod=a.placod and id_doc=a.id_doc and status!='*' AND sn_quincena=1"
    con = con & " AND MONTH(fechaproceso)=" & Month(CDate(VFProceso)) & " AND YEAR(fechaproceso)=" & Year(CDate(VFProceso)) & ") as quincena"
    con = con & " FROM PLACTACTE a LEFT OUTER JOIN PLABOLCTE b ON"
    con = con & " (b.cia=a.cia and b.placod=a.placod and b.id_doc=a.id_doc and b.status!='*') WHERE a.cia='" & wcia & "' and a.placod='" & Trim(Txtcodpla.Text) & "' AND a.STATUS<>'*'"
    con = con & " GROUP BY a.cia,a.placod,a.id_doc,a.tipo,a.importe,a.partes"
Else
    con = "SELECT a.id_doc,a.tipo,a.importe,COALESCE(SUM(b.importe),0) as pago_acuenta,a.partes,"
    con = con & " a.partes-(SELECT COALESCE(SUM(importe),0) FROM plabolcte WHERE cia=a.cia and placod=a.placod and id_doc=a.id_doc and status!='*' AND sn_quincena=0"
    con = con & " AND MONTH(fechaproceso)=" & Month(CDate(VFProceso)) & " AND YEAR(fechaproceso)=" & Year(CDate(VFProceso)) & ") as saldomes,"
    con = con & "(SELECT COALESCE(SUM(importe),0) FROM plabolcte WHERE cia=a.cia and placod=a.placod and id_doc=a.id_doc and status!='*' AND sn_quincena=1"
    con = con & " AND MONTH(fechaproceso)=" & Month(CDate(VFProceso)) & " AND YEAR(fechaproceso)=" & Year(CDate(VFProceso)) & ") as quincena"
    con = con & " FROM PLACTACTE a LEFT OUTER JOIN PLABOLCTE b ON"
    con = con & " (b.cia=a.cia and b.placod=a.placod and b.id_doc=a.id_doc and b.status!='*') WHERE a.cia='" & wcia & "' and a.placod='" & Trim(Txtcodpla.Text) & "' AND a.STATUS<>'*' and a.sn_grati=1"
    con = con & " GROUP BY a.cia,a.placod,a.id_doc,a.tipo,a.importe,a.partes"
End If

RX.Open con, cn, adOpenStatic, adLockReadOnly

   If RX.RecordCount > 0 Then

      rsdesadic.MoveFirst
      Erase ArrDsctoCTACTE
      MAXROW = 0
      Do While Not rsdesadic.EOF
         If rsdesadic("CODIGO") = "07" Then
            Do While Not RX.EOF
                ReDim Preserve ArrDsctoCTACTE(0 To 4, 0 To MAXROW)
                
                ArrDsctoCTACTE(0, MAXROW) = Trim(Txtcodpla.Text)
                ArrDsctoCTACTE(1, MAXROW) = RX!id_doc
                ArrDsctoCTACTE(3, MAXROW) = RX!tipo
                ArrDsctoCTACTE(4, MAXROW) = 0
                
                If RX("partes") >= (RX("IMPORTE") - RX("PAGO_ACUENTA")) Then
                   rsdesadic("MONTO") = rsdesadic("MONTO") + (RX("IMPORTE") - RX("PAGO_ACUENTA"))
                   If wtipodoc = False Then ArrDsctoCTACTE(2, MAXROW) = Round((RX("IMPORTE") - RX("PAGO_ACUENTA")) / 2, 2) Else ArrDsctoCTACTE(2, MAXROW) = (RX("IMPORTE") - RX("PAGO_ACUENTA"))
                Else
                   rsdesadic("MONTO") = rsdesadic("MONTO") + IIf(RX("partes") = 0, 0, RX("partes") - RX("QUINCENA"))
                   If wtipodoc = False Then ArrDsctoCTACTE(2, MAXROW) = Round((RX("partes") - RX("QUINCENA")) / 2, 2) Else ArrDsctoCTACTE(2, MAXROW) = RX("partes") - RX("QUINCENA")
                End If
                
'                If wTipoDoc = False Then
'                    If rsdesadic("MONTO") > 0 Then
'                        rsdesadic("MONTO") = Round(rsdesadic("MONTO") / 2, 2)
'                    End If
'                End If
                
                MAXROW = MAXROW + 1
                
                RX.MoveNext
            Loop
            
             If wtipodoc = False Then
                If rsdesadic("MONTO") > 0 Then
                    rsdesadic("MONTO") = Round(rsdesadic("MONTO") / 2, 2)
                End If
            End If
            
            If rsdesadic("MONTO") < 0 Then
                rsdesadic("MONTO") = 0
            End If
            
            sumporc = 0
            If rsdesadic("MONTO") > 0 Then
                For I = 0 To MAXROW - 1
                
                    ArrDsctoCTACTE(4, I) = Round((ArrDsctoCTACTE(2, I) * 100) / rsdesadic("MONTO"), 2)
                    sumporc = sumporc + Round((ArrDsctoCTACTE(2, I) * 100) / rsdesadic("MONTO"), 2)
                    If I = MAXROW - 1 Then
                        If sumporc > 100 Then
                            ArrDsctoCTACTE(4, I) = ArrDsctoCTACTE(4, I) - (sumporc - 100)
                        ElseIf sumporc < 100 Then
                            ArrDsctoCTACTE(4, I) = ArrDsctoCTACTE(4, I) + (100 - sumporc)
                        End If
                    End If
                Next
            End If
         End If
         rsdesadic.MoveNext
      Loop
   End If

RX.Close


Exit Sub
CORRIGE:
MsgBox "Error :" & Err.Description, vbCritical, Me.Caption

End Sub

Public Function convierte_cant() As Double
Dim RX As New ADODB.Recordset
Dim con As String
Dim importe As Double

con = "select plar.* from plaremunbase plar,platemphist plat where " & _
"plar.placod=plat.placod and plar.concepto='04' and plar.status<>'*'"
RX.Open con, cn, adOpenStatic, adLockReadOnly

'rx.Filter = "concepto='04'"
If RX.RecordCount = 0 Then Exit Function

con = RX("factor_horas")

If con <> 8 Then
   importe = RX("importe") / 30
Else
   importe = RX("importe")
End If

RX.Close

convierte_cant = importe

End Function
Private Sub NuevaBoleta()
Dim rs2 As ADODB.Recordset

On Error GoTo CORRIGE

BolDevengada = False
sn_quinta = True
bImportar = False
cn.CursorLocation = adUseClient

Set Rs = New ADODB.Recordset

Sql$ = " EXEC sp_c_datos_personal '" & wcia & "','" & Trim(Txtcodpla.Text) & "'"

Set Rs = cn.Execute(Sql$)

If Not Rs.EOF Then
   If Rs!TipoTrabajador <> vTipoTra Then
      If bImportar = False Then
         MsgBox "Trabajador no es del tipo seleccionado", vbExclamation, "Codigo N° => " & Txtcodpla.Text
      Else
         LstObs.AddItem Trim(Txtcodpla.Text) & ": Trabajador no es del tipo seleccionado"
      End If
      Txtcodpla.Text = ""
      Limpia_Boleta
      Exit Sub
   End If

   If Not IsNull(Rs!fcese) Then
      If CDate(VFProceso) > CDate(Rs!fcese) Then
          If bImportar = False Then
               MsgBox "Trabajador " & Trim(Txtcodpla.Text) & " ya fue Cesado", vbExclamation, "Con Fecha => " & Format(Rs!fcese, "dd/mm/yyyy")
          Else
               LstObs.AddItem Trim(Txtcodpla.Text) & ": Trabajador ya fue Cesado"
          End If
         Txtcodpla.Text = ""
         Limpia_Boleta
         Exit Sub
      End If
   End If

   LblFingreso.Caption = Format(Rs!fIngreso, "mm/dd/yyyy")
    LblFechaIngreso.Caption = Format(Rs!fIngreso, "dd/mm/yyyy")
   If Val(Right(LblFingreso.Caption, 4)) > Vano Then
      MsgBox "Fecha de Ingreso del Trabajador es Superior", vbCritical, "Calculo de Boletas"
      Limpia_Boleta
      LblFingreso.Caption = ""
      Exit Sub
   ElseIf Val(Right(LblFingreso.Caption, 4)) = Vano And Val(Left(LblFingreso.Caption, 2)) > Vmes Then
      MsgBox "Fecha de Ingreso del Trabajador es Superior", vbCritical, "Calculo de Boletas"
      Limpia_Boleta
      LblFingreso.Caption = ""
      Exit Sub
   ElseIf Val(Right(LblFingreso.Caption, 4)) = Vano And Val(Left(LblFingreso.Caption, 2)) = Vmes And Val(Mid(LblFingreso.Caption, 4, 2)) > Val(Left(VFProceso, 2)) Then
      MsgBox "Fecha de Ingreso del Trabajador es Superior", vbCritical, "Calculo de Boletas"
      Limpia_Boleta
      LblFingreso.Caption = ""
      Exit Sub
   End If
      
   'cargamos si la person esta afecto o no a la quinta categoria
   sn_quinta = True
   If Trim(Rs!quinta & "") = "N" Then sn_quinta = False
      '***********************************************************
      
   Lblnombre.Caption = Rs!nombre
   Lblcodaux.Caption = Rs!codauxinterno
   Lblcodafp.Caption = Rs!CodAfp
   LblCodTrabSunat.Caption = Trim(Rs!CodTipoTrabSunat & "")
   Lblnumafp.Caption = Trim(Rs!NUMAFP)
   VFechaNac = Format(Rs!fnacimiento, "dd/mm/yyyy")
   VFechaJub = Format(Rs!fec_jubila, "dd/mm/yyyy")
   VJubilado = Trim(Rs!jubilado)
   Lbltope.Caption = Rs!tipotasaextra
      
   If vTipoTra = "05" Then Lblcargo.Caption = Rs!Cargo: VAltitud = Rs!altitud: VVacacion = Rs!vacacion
      
   VArea = Trim(Rs!Area)
   Frame2.Enabled = True
   Frame3.Enabled = True
   Frame4.Enabled = True
   Frame5.Enabled = True
      
   LblBasico.Caption = fCadNum(Rs!Basico, "###,##0.00")
      
   If rsccosto.RecordCount > 0 Then rsccosto.MoveFirst
   Do While Not rsccosto.EOF
      rsccosto.Delete
      rsccosto.MoveNext
   Loop
      
   If Trim(UCase(Rs("ESSALUDVIDA"))) = "S" Then
      rsdesadic.MoveFirst
      rsdesadic.FIND "CODIGO='06'", 1 'ESSALUDVIDA
      If Not rsdesadic.EOF Then rsdesadic("MONTO") = Rs!impessalud
   End If
        
   MSINDICATO = False
           
   If Trim(UCase(Rs("sindicato"))) = "S" Then
      rsdesadic.FIND "CODIGO='20'", 1 'SOLIC SINDICATO
      If Not rsdesadic.EOF Then rsdesadic("MONTO") = Rs!impsindicato
         
      MSINDICATO = True
   End If
      
   rsccosto.AddNew
   rsccosto.MoveFirst
   rsccosto!Codigo = Trim(Rs!Area)
   rsccosto!Descripcion = UCase(Rs!DESCRIP)
   rsccosto!Monto = "100.00"
   lbltot.Caption = "100.00"
   rsccosto!Item = vItem
      
   For I = I To 4
      If rsccosto.RecordCount < 5 Then rsccosto.AddNew
   Next I
      
   rsccosto.MoveFirst
   Dgrdccosto.Refresh
            
   Rs.Close
   Txtcodpla.Enabled = False
ElseIf Trim(Txtcodpla.Text) <> "" Then
   MsgBox "Codigo de Planilla No Existe ", vbExclamation, "Codigo N° => " & Txtcodpla.Text
   Txtcodpla.Text = ""
   Limpia_Boleta
   Lblnombre.Caption = ""
   Lblctacte.Caption = "0.00"
   Lblcodaux.Caption = ""
   Lblcodafp.Caption = ""
   LblCodTrabSunat.Caption = ""
   Lblnumafp.Caption = ""
   LblBasico.Caption = ""
   LblFingreso.Caption = ""
   Lbltope.Caption = ""
   Lblcargo.Caption = ""
   VAltitud = ""
   VVacacion = ""
   VArea = ""
   VFechaNac = ""
   VFechaJub = ""
   vPlaCod = ""
   Txtcodpla.SetFocus
   Frame2.Enabled = False
   Frame3.Enabled = False
   Frame4.Enabled = False
   Frame5.Enabled = False
End If
vItem = 0
If VTipobol <> "01" Then
'   Dgrdhoras.Columns(1).Locked = False
'Else
   'Dgrdhoras.Columns(1).Locked = True
End If

Call Carga_Horas_NEW

If VTipobol = "02" Then Otros_Pagos_Vac (False)
If wtipodoc = True And (VTipobol = "01" Or VTipobol = "05") Then
   If rsdesadic.RecordCount > 0 Then rsdesadic.MoveFirst
   Do While Not rsdesadic.EOF
      If rsdesadic!Codigo = "09" Then
         Sql$ = "select totneto from plaquincena where cia='" & wcia & "' and placod='" & Txtcodpla.Text & "'  and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " and status<>'*' "
          '       Sql$ = "select 0"
         'mgirao
          If (fAbrRst(rs2, Sql$)) Then
            rsdesadic!Monto = rs2(0)
         End If
    
         rs2.Close
      End If
      rsdesadic.MoveNext
   Loop
End If
If wtipodoc = True Then
   Sql = "select sum(importe-pago_acuenta) from plactacte where cia='" & wcia & "' and placod='" & Trim(Txtcodpla.Text) & "' and importe-pago_acuenta<>0 and status<>'*' and importe>0"
   If (fAbrRst(Rs, Sql$)) Then
      If IsNull(Rs(0)) Then Lblctacte.Caption = "0.00" Else Lblctacte.Caption = Format(Rs(0), "###,###,###.00")
   End If
   Rs.Close
End If

If VTipobol = "02" And vDevengue = False Then
   Sql = "select i16 from plahistorico where cia='" & wcia & "' and proceso='02' and year(fechaproceso)<=" & Val(Mid(VFProceso, 7, 4)) & " and month(fechaproceso)=1 and status='D' and placod='" & Txtcodpla.Text & "' order by turno"
   If (fAbrRst(Rs, Sql$)) Then
      Rs.MoveFirst
      If rspagadic.RecordCount > 0 Then rspagadic.MoveFirst
      Do While Not rspagadic.EOF
         rspagadic.Delete
         rspagadic.MoveNext
      Loop
      If Rs!I16 <> 0 Then
         rspagadic.AddNew
         rspagadic!Codigo = "16"
         rspagadic!Monto = Rs!I16
         rspagadic!Descripcion = "OTROS PAGOS"
      End If
      MsgBox "Trabajador Tiene Vacaciones Devengadas" & Chr(13) & "No podra modificar Datos, Solo Grabar", vbInformation, "Vacaciones Devengadas"
      BolDevengada = True
      Frame2.Enabled = False
      Frame3.Enabled = False
      Frame4.Enabled = False
      Frame5.Enabled = False
   End If
End If
If VTipobol = "01" Or VTipobol = "03" Or VTipobol = "04" Then MUESTRA_CUENTACORRIENTE
Exit Sub
CORRIGE:
  MsgBox "Error:" & Err.Description, vbCritical, Me.Caption
End Sub

Private Sub Carga_Horas_NEW()
Dim rs2 As ADODB.Recordset
Dim mbol As String
Dim mconceptos As String
Dim mhor As Currency
Dim mdiasfalta As String
Dim mdiasferiado As String
Dim wBeginMonth As String
Dim mHourTra As Currency
Dim con As String
Dim I As Integer
Dim NroTrans As Integer

On Error GoTo CORRIGE
NroTrans = 0
If Trim(Txtcodpla.Text) = "" Then Exit Sub

Sql$ = "select iniciomes from cia where cod_cia='" & _
wcia & "' and status<>'*'"

If (fAbrRst(Rs, Sql$)) Then
   If IsNull(Rs!iniciomes) Then wBeginMonth = "1" Else wBeginMonth = Rs!iniciomes
End If
Rs.Close

If Trim(wBeginMonth) <> "1" Then
   If Vmes = 1 Then
      wBeginMonth = Format(wBeginMonth, "00") & "/12/" & Format(Vano - 1, "0000")
   Else
     wBeginMonth = Format(wBeginMonth, "00") & "/" & Format(Vmes - 1, "00") & "/" & Format(Vano, "0000")
   End If
Else
   wBeginMonth = Format(wBeginMonth, "00") & "/" & Format(Vmes, "00") & "/" & Format(Vano, "0000")
End If

If VPerPago = "02" Then wBeginMonth = VfDel

If rshoras.RecordCount > 0 Then rshoras.MoveFirst
Do While Not rshoras.EOF
   rshoras.Delete
   rshoras.MoveNext
Loop
 
'cn.BeginTrans
Sql$ = wInicioTrans & " CARGA_HORAS_NEW"
cn.Execute Sql$

NroTrans = 1
Sql$ = "delete from platemphist where cia='" & wcia & "' and placod='" & Trim(Txtcodpla.Text) & "'"

cn.Execute Sql
'cn.CommitTrans
Sql$ = wFinTrans & " CARGA_HORAS_NEW"
cn.Execute Sql$
NroTrans = 0

mdiasfalta = 0
mdiasferiado = 0
VHorasnormal = VHoras
If VTipobol = "01" Then mbol = "N"
If VTipobol = "02" Then mbol = "V"
If VTipobol = "03" Then mbol = "G"

wciamae = Funciones.Determina_Maestro("01077")

If VTipobol = "01" Or wtipodoc <> True Then

   Sql$ = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
    
   If wtipodoc = True Then
     
      Select Case VPerPago
             Case Is = "04" 'Mensual
                  Sql$ = Sql$ & "select distinct(concepto) from platareo where cia='" & wcia & "' and fecha " _
                       & "BETWEEN '" & Format(wBeginMonth, FormatFecha) + Coneccion.FormatTimei & _
                       "' AND '" & Format(VFProceso, FormatFecha) + FormatTimef & "' " _
                       & "and codigotrab='" & Trim(Txtcodpla) & "' and status<>'*'"
                       
             Case Is = "02" 'Semanal
                  Sql$ = Sql$ & "select distinct(concepto) from platareo where cia='" & wcia & "' and fecha " _
                       & " BETWEEN '" & Format(VfDel, FormatFecha) + FormatTimei & "' AND '" & Format(VfAl, FormatFecha) + FormatTimef & "' " _
                       & "and codigotrab='" & Trim(Txtcodpla) & "' and status<>'*'"
      End Select
   Else
      Sql$ = Sql$ & "select distinct(concepto) from platareo where cia='" & wcia & "' and fecha " _
           & "BETWEEN '" & Format(wBeginMonth, FormatFecha) + FormatTimei & _
           "' AND '" & Format(VFProceso, FormatFecha) + FormatTimef & "' " _
           & "and codigotrab='" & Trim(Txtcodpla) & "' and status<>'*'"
      mbol = "N"
   End If
   
   If (fAbrRst(Rs, Sql$)) Then
      Rs.MoveFirst
      mconceptos = ""
      Do While Not Rs.EOF
         mconceptos = mconceptos & "'" & Trim(Rs!concepto) & "',"
         Rs.MoveNext
      Loop
      mconceptos = "(cod_maestro2 in (" & Mid(mconceptos, 1, Len(mconceptos) - 1) & ")"
   End If
   
   If Rs.State = 1 Then Rs.Close
   
   If Trim(mconceptos) <> "" Then
      Sql$ = "Select * from maestros_2 where status<>'*' and CHARINDEX('" & mbol & "',flag1 )>0 and " & mconceptos
   Else
      Sql$ = "Select * from maestros_2 where status<>'*' and CHARINDEX('" & mbol & "',flag1 )>0 and "
   End If
   
   If vTipoTra = "01" Then
      If Trim(mconceptos) <> "" Then
         Sql$ = Sql$ & " or cod_maestro2 in('01'))"
      Else
         Sql$ = Sql$ & "cod_maestro2 in('01')"
      End If
   Else
     If Trim(mconceptos) <> "" Then
         Sql$ = Sql$ & " or cod_maestro2 in ('14','01','02','03','04','06','05','07','08','09','10','17','11','12','13','15','18','19'))"
     Else
        Sql$ = Sql$ & "cod_maestro2 in  ('14','01','02','03','04','06','05','07','08','09','10','17','11','12','13','15','18','19')"
     End If
   End If
   
'3     '-feriado
'4     'perm. pag.
'6     '-enferm. no pagadas
'5     '-enferm. pagadas
'7     'accidente de trabajo
'8     'faltas injustificadas
'9     'suspencion
'10    'extras l-s
'11    'extras d-f
'12    'vacaciones
'13    'sobretasa
'15    'otros
   
   
   Sql$ = Sql$ & wciamae
   
   If (fAbrRst(Rs, Sql$)) Then
      Rs.MoveFirst
      con = "140102030406050708091017111213151819"
      
       
      For I = 1 To (Len(con) / 2)
      'Do While Not RS.EOF
         Rs.Filter = "COD_MAESTRO2='" & Mid(con, I + (I - 1), 2) & "'"
        If Not Rs.EOF Then  ' {<MA>} 01/02/2007
            rshoras.AddNew
            rshoras!Codigo = Trim(Rs!cod_maestro2)
            rshoras!Descripcion = UCase(Rs!DESCRIP)
            If Trim(Rs("cod_maestro2")) = "14" Then
               rshoras("MONTO") = 6
            End If

            If vTipoTra = "02" Then
                If rshoras!Codigo = "01" Then
                    'rshoras!Monto = 6 * 8
                    rshoras!Monto = 7 * 8
                End If
            
                If rshoras!Codigo = "02" Then
                    rshoras!Monto = (6 / 6) * 8
                End If
            End If
            
         mhor = 0
         
         Sql$ = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
                  
         Select Case VPerPago
               Case Is = "04"
                    Sql$ = Sql$ & "select concepto,motivo,sum(tiempo) as tiempo from platareo where cia='" & wcia & "' and fecha " _
                         & "BETWEEN '" & Format(wBeginMonth, FormatFecha) + FormatTimei & "' AND '" & Format(VFProceso, FormatFecha) + FormatTimef & "' " _
                         & "and codigotrab='" & Trim(Txtcodpla) & "' and concepto='" & _
                         Trim(Rs!cod_maestro2) & "' and status<>'*'  group by concepto,motivo order by concepto,motivo"
               Case Is = "02"
                    Sql$ = Sql$ & "select concepto,motivo,sum(tiempo) as tiempo from platareo where cia='" & wcia & "' fecha " _
                         & " BETWEEN '" & Format(VfDel, FormatFecha) + FormatTimei & "' AND '" & _
                         Format(VfAl, FormatFecha) + FormatTimef & "' " _
                         & "and codigotrab='" & Trim(Txtcodpla) & "' and concepto='" & Trim(Rs!cod_maestro2) & "' and status<>'*'  group by concepto,motivo order by concepto,motivo"
         End Select
 
         If (fAbrRst(rs2, Sql$)) Then
            rs2.MoveFirst
            Do While Not rs2.EOF
               Select Case rs2!motivo
                      Case Is = "DI"
                            mhor = mhor + (rs2!tiempo * 8 * 60)
                            mdiasfalta = mdiasfalta + rs2!tiempo
                            If Rs!cod_maestro2 = "03" Then mdiasferiado = mdiasferiado + rs2!tiempo
                      Case Is = "HO"
                           mhor = mhor + Int(rs2!tiempo) * 60 + ((rs2!tiempo - Int(rs2!tiempo)) * 100)
                      Case Is = "MI"
                           mhor = mhor + rs2!tiempo
               End Select
               rs2.MoveNext
            Loop
            If rs2.State = 1 Then rs2.Close
            mhor = Int(mhor / 60) + ((mhor Mod 60) / 100)
                        
            rshoras!Monto = IIf(mhor = 0, Null, mhor)
            
            If Trim(Rs!flag2) = "-" Then
               VHorasnormal = VHorasnormal - mhor
            End If
         End If
         
         If rshoras("CODIGO") = "14" And IsNull(rshoras("MONTO")) Then
            rshoras("MONTO") = 6
         End If
         
         Rs.MoveNext
      End If
      Next
      
      'TIPICO
       
      rshoras.MoveFirst
      mHourTra = 0
      If rshoras!Codigo = "01" Then
         If wtipodoc = True Then
            mHourTra = Calc_Horas_FecIng(wBeginMonth)
            rshoras!Monto = VHorasnormal
         Else
            VHorasnormal = 0
            If VNewBoleta = True Then VHorasnormal = Calc_Horas_Quincena(wBeginMonth)
            rshoras!Monto = VHorasnormal
         End If
         If rshoras.RecordCount > 1 Then rshoras.MoveNext
      End If
      If rshoras!Codigo = "02" Then rshoras!Monto = ((6 - mdiasfalta + mdiasferiado) * 8) / 6
      If mHourTra <> 0 Then
         rshoras.AddNew
         rshoras!Codigo = "03"
         rshoras!Descripcion = "FERIADOS"
         rshoras!Monto = mHourTra
      End If
      
   End If
Else
    Sql$ = "Select * from maestros_2 where status<>'*' and CHARINDEX('" & mbol & "',flag1 )>0"
    Sql$ = Sql$ & wciamae
    cn.CursorLocation = adUseClient
    Set Rs = New ADODB.Recordset
    Set Rs = cn.Execute(Sql$, 64)
    If rshoras.RecordCount > 0 Then
       rshoras.MoveFirst
       Do While Not rshoras.EOF
          rshoras.Delete
          rshoras.MoveNext
       Loop
    End If
    If Not Rs.RecordCount > 0 Then MsgBox "No Existen Horas Registradas", vbCritical, TitMsg: Exit Sub
    Rs.MoveFirst
    Do While Not Rs.EOF
       rshoras.AddNew
       rshoras!Codigo = Trim(Rs!cod_maestro2)
       rshoras!Descripcion = Trim(Rs!DESCRIP)
       If VTipobol = "02" Then rshoras!Monto = VHoras Else rshoras!Monto = "0.00"
       Rs.MoveNext
    Loop
    If Rs.State = 1 Then Rs.Close
End If

'solo para datos guardados
Call Llena_Horas_New

Exit Sub
CORRIGE:

If NroTrans = 1 Then
    'cn.RollbackTrans
    Sql$ = wCancelTrans & " CARGA_HORAS_NEW"
    cn.Execute Sql$
End If
MsgBox "Error :" & Err.Description, vbCritical, "Sistema de Planillas"
End Sub


Sub Llena_Horas_New()
    Dim RZ As New ADODB.Recordset
    Dim Rt As New ADODB.Recordset
    Dim con As String, cod As String
    Dim horas As Variant, VALOR As Variant
    Dim I As Integer, ccosto As String, PORC As String
    Dim INGRESOS As String
        
    For I = 1 To 30
         horas = horas & "H" & Format(I, "00") & ","
    Next
    horas = Left(horas, Len(horas) - 1)
    
    'PAGOS ADICIONALES
    If rspagadic.RecordCount > 0 Then
       rspagadic.MoveFirst
       Do While Not rspagadic.EOF
                INGRESOS = INGRESOS & "I" & rspagadic(0) & ","
                rspagadic.MoveNext
       Loop
       INGRESOS = Left(INGRESOS, Len(INGRESOS) - 1)
    End If
    
    'DESCUENTOS
    If rsdesadic.RecordCount > 0 Then
       rsdesadic.MoveFirst
       Do While Not rsdesadic.EOF
                DESCUENTOS = DESCUENTOS & "d" & rsdesadic(0) & ","
                rsdesadic.MoveNext
       Loop
       DESCUENTOS = Left(DESCUENTOS, Len(DESCUENTOS) - 1)
    End If
        
    ccosto = "ccosto1,ccosto2,ccosto3,ccosto4,ccosto5"
    PORC = "PORC1,PORC2,PORC3,PORC4,PORC5"
    
    con = "select " & horas & "," & ccosto & "," & PORC & "," & INGRESOS & _
    "," & DESCUENTOS & " from plahistorico where cia='" & wcia & "' and " & "proceso='" & _
    VTipobol & "' and placod='" & Trim(Txtcodpla.Text) & "' and semana='" & _
    VSemana & "' and Year(fechaproceso)=" & Vano & " and status<>'*'"

    RZ.Open con, cn, adOpenStatic, adLockReadOnly
    
    If RZ.RecordCount = 0 Then Exit Sub
    
    rshoras.MoveFirst
    Do While Not rshoras.EOF
      With rshoras
           cod = "h" & Format(.Fields.Item(0), "00")
           If .Fields.Item(0) <> 14 Then
            VALOR = RZ.Fields.Item(cod).Value
           Else
            VALOR = RZ.Fields.Item(cod).Value / 8
           End If
           .Fields.Item(2) = IIf(VALOR = 0, Null, VALOR)
              
           .MoveNext
      End With
    Loop
  
  'LLENA CENTRO DE COSTO
  con = "Select cod_maestro2,descrip from maestros_2 where status<>'*' " & _
  "and ciamaestro= '01044' ORDER BY cod_maestro2"
  Rt.Open con, cn, adOpenStatic, adLockReadOnly
  
    For I = 0 To 4
        con = "ccosto" & (I + 1)
        'rsccosto.AddNew
        VALOR = RZ.Fields.Item("PORC" & (I + 1))
        If VALOR <> 0 Then
           rsccosto.Fields("MONTO") = VALOR
           
           VALOR = RZ.Fields.Item(con)
           Rt.Filter = "cod_maestro2='" & VALOR & "'"
           rsccosto.Fields("CODIGO") = Trim(Rt("COD_MAESTRO2"))
           rsccosto.Fields("DESCRIPCION") = Rt("DESCRIP")
           rsccosto.MoveNext
        End If
    Next
  
    Rt.Close
    'RZ.Close
    'Set RZ = Nothing
    
    
    'LLENA PAGOSADICIONALES
    rspagadic.MoveFirst
    Do While Not rspagadic.EOF
          VALOR = "I" & rspagadic.Fields("CODIGO")
          rspagadic.Fields("MONTO") = RZ.Fields(VALOR)
          rspagadic.MoveNext
    Loop
    
    'llena descuentos
    rsdesadic.MoveFirst
    Do While Not rsdesadic.EOF
       VALOR = "d" & rsdesadic.Fields("CODIGO")
       rsdesadic.Fields("MONTO") = RZ.Fields(VALOR)
       rsdesadic.MoveNext
    Loop
    
End Sub


Private Sub ProrrateaCtaCte(ByVal pNuevoValor As Currency)
Dim x As Integer

For x = 0 To MAXROW - 1
    ArrDsctoCTACTE(2, x) = Round(pNuevoValor * (ArrDsctoCTACTE(4, x) / 100), 2)
Next x

End Sub
Private Sub Acumula_Mes_SCTR(concepto As String, tipo As String)
Dim mcad As String
Dim SqlAcu As String
Dim rs2 As New Recordset
Dim RsAcumula As New Recordset

macui = 0: macus = 0: macomi = 0
    Sql$ = "select cod_remu from plaafectos where cia='" & wcia & "' and tipo='" & tipo & "' and tboleta='01'  and  codigo='" & concepto & "' and status<>'*'"
    If (fAbrRst(RsAcumula, Sql$)) Then
       RsAcumula.MoveFirst
       mcad = ""
       Do While Not RsAcumula.EOF
          mcad = mcad & "i" & Trim(RsAcumula!cod_remu) & "+"
          RsAcumula.MoveNext
       Loop
       mcad = "(" & Mid(mcad, 1, Len(Trim(mcad)) - 1) & ")"
       
       SqlAcu = "select sum(" & mcad & ") as ing,sum(a" & concepto & ") as ded "
       SqlAcu = SqlAcu & "from plahistorico "
       SqlAcu = SqlAcu & "where cia='" & wcia & "' and status<>'*' and placod='" & Txtcodpla.Text & "' "
       SqlAcu = SqlAcu & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " and "
       SqlAcu = SqlAcu & "proceso in('01','02','04','05')"
       'SqlAcu = SqlAcu & "proceso in('01','02','03','04','05')"
       
       If (fAbrRst(rs2, SqlAcu)) Then
          If Not IsNull(rs2!ing) Then macui = macui + rs2!ing
          If Not IsNull(rs2!ded) Then macus = macus + rs2!ded
       End If
    End If
If rs2.State = 1 Then rs2.Close

End Sub

Private Function Bcodsctr(per As String)
Dim Sql As String
Dim RZ As New ADODB.Recordset

    Sql = "select p.codsctr " & _
    " from planillas p " & _
    " where p.status<>'*'  " & _
    " and p.cia='" & wcia & "' and p.placod='" & per & "' "
    
    RZ.Open Sql, cn, adOpenStatic, adLockReadOnly
    
    If RZ.RecordCount = 0 Then
      Bcodsctr = 0
    Else
      Bcodsctr = RZ!codsctr
    End If
RZ.Close
End Function

Private Sub Otros_Pagos_Vac_Ant()

Dim mCadOtPag As String
Dim mPer As Integer
Dim mDateBeginVac As String
Dim mDateEndVac As String
Dim mdia As Integer
Dim meses As Integer
Dim RX As ADODB.Recordset
Dim FecIni As String, FecFin As String


If Month(VFProceso) = 7 Then
    FecIni = "01/01/" & Year(VFProceso)
    FecFin = "07/01/" & Year(VFProceso)
Else
    FecIni = "07/01/" & Year(VFProceso)
    FecFin = "12/01/" & Year(VFProceso)
End If

mPer = 0
If rspagadic.RecordCount > 0 Then rspagadic.MoveFirst
Do While Not rspagadic.EOF
   If rspagadic!Codigo = "16" Then
      Sql = "select distinct(codinterno),factor,factor_divisionario from platasaanexo where cia='" & wcia & "' and tipomovimiento='" & vTipoTra & "' and modulo='01'  and basecalculo='16' and status<>'*'"
      If (fAbrRst(Rs, Sql)) Then Rs.MoveFirst: mPer = Val(Rs(2)): meses = Val(Rs(1))
      mCadOtPag = ""
      Rs.MoveFirst
      Do While Not Rs.EOF
        If VTipobol = "03" Then
              Sql = "SET DATEFORMAT MDY SELECT COUNT(" & "I" & Rs(0) & ") FROM plahistorico WHERE cia='" & wcia & "' and placod='" & Txtcodpla.Text & "' and status!='*' and (fechaproceso>'" & FecIni & "' and fechaproceso<'" & FecFin & "') and " & "I" & Rs(0) & ">0 AND proceso='01'"
              If (fAbrRst(RX, Sql)) Then
                  If RX(0) >= 3 Then
                       mCadOtPag = mCadOtPag & "I" & Rs(0) & "+"
                  End If
              End If
        Else
           mCadOtPag = mCadOtPag & "I" & Rs(0) & "+"
        End If
         Rs.MoveNext
      Loop
      Rs.Close
      Exit Do
   End If
   rspagadic.MoveNext
Loop

'AGREGAMOS LAS HORAS EXTRAS QUE NO FUESEN CARGADAS
If InStr(1, mCadOtPag, "10") > 0 Or InStr(1, mCadOtPag, "11") > 0 Or InStr(1, mCadOtPag, "21") > 0 Then
    If InStr(1, mCadOtPag, "10") = 0 Then mCadOtPag = mCadOtPag & "I10+"
    If InStr(1, mCadOtPag, "11") = 0 Then mCadOtPag = mCadOtPag & "I11+"
    If InStr(1, mCadOtPag, "21") = 0 Then mCadOtPag = mCadOtPag & "I21+"
End If

Sql = ""

mDateBeginVac = Fecha_Promedios(mPer, VFProceso)

If Val(Mid(VFProceso, 4, 2)) = 1 Then
   mdia = Ultimo_Dia(12, Val(Mid(VFProceso, 7, 4)) - 1)
   mDateEndVac = Format(mdia, "00") & "/12/" & Format(Val(Mid(VFProceso, 7, 4)) - 1, "0000")
Else
   mdia = Ultimo_Dia(Val(Mid(VFProceso, 4, 2) - 1), Val(Mid(VFProceso, 7, 4)))
   mDateEndVac = Format(mdia, "00") & "/" & Format(Val(Mid(VFProceso, 4, 2) - 1), "00") & "/" & Mid(VFProceso, 7, 4)
End If

mDateBeginVac = "01/" & Month(DateAdd("m", (meses - 1) * -1, mDateEndVac)) & "/" & Year(DateAdd("m", -5, mDateEndVac))

If Trim(mCadOtPag) <> "" Then
   mCadOtPag = Mid(mCadOtPag, 1, Len(Trim(mCadOtPag)) - 1)
   mCadOtPag = "sum(" & mCadOtPag & ")"
   '**********codigo modificado giovanni 17092007******************************
   'sql = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " select " & mCadOtPag & " from plahistorico"
   
   '***************************************************************************
   'FECHA DE MODIFICACION : 08/01/2008
   'MODIFICADO POR RICARDO HINOSTROZA
   'SE MODOFICA AÑO DE PROCESO QUE SE INGRESA A F_CARGAPAGOS2
   'sql = sql & "'" & Val(Month(mDateBeginVac)) & "','" & Val(Month(mDateEndVac)) & "','" & Val(Year(FecFin)) & "')=3 THEN "
   'REEMPLAZADO POR
   'sql = sql & "'" & Val(Month(mDateBeginVac)) & "','" & Val(Month(mDateEndVac)) & "','" & PAÑO_PROCESO & "')=3 THEN "
   '***************************************************************************
   
   Dim PAÑO_PROCESO As Integer
   
   If Month(FecFin) >= 1 Or Month(FEFIN) < 7 Then
    If VTipobol = "03" Then
    PAÑO_PROCESO = Val(Year(FecFin))
    Else
    PAÑO_PROCESO = Val(Year(FecFin)) - 1
    End If
        
   Else
        PAÑO_PROCESO = Val(Year(FecFin))
   End If
        
        
   '***MODIFICADO 03/07/2009
'   Sql = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " select "
'   Sql = Sql & "CASE WHEN dbo.fc_cargapagos2('" & wcia & "','" & Txtcodpla.Text & "',"
'   Sql = Sql & "'" & Val(Month(mDateBeginVac)) & "','" & Val(Month(mDateEndVac)) & "','" & PAÑO_PROCESO & "')=3 THEN "
'   Sql = Sql & mCadOtPag & "else 0 end from plahistorico "
    Dim FECINICIO As String
    Dim FECFINAL As String
                
    FECINICIO = Format(mDateBeginVac, "mm/dd/yyyy")
    FECFINAL = Format(mDateEndVac, "mm/dd/yyyy")
                    
    Sql = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " select "
    Sql = Sql & "CASE WHEN dbo.FC_VALIDAPAGOS('" & wcia & "','" & Txtcodpla.Text & "','" & FECINICIO & "','" & FECFINAL & "')>=3 THEN "
    Sql = Sql & mCadOtPag & "else 0 end from plahistorico"

End If

If Trim(Sql) <> "" Then

   Sql = Sql & " where cia='" & wcia & "' and fechaproceso Between '" & Format(mDateBeginVac, FormatFecha) & Space(1) & FormatTimei & "' and '" & Format(mDateEndVac, FormatFecha) & Space(1) & FormatTimef & "'"
   Sql = Sql & " and proceso='01' and placod='" & Txtcodpla.Text & "' and status<>'*'"
   If (fAbrRst(Rs, Sql)) Then Rs.MoveFirst
 '  If wcia = "05" Then
 '   If Not IsNull(rs(0)) Then rspagadic!Monto = rs(0) / 6
 '  Else
    If Not IsNull(Rs(0)) Then rspagadic!Monto = Rs(0) / mPer
'   End If
End If
End Sub

Private Sub TxtHor1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Sendkeys "{tab}"
        KeyAscii = 0
    Else
        If (KeyAscii >= 46 And KeyAscii <= 57 Or KeyAscii = 127 Or KeyAscii = 8) And TxtAct1.Text <> "" Then
            ' 46 es . El 48 es 0 y el 57 es 9, 127 es SUPR y 8 es Backspace
            Exit Sub
        Else
             MsgBox "Solo números para registrar las horas sin puntos, " & _
                    "ni comas, ni cualquier caracter especial y el activo debe estar Ingresado!!"
             KeyAscii = 0
        End If
    End If
End Sub

Private Sub TxtHor2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Sendkeys "{tab}"
        KeyAscii = 0
    Else
        If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 127 Or KeyAscii = 8 Then
            ' El 48 es 0 y el 57 es 9, 127 es SUPR y 8 es Backspace
            Exit Sub
        Else
             MsgBox "Solo números para registrar las horas sin puntos, " & _
                    "ni comas, ni cualquier caracter especial!!"
        End If
    End If
End Sub

Private Sub TxtHor3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Sendkeys "{tab}"
        KeyAscii = 0
    Else
        If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 127 Or KeyAscii = 8 Then
            ' El 48 es 0 y el 57 es 9, 127 es SUPR y 8 es Backspace
            Exit Sub
        Else
            MsgBox "Solo números para registrar las horas sin puntos, " & _
                   "ni comas, ni cualquier caracter especial!!"
        End If
    End If
End Sub
Private Sub TxtHor4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Sendkeys "{tab}"
        KeyAscii = 0
    Else
        If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 127 Or KeyAscii = 8 Then
            ' El 48 es 0 y el 57 es 9, 127 es SUPR y 8 es Backspace
            Exit Sub
        Else
            MsgBox "Solo números para registrar las horas sin puntos, " & _
                   "ni comas, ni cualquier caracter especial!!"
        End If
    End If
End Sub
Private Sub TxtRango_Change()
    Dim I As Integer
    I = Len(TxtRango.Text)
    TxtRango.Text = UCase(TxtRango.Text)
    TxtRango.SelStart = I
End Sub
Private Sub Agrega_Promedios(xBoleta As String, esliquid As String)
Dim mDateBeginVac As String
Dim mDateEndVac As String
Dim mCalculaExtras As Boolean
Dim mPer As Integer
Dim meses As Integer
Dim mcad As String
Dim RX As ADODB.Recordset
Dim mImporte As Double

If esliquid = "S" Then
   Dim Vano As Integer
   Dim Vmes As Integer
   Dim Vdia As Integer
   Dim gano As Integer
   Dim gmes As Integer
   Dim gdia As Integer
   Dim cano As Integer
   Dim cmes As Integer
   Dim cdia As Integer

   Vano = 0: Vmes = 0: Vdia = 0
   gano = 0: gmes = 0: gdia = 0
   cano = 0: cmes = 0: cdia = 0

   RsLiquid.MoveFirst
   Do While Not RsLiquid.EOF
      If RsLiquid!tipo = "02" Then Vano = RsLiquid!ano: Vmes = RsLiquid!Mes: Vdia = RsLiquid!Dia:
      If RsLiquid!tipo = "03" Then gano = RsLiquid!ano: gmes = RsLiquid!Mes: gdia = RsLiquid!Dia:
      If RsLiquid!tipo = "06" Then cano = RsLiquid!ano: cmes = RsLiquid!Mes: cdia = RsLiquid!Dia:
      RsLiquid.MoveNext
   Loop
   
   Sql$ = "Usp_platmphisliquid " & Vano & "," & Vmes & "," & Vdia & "," & gano & "," & gmes & "," & gdia & "," & cano & "," & cmes & "," & cdia & ""
   cn.Execute Sql
End If

Sql = "select distinct(codinterno),factor,factor_divisionario from platasaanexo where cia='" & wcia & "' and tipomovimiento='" & vTipoTra & "' and modulo='01'  and basecalculo='16' and status<>'*'"
If (fAbrRst(Rs, Sql)) Then Rs.MoveFirst: mPer = Val(Rs(2)): meses = Val(Rs(1))

mDateBeginVac = Fecha_Promedios(mPer, VFProceso)
If Val(Mid(VFProceso, 4, 2)) = 1 Then
   mdia = Ultimo_Dia(12, Val(Mid(VFProceso, 7, 4)) - 1)
   mDateEndVac = Format(mdia, "00") & "/12/" & Format(Val(Mid(VFProceso, 7, 4)) - 1, "0000")
Else
   mdia = Ultimo_Dia(Val(Mid(VFProceso, 4, 2) - 1), Val(Mid(VFProceso, 7, 4)))
   mDateEndVac = Format(mdia, "00") & "/" & Format(Val(Mid(VFProceso, 4, 2) - 1), "00") & "/" & Mid(VFProceso, 7, 4)
End If
mDateBeginVac = "01/" & Format(Month(DateAdd("m", (meses - 1) * -1, mDateEndVac)), "00") & "/" & Format(Year(DateAdd("m", -5, mDateEndVac)), "0000")

Dim FECINICIO As String
Dim FECFINAL As String
                
FECINICIO = Mid(mDateBeginVac, 4, 2) & "/" & Mid(mDateBeginVac, 1, 2) & "/" & Mid(mDateBeginVac, 7, 4)
FECFINAL = Mid(mDateEndVac, 4, 2) & "/" & Mid(mDateEndVac, 1, 2) & "/" & Mid(mDateEndVac, 7, 4)

mCalculaExtras = False
Sql = "SET DATEFORMAT mdy select dbo.fc_validaHEextras('" & wcia & "','" & Txtcodpla.Text & "','" & FECINICIO & "','" & FECFINAL & "')"
If (fAbrRst(RX, Sql)) Then
   If Not IsNull(RX(0)) Then If RX(0) >= 3 Then mCalculaExtras = True
End If
RX.Close: Set RX = Nothing

mcad = ""
Do While Not Rs.EOF
   mcad = ""
   Select Case Rs(0)
   Case "10", "11", "21", "24", "25"
        If mCalculaExtras = True Then mcad = mcad & "I" & Rs(0)
   Case Else
        Sql = "SET DATEFORMAT MDY "
        Sql = Sql & "select count(1) from "
        Sql = Sql & "(SELECT COUNT(1) as c FROM plahistorico "
        Sql = Sql & "WHERE cia='" & wcia & "' and placod='" & Txtcodpla.Text & "' and status!='*' "
        Sql = Sql & "and (fechaproceso>'" & FECINICIO + " 00:00:00" & "' and fechaproceso<'" & FECFINAL + " 23:59:59" & "') "
        Sql = Sql & "and " & "I" & Rs(0) & ">0 AND proceso='01' "
        Sql = Sql & "group by month(fechaproceso)) Temp"
        If (fAbrRst(RX, Sql)) Then
           If RX(0) >= 3 Then mcad = mcad & "I" & Rs(0)
        End If
        RX.Close: Set RX = Nothing
   End Select
   mImporte = 0
   If Trim(mcad) <> "" Then
      Sql = "SET DATEFORMAT MDY select sum(" & mcad & ") from plahistorico "
      Sql = Sql & " where cia='" & wcia & "' and fechaproceso Between '" & FECINICIO + " 00:00:00" & "' and '" & FECFINAL + "  23:59:59" & "'"
      Sql = Sql & " and proceso='01' and placod='" & Txtcodpla.Text & "' and status<>'*'"
      If (fAbrRst(RX, Sql)) Then RX.MoveFirst
      If Not IsNull(RX(0)) Then mImporte = RX(0) / mPer
      RX.Close: Set RX = Nothing
   End If
   If mImporte > 0 Then
      If esliquid <> "S" Then
         Sql = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
         Sql = Sql & "Update platemphist set " & mcad & "" & " = " & "" & mcad & "" & " + " & "" & mImporte & ""
         Sql = Sql & " where cia='" & wcia & "' and placod='" & _
         Trim(Txtcodpla.Text) & "' and codauxinterno='" & _
         Trim(Lblcodaux.Caption) & "' and proceso='" & _
         Trim(xBoleta) & "' and fechaproceso='" & _
         Format(VFProceso, FormatFecha) & _
         "' and semana='" & VSemana & "' and status='" & _
         Mstatus & "'"
         cn.Execute Sql
      Else
         Sql = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
         Sql = Sql & "Update platmphisliquid set " & mcad & "" & " = " & "" & mcad & "" & " + " & "" & mImporte & ""
         cn.Execute Sql
      End If
   End If
   Rs.MoveNext
Loop
Rs.Close: Set Rs = Nothing
End Sub
Private Function Verifica_Ingreso(fAno As Integer, fMes As Integer, fCodigo As String, fIngreso As String) As Currency
Verifica_Ingreso = 0
Dim Sql As String
Sql = "uSp_Pla_Verifica_Ingreso '" & wcia & "'," & fAno & "," & fMes & ",'" & fCodigo & "','" & fIngreso & "'"
Dim Rq As ADODB.Recordset
If fAbrRst(Rq, Sql) Then
   If Trim(Rq(0) & "") = "" Then Verifica_Ingreso = 0 Else Verifica_Ingreso = Rq(0)
End If
Rq.Close: Set Rq = Nothing
End Function
Private Function VerificaGrati(mcod, mmes As Integer) As Boolean
VerificaGrati = True
If (mmes <> 12 And mmes <> 7) Then Exit Function
If mmes = 7 Then
   If DateDiff("D", LblFingreso.Caption, "01/07/2012") < 30 Then Exit Function
End If
If mmes = 12 Then
   If DateDiff("D", LblFingreso.Caption, "01/12/2012") < 30 Then Exit Function
End If
  
If VTipobol <> "01" And VTipobol <> "02" Then Exit Function
If Mid(LblFingreso.Caption, 7, 4) = Mid(VFProceso, 7, 4) And Mid(LblFingreso.Caption, 4, 2) = Mid(VFProceso, 4, 2) Then Exit Function

If vTipoTra = "01" Then
   If mmes = 12 Then
      If Busca_Grati <> 0 Then
         'MsgBox "Debe Generar Boleta de Gratificación Previamente", vbInformation
       '  VerificaGrati = False
         VerificaGrati = True
         Exit Function
      End If
   End If
   If mmes = 7 Then
      If Busca_Grati <> 1 Then
        MsgBox "Debe Generar Boleta de Gratificación Previamente", vbInformation
         VerificaGrati = False
         Exit Function
      End If
   End If
Else
   If mmes = 12 Then
      If VTipobol = "02" Then
        'Obtenemos la Semana de Vacaciones
        Dim vSv As String
        Dim SqlFec As String
        Dim RqVSem As ADODB.Recordset
        vSv = ""
        SqlFec = " select semana from plasemanas where cia='" & wcia & "' and ano=" & Vano & " and status<>'*' and '" & Format(VFProceso, FormatFecha) & "' between fechai and fechaf"
        If (fAbrRst(RqVSem, SqlFec)) Then
           vSv = Trim(RqVSem!semana & "")
        End If
        RqVSem.Close: Set RqVSem = Nothing
        'Fin Sem Vaca
      
         Dim Rq As ADODB.Recordset
         Sql$ = "select max(convert(int,semana))-2 from plasemanas where cia='" & wcia & "' and status<>'*' and ano=" & Mid(VFProceso, 7, 4) & ""
         If fAbrRst(Rq, Sql$) Then
            If Val(vSv) >= Val(Rq(0)) Then
               If Busca_Grati <> 0 Then
                  '@01 JCJS COMENT 1L
                  'MsgBox "Debe Generar Boleta de Gratificación Previamente", vbInformation
                  '@01 JCJS ADD INI
                  Dim mensaje2 As Integer
                  mensaje2 = MsgBox("Debe Generar Boleta de Gratificación Previamente, desea continuar?", vbInformation + vbYesNo + vbDefaultButton2, "Mensaje de Alerta")
                  If mensaje2 = 7 Then
                    VerificaGrati = False
                    Exit Function
                  End If
                  '@01 JCJS ADD FIN
               End If
            End If
         End If
         Rq.Close
      Else
         Sql$ = "select max(convert(int,semana))-2 from plasemanas where cia='" & wcia & "' and status<>'*' and ano=" & Mid(VFProceso, 7, 4) & ""
         If fAbrRst(Rq, Sql$) Then
            If Val(VSemana) >= Val(Rq(0)) Then
               If Busca_Grati <> 0 Then
                  '@01 JCJS COMENT 1L
                  'MsgBox "Debe Generar Boleta de Gratificación Previamente", vbInformation
                  '@01 JCJS ADD INI
                  Dim mensaje As Integer
                  mensaje = MsgBox("Debe Generar Boleta de Gratificación Previamente, desea continuar?", vbInformation + vbYesNo + vbDefaultButton2, "Mensaje de Alerta")
                  If mensaje = 7 Then
                    VerificaGrati = False
                    Exit Function
                  End If
                  '@01 JCJS ADD FIN
               End If
            End If
         End If
         Rq.Close
      End If
   End If
End If
End Function
Private Function Calcula_Promedios_Qta() As Double
Dim mDateBegin As String
Dim mDateEnd As String
Dim mCalculaExtras As Boolean
Dim meses As Integer
Dim mcad As String
Dim RX As ADODB.Recordset
Dim mImporte As Double
mImporte = 0

Sql = "select distinct(codinterno) from placonstante where cia='" & wcia & "' and tipomovimiento='02' and promqta='S' and status<>'*'"
If (fAbrRst(Rs, Sql)) Then Rs.MoveFirst
meses = 6

mDateBegin = Fecha_Promedios(meses, VFProceso)

If Val(Mid(VFProceso, 4, 2)) = 1 Then
   mdia = Ultimo_Dia(12, Val(Mid(VFProceso, 7, 4)) - 1)
   mDateEnd = Format(mdia, "00") & "/12/" & Format(Val(Mid(VFProceso, 7, 4)) - 1, "0000")
Else
   mdia = Ultimo_Dia(Val(Mid(VFProceso, 4, 2) - 1), Val(Mid(VFProceso, 7, 4)))
   mDateEnd = Format(mdia, "00") & "/" & Format(Val(Mid(VFProceso, 4, 2) - 1), "00") & "/" & Mid(VFProceso, 7, 4)
End If
mDateBegin = "01/" & Format(Month(DateAdd("m", (meses - 1) * -1, mDateEnd)), "00") & "/" & Format(Year(DateAdd("m", -5, mDateEnd)), "0000")

Dim FECINICIO As String
Dim FECFINAL As String
                
FECINICIO = Mid(mDateBegin, 4, 2) & "/" & Mid(mDateBegin, 1, 2) & "/" & Mid(mDateBegin, 7, 4)
FECFINAL = Mid(mDateEnd, 4, 2) & "/" & Mid(mDateEnd, 1, 2) & "/" & Mid(mDateEnd, 7, 4)

mCalculaExtras = False
Sql = "SET DATEFORMAT mdy select dbo.fc_validaHEextras('" & wcia & "','" & Txtcodpla.Text & "','" & FECINICIO & "','" & FECFINAL & "')"
If (fAbrRst(RX, Sql)) Then
   If Not IsNull(RX(0)) Then If RX(0) >= 3 Then mCalculaExtras = True
End If
RX.Close: Set RX = Nothing


mcad = ""
Do While Not Rs.EOF
   mcad = ""
   Select Case Rs(0)
   Case "10", "11", "21", "24", "25"
        If mCalculaExtras = True Then mcad = mcad & "I" & Rs(0)
   Case Else
        Sql = "SET DATEFORMAT MDY "
        Sql = Sql & "select count(1) from "
        Sql = Sql & "(SELECT COUNT(1) as c FROM plahistorico "
        Sql = Sql & "WHERE cia='" & wcia & "' and placod='" & Txtcodpla.Text & "' and status!='*' "
        Sql = Sql & "and (fechaproceso>'" & FECINICIO + " 00:00:00" & "' and fechaproceso<'" & FECFINAL + " 23:59:59" & "') "
        Sql = Sql & "and " & "I" & Rs(0) & ">0 AND proceso='01' "
        Sql = Sql & "group by month(fechaproceso)) Temp"
        If (fAbrRst(RX, Sql)) Then
           If RX(0) >= 3 Then mcad = mcad & "I" & Rs(0)
        End If
        RX.Close: Set RX = Nothing
   End Select
   
   If Trim(mcad) <> "" Then
      Sql = "SET DATEFORMAT MDY select sum(" & mcad & ") from plahistorico "
      Sql = Sql & " where cia='" & wcia & "' and fechaproceso Between '" & FECINICIO + " 00:00:00" & "' and '" & FECFINAL + "  23:59:59" & "'"
      Sql = Sql & " and proceso='01' and placod='" & Txtcodpla.Text & "' and status<>'*'"
      If (fAbrRst(RX, Sql)) Then RX.MoveFirst
      If Not IsNull(RX(0)) Then mImporte = mImporte + RX(0)
      RX.Close: Set RX = Nothing
   End If
   Rs.MoveNext
Loop
If mImporte > 0 Then mImporte = Round(mImporte / meses, 2)
Rs.Close: Set Rs = Nothing
Calcula_Promedios_Qta = mImporte
End Function
Private Function Calcula_Factor_Sobretasa_Noche() As Double

Dim rsTemp As New ADODB.Recordset
Dim ftemp As Currency
Dim codauxinterno As String
Calcula_Factor_Sobretasa_Noche = 0
ftemp = 0
codauxinterno = ""

Sql$ = "select h20 from platemphist "
Sql$ = Sql$ & "where cia='" & wcia & "' and status='" & Mstatus & "' and placod='" & Trim(Txtcodpla.Text) & "' "
Sql$ = Sql$ & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
If (fAbrRst(rsTemp, Sql$)) Then
   If rsTemp!h20 = 0 Then Exit Function
Else
   Exit Function
End If
rsTemp.Close

Sql$ = "select codauxinterno,sum(case when concepto='02' then case tipo when '02' then ((importe*4)/240) when '03' then ((importe*2)/240) else importe/factor_horas end  else importe/factor_horas end) as factor"
Sql$ = Sql$ & " From plaremunbase"
Sql$ = Sql$ & " where cia='" & wcia & "' and placod='" & Trim(Txtcodpla.Text) & "' AND Concepto in('01','02','04','05','06','08') and status !='*' group by codauxinterno"

If (fAbrRst(rsTemp, Sql$)) Then
    codauxinterno = rsTemp!codauxinterno
    ftemp = Round(rsTemp!factor, 2)
End If

If ftemp < Round((sueldominimo * 1.35 / 240), 2) Then
       ftemp = Round((sueldominimo * 1.35 / 240), 2) - ftemp
Else
       ftemp = 0
End If
Calcula_Factor_Sobretasa_Noche = Round(ftemp, 2)

End Function
Public Function Verifica_Subsidio() As Boolean
Dim sSQL As String
Dim RX As New ADODB.Recordset

Verifica_Subsidio = False

sSQL = "select h01 as Normal,h04 as Pemisos,h05 as EnfPag,h06 as SubsEnf,h27 as Subs_Mat,h29 as LicSind,h25 as LicGoce,h30 as SubsAcc  from platemphist where cia='" & wcia & "' and status='" & Mstatus & "' and placod='" & Trim(Txtcodpla.Text) & "' "
sSQL = sSQL & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
If (fAbrRst(RX, sSQL)) Then
   If RX!Normal + RX!Pemisos + RX!EnfPag + RX!LicSind + RX!LicGoce = 0 And (RX!SubsEnf <> 0 Or RX!Subs_Mat <> 0 Or SubsAcc <> 0) Then Verifica_Subsidio = True
End If
RX.Close: Set RX = Nothing
End Function
Private Function Verifica_Cuenta_Corriente() As Boolean
'mgirao cambio para solo quinta
'Verifica_Cuenta_Corriente = False
Verifica_Cuenta_Corriente = True
If rsdesadic.RecordCount > 0 Then rsdesadic.MoveFirst
Do While Not rsdesadic.EOF
    If rsdesadic("CODIGO") = "07" Then
        If Not IsNumeric(rsdesadic("MONTO")) Then rsdesadic("MONTO") = 0
        If rsdesadic("MONTO") > CCur(Lblctacte.Caption) Then
           Verifica_Cuenta_Corriente = False
        End If
        Exit Do
    End If
    rsdesadic.MoveNext
Loop
End Function
Private Function Cierre_Planilla(ano As Integer, Mes As Integer) As Boolean
Dim Rq As ADODB.Recordset
Cierre_Planilla = False
Sql = "Select User_crea From Pla_Cierre where cia='" & wcia & "' and ano=" & ano & " and mes=" & Mes & " and status<>'*'"
If fAbrRst(Rq, Sql) Then
   MsgBox "Planilla ya se Cerro no se pueden realizar modificaciones", vbInformation
   Cierre_Planilla = True
End If
If Rq.State = 1 Then Rq.Close: Set Rq = Nothing
End Function

Private Function Verifica_Activo() As Boolean
Dim Rq As ADODB.Recordset
Verifica_Activo = True
If Trim(TxtAct1.Text & "") <> "" Then
    Sql$ = "select modelo from activo where cod_cia='" & wcia & "' and placa='" & Trim(TxtAct1.Text & "") & "' and status<>'*'"
    If Not fAbrRst(Rq, Sql) Then
       MsgBox "Primer Activo no registrado", vbInformation
       Verifica_Activo = False
       If Rq.State = 1 Then Rq.Close: Set Rq = Nothing
    End If
End If
If Trim(TxtAct2.Text & "") <> "" Then
    Sql$ = "select modelo from activo where cod_cia='" & wcia & "' and placa='" & Trim(TxtAct2.Text & "") & "' and status<>'*'"
    If Not fAbrRst(Rq, Sql) Then
       MsgBox "Segundo Activo no registrado", vbInformation
       Verifica_Activo = False
       If Rq.State = 1 Then Rq.Close: Set Rq = Nothing
    End If
End If
If Trim(TxtAct3.Text & "") <> "" Then
    Sql$ = "select modelo from activo where cod_cia='" & wcia & "' and placa='" & Trim(TxtAct3.Text & "") & "' and status<>'*'"
    If Not fAbrRst(Rq, Sql) Then
       MsgBox "Tercer Activo no registrado", vbInformation
       Verifica_Activo = False
       If Rq.State = 1 Then Rq.Close: Set Rq = Nothing
    End If
End If
If Trim(TxtAct4.Text & "") <> "" Then
    Sql$ = "select modelo from activo where cod_cia='" & wcia & "' and placa='" & Trim(TxtAct4.Text & "") & "' and status<>'*'"
    If Not fAbrRst(Rq, Sql) Then
       MsgBox "Cuarto Activo no registrado", vbInformation
       Verifica_Activo = False
       If Rq.State = 1 Then Rq.Close: Set Rq = Nothing
    End If
End If
If Verifica_Activo = True Then
    If Trim(TxtAct3.Text & "") <> "" Or Trim(TxtAct2.Text & "") <> "" Or Trim(TxtAct1.Text & "") <> "" Or Trim(TxtAct4.Text & "") <> "" Then
        If Val(TxtHor1.Text) + Val(TxtHor2.Text) + Val(TxtHor3.Text) + Val(TxtHor4.Text) <> t_horas_trabajadas Then
              ' MsgBox "Suma de Horas Aignadas a los Activos no suma el 100 % de las horas Trabajadas", vbInformation
              ' Verifica_Activo = False
               If Rq.State = 1 Then Rq.Close: Set Rq = Nothing
        End If
    End If
End If
End Function
Public Sub Importar_Utilidades()
Dim VanoUti As Integer
VanoUti = Val(Mid(VFProceso, 7, 4)) - 1

Sql$ = "Usp_Pla_Importa_Utilidades '" & wcia & "'," & VanoUti & ",'" & vTipoTra & "'"
cn.CursorLocation = adUseClient
Set rsExport = New ADODB.Recordset
Set rsExport = cn.Execute(Sql$, 64)
         
If rsExport.RecordCount <= 0 Then
    MsgBox "No se ha efectuado el calculo de utilidades", vbCritical, Me.Caption
    GoTo Salir:
End If
    
rsExport.MoveLast
rsExport.MoveFirst
Set DGrd.DataSource = rsExport
DGrd.Refresh
    
'fc_SumaTotalesImportacion rsExport, DGrd
    
Salir:

End Sub

Private Function Verifica_Meses_Grati() As Boolean
Verifica_Meses_Grati = True
Dim mmes As Currency
Dim mdias As Currency
If rshoras.RecordCount > 0 Then
   rshoras.MoveFirst
   Do While Not rshoras.EOF
      Select Case rshoras!Codigo
             Case "21"
                  If IsNumeric(rshoras!Monto) Then
                     If rshoras!Monto > 6 Then
                        If bImportar = False Then
                           MsgBox "No puede excederse de los 6 meses", vbInformation
                        Else
                           LstObs.AddItem Trim(Txtcodpla.Text) & ": Meses no pueden exceder a 6"
                        End If
                        mmes = rshoras!Monto
                        Verifica_Meses_Grati = False
                     End If
                  Else
                     If bImportar = False Then
                        MsgBox "Ingrese Correctamente Nùmero de Meses", vbInformation
                     Else
                        LstObs.AddItem Trim(Txtcodpla.Text) & ": Ingrese Correctamente Nùmero de Meses"
                     End If
                     Verifica_Meses_Grati = False
                     mmes = 0
                  End If
             Case "14"
                  If IsNumeric(rshoras!Monto) Then
                     If rshoras!Monto > 30 Then
                        If bImportar = False Then
                           MsgBox "Dìas no pueden exceder a 30", vbInformation
                        Else
                           LstObs.AddItem Trim(Txtcodpla.Text) & ": Dìas no pueden exceder a 30"
                        End If
                        Verifica_Meses_Grati = False
                        mdias = rshoras!Monto
                     End If
                  Else
                     If bImportar = False Then
                        MsgBox "Ingrese Correctamente Nùmero de Dìas", vbInformation
                     Else
                        LstObs.AddItem Trim(Txtcodpla.Text) & ": Ingrese Correctamente Nùmero de Dìas"
                     End If
                     Verifica_Meses_Grati = False
                     mdias = 0
                  End If
      End Select
      rshoras.MoveNext
   Loop
   If mmes >= 6 And mdias <> 0 Then
      If bImportar = False Then
         MsgBox "No puede excederse de los 6 meses", vbInformation
      Else
         LstObs.AddItem Trim(Txtcodpla.Text) & ": No puede excederse de los 6 meses"
      End If
      Verifica_Meses_Grati = False
   End If
   
   If mmes <> CInt(mmes) Then
      If bImportar = False Then
         MsgBox "Nùmero de Meses debe ser entero", vbInformation
      Else
         LstObs.AddItem Trim(Txtcodpla.Text) & ": Nùmero de Meses debe ser entero"
      End If
      Verifica_Meses_Grati = False
   End If
   
   If mdias <> CInt(mdias) Then
      If bImportar = False Then
         MsgBox "Nùmero de Dìas debe ser entero", vbInformation
      Else
         LstObs.AddItem Trim(Txtcodpla.Text) & ": Nùmero de Dìas debe ser entero"
      End If
      Verifica_Meses_Grati = False
   End If
End If
End Function
