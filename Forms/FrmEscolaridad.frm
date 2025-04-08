VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmEscolaridad 
   Caption         =   "Escolaridad"
   ClientHeight    =   8820
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15420
   Icon            =   "FrmEscolaridad.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8820
   ScaleWidth      =   15420
   Begin VB.Frame FrameTxtBco 
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   840
      TabIndex        =   17
      Top             =   3720
      Visible         =   0   'False
      Width           =   7935
      Begin VB.ComboBox CmbBco 
         Height          =   315
         Left            =   3600
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   120
         Width           =   3015
      End
      Begin MSComCtl2.DTPicker Cbofecha 
         Height          =   285
         Left            =   1560
         TabIndex        =   19
         Top             =   120
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Format          =   62980097
         CurrentDate     =   37265
      End
      Begin MSForms.CommandButton CommandButton1 
         Height          =   375
         Left            =   7320
         TabIndex        =   24
         Top             =   120
         Width           =   480
         PicturePosition =   327683
         Size            =   "847;661"
         Picture         =   "FrmEscolaridad.frx":030A
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Deposito"
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
         Left            =   120
         TabIndex        =   22
         Top             =   180
         Width           =   1350
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Banco"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3000
         TabIndex        =   21
         Top             =   180
         Width           =   510
      End
      Begin MSForms.CommandButton CommandButton2 
         Height          =   375
         Left            =   6720
         TabIndex        =   20
         Top             =   120
         Width           =   480
         PicturePosition =   327683
         Size            =   "847;661"
         Picture         =   "FrmEscolaridad.frx":0624
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
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
      Height          =   7050
      Left            =   9720
      TabIndex        =   11
      Top             =   1560
      Width           =   5610
      Begin TrueOleDBGrid70.TDBGrid DGrd 
         Height          =   5835
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   10292
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
      Begin Threed.SSCommand SSCommand1 
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   6480
         Width           =   5415
         _Version        =   65536
         _ExtentX        =   9551
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "IMPORTAR"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "FrmEscolaridad.frx":0BBE
      End
      Begin MSComctlLib.ProgressBar BarraImporta 
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   6120
         Width           =   5445
         _ExtentX        =   9604
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   0
      TabIndex        =   3
      Top             =   600
      Width           =   15375
      Begin VB.TextBox Txtarchivos 
         BackColor       =   &H8000000A&
         Height          =   285
         Left            =   8955
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   360
         Width           =   5370
      End
      Begin VB.TextBox Txtano 
         Height          =   285
         Left            =   855
         TabIndex        =   5
         Top             =   360
         Width           =   825
      End
      Begin VB.ComboBox Cmbmes 
         Height          =   315
         ItemData        =   "FrmEscolaridad.frx":0BDA
         Left            =   1740
         List            =   "FrmEscolaridad.frx":0C02
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   360
         Width           =   2475
      End
      Begin Threed.SSCommand cmdVerArchivo 
         Height          =   615
         Left            =   14520
         TabIndex        =   8
         Top             =   120
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Picture         =   "FrmEscolaridad.frx":0C6A
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   375
         Left            =   4440
         TabIndex        =   23
         Top             =   360
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "Archivo Banco"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "FrmEscolaridad.frx":0F84
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "La Hoja debe tener como nombre ""escolaridad"""
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
         Index           =   2
         Left            =   8880
         TabIndex        =   16
         Top             =   120
         Width           =   4050
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
         Left            =   7920
         TabIndex        =   10
         Top             =   375
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Periodo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   6
         Top             =   405
         Width           =   660
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Escolaridad"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   17895
      Begin VB.ComboBox Cmbcia 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   90
         Width           =   14250
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
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   825
      End
   End
   Begin TrueOleDBGrid70.TDBGrid Dgrdescolaridad 
      Height          =   6705
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   9435
      _ExtentX        =   16642
      _ExtentY        =   11827
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Codigo"
      Columns(0).DataField=   "placod"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Nombre"
      Columns(1).DataField=   "Nombre"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Hijos"
      Columns(2).DataField=   "hijos"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Importe_Hijo"
      Columns(3).DataField=   "importe_hijo"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Importe_Total"
      Columns(4).DataField=   "Importe_Total"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   5
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).FetchRowStyle=   -1  'True
      Splits(0).DividerColor=   13160660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=5"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1349"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1270"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=9181"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=9102"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(9)=   "Column(2).Width=979"
      Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=900"
      Splits(0)._ColumnProps(12)=   "Column(2)._ColStyle=2"
      Splits(0)._ColumnProps(13)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(14)=   "Column(3).Width=1931"
      Splits(0)._ColumnProps(15)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(16)=   "Column(3)._WidthInPix=1852"
      Splits(0)._ColumnProps(17)=   "Column(3)._ColStyle=2"
      Splits(0)._ColumnProps(18)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(19)=   "Column(4).Width=2461"
      Splits(0)._ColumnProps(20)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(21)=   "Column(4)._WidthInPix=2381"
      Splits(0)._ColumnProps(22)=   "Column(4)._ColStyle=2"
      Splits(0)._ColumnProps(23)=   "Column(4).Order=5"
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
      _StyleDefs(43)  =   "Splits(0).Columns(1).Style:id=66,.parent=13"
      _StyleDefs(44)  =   "Splits(0).Columns(1).HeadingStyle:id=63,.parent=14"
      _StyleDefs(45)  =   "Splits(0).Columns(1).FooterStyle:id=64,.parent=15"
      _StyleDefs(46)  =   "Splits(0).Columns(1).EditorStyle:id=65,.parent=17"
      _StyleDefs(47)  =   "Splits(0).Columns(2).Style:id=54,.parent=13,.alignment=1"
      _StyleDefs(48)  =   "Splits(0).Columns(2).HeadingStyle:id=51,.parent=14"
      _StyleDefs(49)  =   "Splits(0).Columns(2).FooterStyle:id=52,.parent=15"
      _StyleDefs(50)  =   "Splits(0).Columns(2).EditorStyle:id=53,.parent=17"
      _StyleDefs(51)  =   "Splits(0).Columns(3).Style:id=58,.parent=13,.alignment=1"
      _StyleDefs(52)  =   "Splits(0).Columns(3).HeadingStyle:id=55,.parent=14"
      _StyleDefs(53)  =   "Splits(0).Columns(3).FooterStyle:id=56,.parent=15"
      _StyleDefs(54)  =   "Splits(0).Columns(3).EditorStyle:id=57,.parent=17"
      _StyleDefs(55)  =   "Splits(0).Columns(4).Style:id=62,.parent=13,.alignment=1"
      _StyleDefs(56)  =   "Splits(0).Columns(4).HeadingStyle:id=59,.parent=14"
      _StyleDefs(57)  =   "Splits(0).Columns(4).FooterStyle:id=60,.parent=15"
      _StyleDefs(58)  =   "Splits(0).Columns(4).EditorStyle:id=61,.parent=17"
      _StyleDefs(59)  =   "Named:id=33:Normal"
      _StyleDefs(60)  =   ":id=33,.parent=0"
      _StyleDefs(61)  =   "Named:id=34:Heading"
      _StyleDefs(62)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(63)  =   ":id=34,.wraptext=-1"
      _StyleDefs(64)  =   "Named:id=35:Footing"
      _StyleDefs(65)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(66)  =   "Named:id=36:Selected"
      _StyleDefs(67)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(68)  =   "Named:id=37:Caption"
      _StyleDefs(69)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(70)  =   "Named:id=38:HighlightRow"
      _StyleDefs(71)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(72)  =   "Named:id=39:EvenRow"
      _StyleDefs(73)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(74)  =   "Named:id=40:OddRow"
      _StyleDefs(75)  =   ":id=40,.parent=33"
      _StyleDefs(76)  =   "Named:id=41:RecordSelector"
      _StyleDefs(77)  =   ":id=41,.parent=34"
      _StyleDefs(78)  =   "Named:id=42:FilterBar"
      _StyleDefs(79)  =   ":id=42,.parent=33"
   End
   Begin MSComDlg.CommonDialog Box 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label LblTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
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
      Left            =   8040
      TabIndex        =   14
      Top             =   8400
      Width           =   1455
   End
End
Attribute VB_Name = "FrmEscolaridad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsExport As New ADODB.Recordset
Dim rsdepo As New Recordset

Private Sub CmbBco_Click()
If rsdepo.RecordCount > 0 Then
   rsdepo.MoveFirst
   Do While Not rsdepo.EOF
      rsdepo.Delete
      rsdepo.MoveNext
   Loop
End If

Dim VBcoPago As String
VBcoPago = fc_CodigoComboBox(CmbBco, 2)
Dim rsbco As ADODB.Recordset
Dim xSem As String

Sql = "Usp_Pla_Archivo_Importa_Escolaridad '" & wcia & "','" & VBcoPago & "'," & Format(Txtano.Text, "0000") & "," & Format(Cmbmes.ListIndex + 1, "00") & ""
If (fAbrRst(rsbco, Sql)) Then rsbco.MoveFirst
Do While Not rsbco.EOF
    rsdepo.AddNew
    rsdepo!codigo = Trim(rsbco!codigo & "")
    rsdepo!NOM_CLIE = Trim(rsbco!NOM_CLIE & "")
    rsdepo!NUMEROCTA = Trim(rsbco!NUMEROCTA & "")
    rsdepo!importe = rsbco!importe
    rsdepo!TIPO_REG = Trim(rsbco!TIPO_REG & "")
    rsdepo!Cuenta = Trim(rsbco!Cuenta & "")
    rsdepo!tipo_doc = Trim(rsbco!tipo_doc & "")
    rsdepo!NUMERO_DOC = Trim(rsbco!NUMERO_DOC & "")
    rsdepo!REF_TRAB = Trim(rsbco!REF_TRAB & "")
    rsdepo!REF_EMP = Trim(rsbco!REF_EMP & "")
    rsdepo!moneda = Trim(rsbco!moneda & "")
    rsdepo!FLAG = Trim(rsbco!FLAG & "")
    rsdepo!cta = Trim(rsbco!cta & "")
    rsdepo!CTAINTER = Trim(rsbco!CTAINTER & "")
    rsdepo!Sucursal = Trim(rsbco!Sucursal & "")
    rsdepo!PAGOCUENTA = Trim(rsbco!PAGOCUENTA & "")
    rsdepo!Excluir = ""
    rsbco.MoveNext
Loop
If rsdepo.RecordCount > 0 Then rsdepo.MoveFirst

End Sub

Private Sub Cmbmes_Click()
Carga_Consulta
End Sub

Private Sub cmdVerArchivo_Click()
    If Not IsNumeric(Txtano.Text) Then
       MsgBox "Ingrese Año Correctamente"
    End If
    
    
    Set rsExport = Nothing
    LimpiarRsT rsExport, DGrd
    AbrirFile ("*.xls")
    If Trim(Txtarchivos.Text) <> "" Then Importar_Excel
End Sub

Public Sub AbrirFile(pextension As String)
If Not Cuadro_Dialogo_Abrir(pextension) Then
    Txtarchivos.Text = ""
    Exit Sub
End If
Debug.Print Box.DefaultExt

If UCase(Right(Box.FileName, 3)) <> UCase(Right(pextension, 3)) Then
   MsgBox "La Extensión de archivo no concuerda con el formato elegido", vbCritical, "Archivo Inválido"
   Exit Sub
End If
Txtarchivos.Text = Box.FileName
Txtarchivos.ToolTipText = Box.FileName

End Sub

Private Sub CommandButton1_Click()
FrameTxtBco.Visible = False
End Sub

Private Sub CommandButton2_Click()
Dim VBcoPago As String
VBcoPago = fc_CodigoComboBox(CmbBco, 2)

Select Case VBcoPago
   Case "01": Print_TxtBcoCredito
   Case "29": Print_TxtBcoScotia
   Case Else
      MsgBox "Implementación no desarrollada para el banco seleccionado", vbInformation
End Select

End Sub

Private Sub Form_Load()
Me.Top = 0: Me.Left = 0
Me.Width = 15540: Me.Height = 9330
Txtano.Text = Format(Year(Date), "0000")
If Month(Date) = 1 Then
   Cmbmes.ListIndex = 0
Else
   Cmbmes.ListIndex = Month(Date) - 1
End If
Carga_Consulta
Cbofecha.Value = Date
Call fc_Descrip_Maestros2("01007", "", CmbBco, False)
Crea_Rs
End Sub
Public Sub Importar_Excel()

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
   
On Error GoTo ERR
        
Set xlApp2 = xlApp1.Application
Dim Col As Integer, Fila As Integer

Set xLibro = xlApp2.Workbooks.Open(Txtarchivos.Text)
  
xlApp2.Visible = False
    
Dim xTipoTrab As String
    
If Not xlApp2 Is Nothing Then Set xlApp2 = Nothing
If Not xlApp1 Is Nothing Then Set xlApp1 = Nothing
If Not xLibro Is Nothing Then Set xlBook = Nothing
    
Dim conexion As ADODB.Connection, rs As ADODB.Recordset
  
Set conexion = New ADODB.Connection
       
conexion.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                  "Data Source=" & Txtarchivos.Text & _
                  ";Extended Properties=""Excel 8.0;HDR=Yes;"""

           
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
         
rsExport.Open "SELECT * FROM [escolaridad$] Where Codigo<>'' Order by Codigo", conexion, , , adCmdText
         
Set DGrd.DataSource = rsExport
    
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
Private Sub Importar_Escolaridad()
   'Confirmar proceso de importación
   Dim NroTrans As Integer
   'On Error GoTo Salir
   NroTrans = 0
     
    
   If MsgBox("Seguro de Cargar Escolaridad", vbYesNo + vbQuestion, "Escolaridad") = vbNo Then Exit Sub
   
   Sql$ = "Update pla_escolaridad set status='*',user_modi='" & wuser & "',fec_modi=getdate() where cia='" & wcia & "' and ayo = " & Txtano.Text & " and mes = " & Cmbmes.ListIndex + 1 & " and status<>'*'"
   cn.Execute Sql$
   Dim lPasa As Integer
   Dim Rq As ADODB.Recordset
   lPasa = 0
   BarraImporta.Max = rsExport.RecordCount
   Screen.MousePointer = 11
   With rsExport
    If .RecordCount > 0 Then
        '//*** Verifica si la estuctura es la correcta ***///
        For I = 0 To DGrd.Columns.count - 1
            If UCase(DGrd.Columns(I).Caption) = UCase("IMPORTEHIJO") Then lPasa = lPasa + 1
            If UCase(DGrd.Columns(I).Caption) = UCase("CODIGO") Then lPasa = lPasa + 1
            If UCase(DGrd.Columns(I).Caption) = UCase("HIJOS") Then lPasa = lPasa + 1
        Next
        If lPasa <> 3 Then
           MsgBox "Los nombres de los campos debebn ser" & Chr(13) & "CODIGO" & Chr(13) & "HIJOS" & Chr(13) & "IMPORTEHIJO", vbInformation
           MsgBox "No se realizó la carga", vbInformation
           Screen.MousePointer = 0
           Exit Sub
        End If
        I = 1
    
        .MoveFirst
        lItems = 0
        Do While Not .EOF
           If Trim(.Fields("CODIGO").Value + "") <> "" Then
              Sql$ = "Select fcese from planillas where cia='" & wcia & "' and placod='" & Trim(.Fields("CODIGO").Value + "") & "' and status<>'*'"
              If fAbrRst(Rq, Sql$) Then
                 If Not IsNull(Rq!fcese) Then
                    MsgBox "Trabajador es cesado = > " & Trim(.Fields("CODIGO").Value + "") & Chr(13) & "No se cargara el trabajador", vbInformation
                 End If
              Else
                 MsgBox "CODIGO NO REGISTRADO => " & Trim(.Fields("CODIGO").Value + ""), vbInformation
              End If
              Rq.Close: Set Rq = Nothing
           End If
           Sql$ = "insert into pla_escolaridad VALUES('" & wcia & "'," & Txtano.Text & "," & Cmbmes.ListIndex + 1 & ",'" & Trim(.Fields("CODIGO").Value + "") & "'," & CCur(.Fields("hijos").Value) & "," & CCur(.Fields("IMPORTEHIJO").Value) & "," & Round(CCur(.Fields("hijos").Value) * CCur(.Fields("IMPORTEHIJO").Value), 2) & ",'','" & wuser & "',getdate(),'',getdate())"
           cn.Execute Sql$
           lItems = lItems + 1
           BarraImporta.Value = I
           I = I + 1
           .MoveNext
        Loop
    End If
    BarraImporta.Visible = False

    Screen.MousePointer = 0
    MsgBox "Terminó la carga Correctamente " & lItems & " Registros importados", vbInformation, Me.Caption
End With
Screen.MousePointer = 0

Exit Sub
Salir:
    If NroTrans = 1 Then
    End If
    Screen.MousePointer = 0
    MsgBox ERR.Description, vbCritical, Me.Caption
    

End Sub

Private Sub SSCommand1_Click()
Importar_Escolaridad
Carga_Consulta
End Sub
Private Sub Carga_Consulta()
If Not IsNumeric(Txtano.Text) Then
   Set Dgrdescolaridad.DataSource = Nothing
   LblTotal.Caption = ""
   Exit Sub
End If
If Val(Txtano.Text) < 2015 Then
   Set Dgrdescolaridad.DataSource = Nothing
   LblTotal.Caption = ""
   Exit Sub
End If

Dim Rq As ADODB.Recordset
Sql$ = "select isnull(sum(importe_total),0) as Total from pla_escolaridad where cia='" & wcia & "' and ayo=" & Txtano.Text & " and mes=" & Cmbmes.ListIndex + 1 & " and status<>'*'"
If fAbrRst(Rq, Sql$) Then
   LblTotal.Caption = Rq!Total
End If
Rq.Close: Set Rq = Nothing

Sql$ = "select E.*,RTRIM(AP_PAT)+ ' '+RTRIM(AP_MAT)+ ' '+RTRIM(NOM_1)+ ' '+RTRIM(NOM_2) AS NOMBRE from pla_escolaridad E,PLANILLAS P where E.cia='" & wcia & "' and E.ayo=" & Txtano.Text & " and E.mes=" & Cmbmes.ListIndex + 1 & " and E.status<>'*' AND P.cia=E.cia AND P.placod=E.placod AND P.status<>'*' order by e.placod"

Dim rstrab As New ADODB.Recordset
cn.CursorLocation = adUseClient
Set rstrab = New ADODB.Recordset
Set rstrab = cn.Execute(Sql$, 64)
Set Dgrdescolaridad.DataSource = rstrab

End Sub

Private Sub SSCommand2_Click()
If rsdepo.RecordCount > 0 Then
   rsdepo.MoveFirst
   Do While Not rsdepo.EOF
      rsdepo.Delete
      rsdepo.MoveNext
   Loop
End If


FrameTxtBco.Visible = True

End Sub

Private Sub Txtano_Change()
Carga_Consulta
End Sub

Private Sub Print_TxtBcoCredito()
Dim mfproceso As String
mfproceso = Str(Cbofecha.Year) + Format(Cbofecha.Month, "00") + Format(Cbofecha.Day, "00")

Dim nTrab As Integer
Dim mneto As Double
Dim mscta As Double

nTrab = 0: mneto = 0
If rsdepo.RecordCount > 0 Then rsdepo.MoveFirst
Do While Not rsdepo.EOF
   If rsdepo!importe < 0 Then
      MsgBox "Existen Netos en negativos, no se generara el archivo", vbInformation
      rsdepo.Close: Set rsdepo = Nothing
      Exit Sub
   End If
   If Trim((rsdepo!cta & "") & "") = "" Then
      MsgBox "Existen trabajadores sin cuenta, no se generara el archivo", vbInformation
      rsdepo.Close: Set rsdepo = Nothing
      Exit Sub
   End If

   If UCase(Trim(rsdepo!Excluir & "")) <> "S" Then
      nTrab = nTrab + 1
      mscta = mscta + rsdepo!cta
      mneto = mneto + rsdepo!importe
   End If
   rsdepo.MoveNext
Loop

RUTA$ = App.Path & "\REPORTS\EscoCred.txt"
Open RUTA$ For Output As #1

Dim mcad As String
Dim mt As String
mcad = "1" & Llenar_Ceros(Trim(Str(nTrab)), 6) & Trim(mfproceso)
mt = "O"

mcad = mcad & mt & "C00011910215470064       " & Llenar_Ceros(Trim(fCadNum(mneto, "#######0.00")), 17)
mcad = mcad & "Planilla de Haberes " & Space(20)

Dim MCOD_CTA As Double
MCOD_CTA = 215470064
mscta = mscta + MCOD_CTA

mcad = mcad & Llenar_Ceros(Trim(Str(mscta)), 15)
Print #1, mcad

If rsdepo.RecordCount > 0 Then rsdepo.MoveFirst

Do While Not rsdepo.EOF
   If rsdepo!importe > 0 And UCase(Trim(rsdepo!Excluir & "")) <> "S" Then
      Print #1, rsdepo!TIPO_REG + rsdepo!Cuenta + rsdepo!NUMEROCTA + rsdepo!tipo_doc + rsdepo!NUMERO_DOC + Space(3) + rsdepo!NOM_CLIE + rsdepo!REF_TRAB + rsdepo!REF_EMP + rsdepo!moneda + Llenar_Ceros(Trim(fCadNum(rsdepo!importe, "#######0.00")), 17) + rsdepo!FLAG
   End If
   rsdepo.MoveNext
 Loop
Close #1
Call Imprime_Txt("EscoCred.txt", RUTA$)

FrameTxtBco.Visible = False
End Sub

Private Sub Print_TxtBcoScotia()
Dim mfproceso As String
mfproceso = Str(Cbofecha.Year) + Format(Cbofecha.Month, "00") + Format(Cbofecha.Day, "00")

If rsdepo.RecordCount > 0 Then rsdepo.MoveFirst
Do While Not rsdepo.EOF
   If rsdepo!importe < 0 Then
      MsgBox "Existen Netos en negativos, no se generara el archivo", vbInformation
      rsdepo.Close: Set rsdepo = Nothing
      Exit Sub
   End If
   If Trim((rsdepo!cta & "") & "") = "" Then
      MsgBox "Existen trabajadores sin cuenta, no se generara el archivo", vbInformation
      rsdepo.Close: Set rsdepo = Nothing
      Exit Sub
   End If
   rsdepo.MoveNext
Loop

RUTA$ = App.Path & "\REPORTS\Escoscot.txt"
Open RUTA$ For Output As #1

If rsdepo.RecordCount > 0 Then rsdepo.MoveFirst
Dim mcpago As String
mcpago = "Planilla Normal             "

Do While Not rsdepo.EOF
   If rsdepo!importe > 0 And UCase(Trim(rsdepo!Excluir & "")) <> "S" Then
      Print #1, rsdepo!codigo + Mid(rsdepo!NOM_CLIE, 1, 30) + Mid(mcpago, 1, 20) + Trim(mfproceso) + Llenar_Ceros(Int(rsdepo!importe * 100), 11) + "3" + rsdepo!Sucursal + Mid(rsdepo!PAGOCUENTA, 1, 7) + Mid(rsdepo!NUMERO_DOC, 1, 8) + " " + rsdepo!CTAINTER
   End If
   rsdepo.MoveNext
 Loop
Close #1

Call Imprime_Txt("Escoscot.txt", RUTA$)

FrameTxtBco.Visible = False
End Sub
Private Sub Crea_Rs()
    
    If rsdepo.State = 1 Then rsdepo.Close
    rsdepo.Fields.Append "TIPO_REG", adChar, 1, adFldIsNullable
    rsdepo.Fields.Append "CUENTA", adChar, 1, adFldIsNullable
    rsdepo.Fields.Append "NUMEROCTA", adChar, 20, adFldIsNullable
    rsdepo.Fields.Append "TIPO_DOC", adChar, 1, adFldIsNullable
    rsdepo.Fields.Append "NUMERO_DOC", adChar, 12, adFldIsNullable
    rsdepo.Fields.Append "NOM_CLIE", adChar, 75, adFldIsNullable
    rsdepo.Fields.Append "REF_TRAB", adChar, 40, adFldIsNullable
    rsdepo.Fields.Append "REF_EMP", adChar, 20, adFldIsNullable
    rsdepo.Fields.Append "MONEDA", adChar, 4, adFldIsNullable
    rsdepo.Fields.Append "IMPORTE", adDouble, 2, adFldIsNullable
    rsdepo.Fields.Append "FLAG", adChar, 1, adFldIsNullable
    rsdepo.Fields.Append "cta", adChar, 20, adFldIsNullable
    rsdepo.Fields.Append "CODIGO", adChar, 8, adFldIsNullable
    rsdepo.Fields.Append "CTAINTER", adChar, 20, adFldIsNullable
    rsdepo.Fields.Append "SUCURSAL", adChar, 3, adFldIsNullable
    rsdepo.Fields.Append "PAGOCUENTA", adChar, 20, adFldIsNullable
    
  
    rsdepo.Fields.Append "EXCLUIR", adChar, 1, adFldIsNullable

    rsdepo.Open

End Sub

