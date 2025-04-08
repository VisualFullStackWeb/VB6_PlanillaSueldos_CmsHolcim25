VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmCtaCte 
   Caption         =   "Saldos de Cuenta Corriente de los Trabajadores"
   ClientHeight    =   8295
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8655
   Icon            =   "FrmCtaCte.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8295
   ScaleWidth      =   8655
   Begin Threed.SSCommand SSCommand4 
      Height          =   480
      Left            =   7320
      TabIndex        =   16
      ToolTipText     =   "Salir"
      Top             =   600
      Width           =   1305
      _Version        =   65536
      _ExtentX        =   2302
      _ExtentY        =   847
      _StockProps     =   78
      Caption         =   "Reporte Mensual"
      BevelWidth      =   1
      Picture         =   "FrmCtaCte.frx":030A
   End
   Begin VB.CheckBox ChkCeros 
      Caption         =   "Incluir trabajadores sin saldo"
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
      Left            =   3600
      TabIndex        =   15
      Top             =   720
      Width           =   2895
   End
   Begin VB.ComboBox Cmbtipotrabajador 
      Height          =   315
      ItemData        =   "FrmCtaCte.frx":0326
      Left            =   1320
      List            =   "FrmCtaCte.frx":0328
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   720
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8655
      Begin VB.ComboBox Cmbcia 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   60
         Width           =   5490
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
         TabIndex        =   3
         Top             =   120
         Width           =   825
      End
      Begin VB.Label Lblfecha 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6675
         TabIndex        =   2
         Top             =   120
         Width           =   1815
      End
   End
   Begin Threed.SSCommand SSCommand2 
      Height          =   480
      Left            =   6720
      TabIndex        =   14
      ToolTipText     =   "Salir"
      Top             =   600
      Width           =   585
      _Version        =   65536
      _ExtentX        =   1032
      _ExtentY        =   847
      _StockProps     =   78
      BevelWidth      =   1
      Picture         =   "FrmCtaCte.frx":032A
   End
   Begin VB.Frame FrameKardex 
      Height          =   7815
      Left            =   0
      TabIndex        =   8
      Top             =   360
      Visible         =   0   'False
      Width           =   8775
      Begin TrueOleDBGrid70.TDBGrid DgrdMovi 
         Height          =   7185
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   8505
         _ExtentX        =   15002
         _ExtentY        =   12674
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Fecha"
         Columns(0).DataField=   "fecha"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Motivo"
         Columns(1).DataField=   "motivo"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Cargo"
         Columns(2).DataField=   "cargo"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Abono"
         Columns(3).DataField=   "abono"
         Columns(3).NumberFormat=   "General Number"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Saldo"
         Columns(4).DataField=   "saldo"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Semana"
         Columns(5).DataField=   "semana"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Referencia"
         Columns(6).DataField=   "refe"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "Id"
         Columns(7).DataField=   "Id_CtaCte"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "tipo"
         Columns(8).DataField=   "tipo"
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   9
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).FetchRowStyle=   -1  'True
         Splits(0).DividerColor=   13160660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=9"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2223"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2143"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=3201"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=3122"
         Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(9)=   "Column(2).Width=2487"
         Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=2408"
         Splits(0)._ColumnProps(12)=   "Column(2)._ColStyle=2"
         Splits(0)._ColumnProps(13)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(14)=   "Column(3).Width=2170"
         Splits(0)._ColumnProps(15)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(16)=   "Column(3)._WidthInPix=2090"
         Splits(0)._ColumnProps(17)=   "Column(3)._ColStyle=2"
         Splits(0)._ColumnProps(18)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(19)=   "Column(4).Width=2408"
         Splits(0)._ColumnProps(20)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(21)=   "Column(4)._WidthInPix=2328"
         Splits(0)._ColumnProps(22)=   "Column(4)._ColStyle=2"
         Splits(0)._ColumnProps(23)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(24)=   "Column(5).Width=1508"
         Splits(0)._ColumnProps(25)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(26)=   "Column(5)._WidthInPix=1429"
         Splits(0)._ColumnProps(27)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(28)=   "Column(6).Width=2725"
         Splits(0)._ColumnProps(29)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(30)=   "Column(6)._WidthInPix=2646"
         Splits(0)._ColumnProps(31)=   "Column(6).Visible=0"
         Splits(0)._ColumnProps(32)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(33)=   "Column(7).Width=2725"
         Splits(0)._ColumnProps(34)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(35)=   "Column(7)._WidthInPix=2646"
         Splits(0)._ColumnProps(36)=   "Column(7).Visible=0"
         Splits(0)._ColumnProps(37)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(38)=   "Column(8).Width=2725"
         Splits(0)._ColumnProps(39)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(40)=   "Column(8)._WidthInPix=2646"
         Splits(0)._ColumnProps(41)=   "Column(8).Visible=0"
         Splits(0)._ColumnProps(42)=   "Column(8).Order=9"
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
         _StyleDefs(39)  =   "Splits(0).Columns(0).Style:id=46,.parent=13"
         _StyleDefs(40)  =   "Splits(0).Columns(0).HeadingStyle:id=43,.parent=14"
         _StyleDefs(41)  =   "Splits(0).Columns(0).FooterStyle:id=44,.parent=15"
         _StyleDefs(42)  =   "Splits(0).Columns(0).EditorStyle:id=45,.parent=17"
         _StyleDefs(43)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
         _StyleDefs(44)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
         _StyleDefs(45)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
         _StyleDefs(46)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
         _StyleDefs(47)  =   "Splits(0).Columns(2).Style:id=50,.parent=13,.alignment=1,.fgcolor=&HFF&"
         _StyleDefs(48)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=14"
         _StyleDefs(49)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=15"
         _StyleDefs(50)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=17"
         _StyleDefs(51)  =   "Splits(0).Columns(3).Style:id=32,.parent=13,.alignment=1,.fgcolor=&HFF0000&"
         _StyleDefs(52)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
         _StyleDefs(53)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
         _StyleDefs(54)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17,.fgcolor=&H80000001&"
         _StyleDefs(55)  =   "Splits(0).Columns(4).Style:id=54,.parent=13,.alignment=1,.fgcolor=&HFF&"
         _StyleDefs(56)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
         _StyleDefs(57)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
         _StyleDefs(58)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
         _StyleDefs(59)  =   "Splits(0).Columns(5).Style:id=58,.parent=13"
         _StyleDefs(60)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
         _StyleDefs(61)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
         _StyleDefs(62)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
         _StyleDefs(63)  =   "Splits(0).Columns(6).Style:id=66,.parent=13"
         _StyleDefs(64)  =   "Splits(0).Columns(6).HeadingStyle:id=63,.parent=14"
         _StyleDefs(65)  =   "Splits(0).Columns(6).FooterStyle:id=64,.parent=15"
         _StyleDefs(66)  =   "Splits(0).Columns(6).EditorStyle:id=65,.parent=17"
         _StyleDefs(67)  =   "Splits(0).Columns(7).Style:id=62,.parent=13"
         _StyleDefs(68)  =   "Splits(0).Columns(7).HeadingStyle:id=59,.parent=14"
         _StyleDefs(69)  =   "Splits(0).Columns(7).FooterStyle:id=60,.parent=15"
         _StyleDefs(70)  =   "Splits(0).Columns(7).EditorStyle:id=61,.parent=17"
         _StyleDefs(71)  =   "Splits(0).Columns(8).Style:id=70,.parent=13"
         _StyleDefs(72)  =   "Splits(0).Columns(8).HeadingStyle:id=67,.parent=14"
         _StyleDefs(73)  =   "Splits(0).Columns(8).FooterStyle:id=68,.parent=15"
         _StyleDefs(74)  =   "Splits(0).Columns(8).EditorStyle:id=69,.parent=17"
         _StyleDefs(75)  =   "Named:id=33:Normal"
         _StyleDefs(76)  =   ":id=33,.parent=0"
         _StyleDefs(77)  =   "Named:id=34:Heading"
         _StyleDefs(78)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(79)  =   ":id=34,.wraptext=-1"
         _StyleDefs(80)  =   "Named:id=35:Footing"
         _StyleDefs(81)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(82)  =   "Named:id=36:Selected"
         _StyleDefs(83)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(84)  =   "Named:id=37:Caption"
         _StyleDefs(85)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(86)  =   "Named:id=38:HighlightRow"
         _StyleDefs(87)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(88)  =   "Named:id=39:EvenRow"
         _StyleDefs(89)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(90)  =   "Named:id=40:OddRow"
         _StyleDefs(91)  =   ":id=40,.parent=33"
         _StyleDefs(92)  =   "Named:id=41:RecordSelector"
         _StyleDefs(93)  =   ":id=41,.parent=34"
         _StyleDefs(94)  =   "Named:id=42:FilterBar"
         _StyleDefs(95)  =   ":id=42,.parent=33"
      End
      Begin Threed.SSCommand SSCommand3 
         Height          =   360
         Left            =   8160
         TabIndex        =   12
         ToolTipText     =   "Salir"
         Top             =   120
         Width           =   465
         _Version        =   65536
         _ExtentX        =   820
         _ExtentY        =   635
         _StockProps     =   78
         BevelWidth      =   1
         Picture         =   "FrmCtaCte.frx":0644
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   360
         Left            =   7440
         TabIndex        =   13
         ToolTipText     =   "Salir"
         Top             =   120
         Width           =   465
         _Version        =   65536
         _ExtentX        =   820
         _ExtentY        =   635
         _StockProps     =   78
         BevelWidth      =   1
         Picture         =   "FrmCtaCte.frx":0BDE
      End
      Begin VB.Label LblNomTrab 
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
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1320
         TabIndex        =   11
         Top             =   165
         Width           =   6015
      End
      Begin VB.Label LblCodTrab 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   165
         Width           =   1095
      End
   End
   Begin VB.Frame Frame7 
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
      Height          =   7170
      Left            =   0
      TabIndex        =   6
      Top             =   1080
      Width           =   8610
      Begin Threed.SSPanel PnlReporte 
         Height          =   1575
         Left            =   1920
         TabIndex        =   17
         Top             =   960
         Visible         =   0   'False
         Width           =   4695
         _Version        =   65536
         _ExtentX        =   8281
         _ExtentY        =   2778
         _StockProps     =   15
         BackColor       =   8421504
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
         BevelInner      =   2
         Begin VB.TextBox Txtano 
            Height          =   315
            Left            =   975
            MaxLength       =   4
            TabIndex        =   19
            Top             =   240
            Width           =   615
         End
         Begin VB.ComboBox Cmbmes 
            Height          =   315
            ItemData        =   "FrmCtaCte.frx":0EF8
            Left            =   1650
            List            =   "FrmCtaCte.frx":0F20
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   240
            Width           =   2415
         End
         Begin Threed.SSCommand SSCommand5 
            Height          =   600
            Left            =   960
            TabIndex        =   21
            ToolTipText     =   "Salir"
            Top             =   720
            Width           =   3105
            _Version        =   65536
            _ExtentX        =   5477
            _ExtentY        =   1058
            _StockProps     =   78
            BevelWidth      =   1
            Picture         =   "FrmCtaCte.frx":0F88
         End
         Begin Threed.SSCommand SSCommand6 
            Height          =   360
            Left            =   4280
            TabIndex        =   22
            ToolTipText     =   "Salir"
            Top             =   80
            Width           =   345
            _Version        =   65536
            _ExtentX        =   609
            _ExtentY        =   635
            _StockProps     =   78
            Caption         =   "X"
            ForeColor       =   255
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   1
            Picture         =   "FrmCtaCte.frx":12A2
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000C&
            Caption         =   "Peiodo"
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
            Left            =   240
            TabIndex        =   20
            Top             =   280
            Width           =   600
         End
      End
      Begin TrueOleDBGrid70.TDBGrid DGrdCtaCte 
         Height          =   6825
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   8385
         _ExtentX        =   14790
         _ExtentY        =   12039
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Codigo"
         Columns(0).DataField=   "codigo"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Nombre"
         Columns(1).DataField=   "nombre"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Saldo"
         Columns(2).DataField=   "saldo"
         Columns(2).NumberFormat=   "General Number"
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
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1826"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1746"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=9525"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=9446"
         Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(9)=   "Column(2).Width=2355"
         Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=2275"
         Splits(0)._ColumnProps(12)=   "Column(2)._ColStyle=2"
         Splits(0)._ColumnProps(13)=   "Column(2).Order=3"
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
         _StyleDefs(39)  =   "Splits(0).Columns(0).Style:id=46,.parent=13"
         _StyleDefs(40)  =   "Splits(0).Columns(0).HeadingStyle:id=43,.parent=14"
         _StyleDefs(41)  =   "Splits(0).Columns(0).FooterStyle:id=44,.parent=15"
         _StyleDefs(42)  =   "Splits(0).Columns(0).EditorStyle:id=45,.parent=17"
         _StyleDefs(43)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
         _StyleDefs(44)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
         _StyleDefs(45)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
         _StyleDefs(46)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
         _StyleDefs(47)  =   "Splits(0).Columns(2).Style:id=32,.parent=13,.alignment=1"
         _StyleDefs(48)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
         _StyleDefs(49)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
         _StyleDefs(50)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
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
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Tipo Trabajador"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   1125
   End
End
Attribute VB_Name = "FrmCtaCte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ChkCeros_Click()
Carga_Saldos_CtaCte
End Sub

Private Sub Cmbcia_Click()
wcia = Trim(Right("00" & Cmbcia.ItemData(Cmbcia.ListIndex), 2))
Call fc_Descrip_Maestros2("01055", "", Cmbtipotrabajador)
Cmbtipotrabajador.Enabled = True
End Sub

Private Sub Cmbtipotrabajador_Click()
Carga_Saldos_CtaCte
End Sub

Private Sub DGrdCtaCte_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
Cancel = True
End Sub

Private Sub DGrdCtaCte_DblClick()
Carga_Kardex (Trim(DgrdCtaCte.Columns(0) & ""))
End Sub
Public Sub Carga_Kardex(lCod As String)
On Error GoTo CORRIGE
If Trim(DgrdCtaCte.Columns(0) & "") <> "" Then
   Dim rsmovi As New ADODB.Recordset
   Sql$ = "Usp_Pla_Carga_movi_ctacte '" & wcia & "','" & lCod & "'"
   LblCodTrab.Caption = Trim(DgrdCtaCte.Columns(0) & "")
   LblNomTrab.Caption = Trim(DgrdCtaCte.Columns(1) & "")
   cn.CursorLocation = adUseClient
   Set rsmovi = New ADODB.Recordset
   Set rsmovi = cn.Execute(Sql$, 64)
   Set DgrdMovi.DataSource = rsmovi
   FrameKardex.ZOrder 0
   FrameKardex.Visible = True
End If
CORRIGE:
End Sub
Private Sub DgrdMovi_DblClick()
On Error GoTo CORRIGE
If Trim(DgrdMovi.Columns(8) & "") = "C" Then
   Frmplacte.Carga_Prestamo (Trim(DgrdMovi.Columns(7) & ""))
End If
CORRIGE:

End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
Me.Width = 8835
Me.Height = 8805

Lblfecha.Caption = Format(Date, "dd/mm/yyyy")
Call rCarCbo(Cmbcia, Carga_Cia, "C", "00")
Call rUbiIndCmbBox(Cmbcia, wcia, "00")
End Sub
Public Sub Carga_Saldos_CtaCte()
Dim rsSaldos As New ADODB.Recordset

Dim lCeros As String
lCeros = "N"
If ChkCeros.Value = 1 Then lCeros = "S"

Sql$ = "Usp_Pla_Carga_Saldos_Ctacte '" & wcia & "','" & fc_CodigoComboBox(Cmbtipotrabajador, 2) & "','*','" & lCeros & "'"

cn.CursorLocation = adUseClient
Set rsSaldos = New ADODB.Recordset
Set rsSaldos = cn.Execute(Sql$, 64)

Set DgrdCtaCte.DataSource = rsSaldos
End Sub
Public Sub Nuevo()

End Sub

Private Sub SSCommand1_Click()
Reporte_Kardex
End Sub

Private Sub SSCommand2_Click()
Reporte_Saldos
End Sub

Private Sub SSCommand3_Click()
FrameKardex.Visible = False
End Sub
Private Sub Reporte_Saldos()

Dim rs As Object
Dim xlApp2 As Excel.Application
Dim xlApp1  As Excel.Application
Dim xlBook As Excel.Workbook
Dim xlSheet As Excel.Worksheet

Dim I As Integer
Dim Fila As Integer

Set xlApp1 = CreateObject("Excel.Application")
xlApp1.Workbooks.Add
Set xlApp2 = xlApp1.Application
Set xlBook = xlApp2.Workbooks(1)
Set xlSheet = xlApp2.Worksheets("HOJA1")

xlSheet.Range("B:B").ColumnWidth = 50
xlSheet.Range("C:C").NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
xlSheet.Cells(1, 1).Value = Trim(Cmbcia.Text)
xlSheet.Cells(1, 1).Font.Bold = True
xlSheet.Cells(3, 2).Value = " SALDOS DE CUENTA CORRIENTE " & Trim(Cmbtipotrabajador.Text & "")
xlSheet.Cells(3, 3).Value = Format(Date, "dd/mm/yyyy")
xlSheet.Cells(3, 3).NumberFormat = "m/d/yyyy"
xlSheet.Cells(3, 2).Font.Bold = True
xlSheet.Cells(3, 2).HorizontalAlignment = xlCenter
'xlSheet.Range("B2:U2").Merge

Fila = 5
xlSheet.Cells(Fila, 1).Value = "Codigo"
xlSheet.Cells(Fila, 2).Value = "Nombre Trabajador"
xlSheet.Cells(Fila, 3).Value = "Saldo"
xlSheet.Range(xlSheet.Cells(Fila, 1), xlSheet.Cells(Fila, 3)).Borders.LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(Fila, 1), xlSheet.Cells(Fila, 3)).Font.Bold = True
xlSheet.Range(xlSheet.Cells(Fila, 1), xlSheet.Cells(Fila, 3)).HorizontalAlignment = xlCenter


Dim lCeros As String
lCeros = "N"
If ChkCeros.Value = 1 Then lCeros = "S"

Sql = "Usp_Pla_Carga_Saldos_Ctacte '" & wcia & "','" & fc_CodigoComboBox(Cmbtipotrabajador, 2) & "','*','" & lCeros & "'"

Fila = 6
Dim msum As Integer
msum = 1
If (fAbrRst(rs, Sql)) Then
    Do While Not rs.EOF
        xlSheet.Cells(Fila, 1).Value = Trim(rs!Codigo & "")
        xlSheet.Cells(Fila, 2).Value = Trim(rs!nombre & "")
        xlSheet.Cells(Fila, 3).Value = rs!saldo
        Fila = Fila + 1
        msum = msum + 1
        rs.MoveNext
    Loop
End If
Fila = Fila + 1
xlSheet.Cells(Fila, 3).Value = "=SUM(R[" & msum * -1 & "]C:R[-1]C)"

xlApp2.Application.ActiveWindow.DisplayGridLines = False
xlSheet.Range("A1:A1").Select
'xlApp2.ActiveWindow.Zoom = 80
xlApp2.Application.Visible = True

End Sub
Private Sub Reporte_Kardex()

Dim rs As Object
Dim xlApp2 As Excel.Application
Dim xlApp1  As Excel.Application
Dim xlBook As Excel.Workbook
Dim xlSheet As Excel.Worksheet

Dim I As Integer
Dim Fila As Integer

Set xlApp1 = CreateObject("Excel.Application")
xlApp1.Workbooks.Add
Set xlApp2 = xlApp1.Application
Set xlBook = xlApp2.Workbooks(1)
Set xlSheet = xlApp2.Worksheets("HOJA1")

xlSheet.Range("B:B").ColumnWidth = 20
xlSheet.Range("F:F").ColumnWidth = 9
xlSheet.Range("G:G").ColumnWidth = 30
xlSheet.Range("F:F").HorizontalAlignment = xlCenter
xlSheet.Cells(1, 1).Value = Trim(Cmbcia.Text)
xlSheet.Cells(1, 1).Font.Bold = True
xlSheet.Cells(3, 1).Value = " SALDOS DE CUENTA CORRIENTE "
xlSheet.Cells(1, 7).Value = Format(Date, "dd/mm/yyyy")
xlSheet.Cells(3, 2).Font.Bold = True
xlSheet.Cells(3, 2).HorizontalAlignment = xlCenter
xlSheet.Range("A3:G3").Merge

Fila = 5
xlSheet.Cells(Fila, 1).Value = "Fecha"
xlSheet.Cells(Fila, 2).Value = "Motivo"
xlSheet.Cells(Fila, 3).Value = "Cargo"
xlSheet.Cells(Fila, 4).Value = "Abono"
xlSheet.Cells(Fila, 5).Value = "Saldo"
xlSheet.Cells(Fila, 6).Value = "Semana"
xlSheet.Cells(Fila, 7).Value = "Referencia"
xlSheet.Range(xlSheet.Cells(Fila, 1), xlSheet.Cells(Fila, 7)).Borders.LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(Fila, 1), xlSheet.Cells(Fila, 7)).Font.Bold = True
xlSheet.Range(xlSheet.Cells(Fila, 1), xlSheet.Cells(Fila, 7)).HorizontalAlignment = xlCenter

xlSheet.Range("C:E").NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "

Sql$ = "Usp_Pla_Carga_movi_ctacte '" & wcia & "','" & LblCodTrab.Caption & "'"

Fila = 7
If (fAbrRst(rs, Sql)) Then
    xlSheet.Cells(Fila, 1).Value = LblCodTrab.Caption
    xlSheet.Cells(Fila, 2).Value = LblNomTrab
    Fila = Fila + 2
    Do While Not rs.EOF
        xlSheet.Cells(Fila, 1).Value = rs!fecha
        xlSheet.Cells(Fila, 2).Value = Trim(rs!motivo & "")
        xlSheet.Cells(Fila, 3).Value = rs!Cargo
        xlSheet.Cells(Fila, 4).Value = rs!abono
        xlSheet.Cells(Fila, 5).Value = rs!saldo
        xlSheet.Cells(Fila, 6).Value = rs!semana
        xlSheet.Cells(Fila, 7).Value = rs!refe
        Fila = Fila + 1
        rs.MoveNext
    Loop
End If

xlApp2.Application.ActiveWindow.DisplayGridLines = False
xlSheet.Range("A1:A1").Select
xlApp2.Application.Caption = "PROVISION DE CTS"
'xlApp2.ActiveWindow.Zoom = 80
xlApp2.Application.Visible = True

End Sub

Private Sub Reporte_Mensual()

Dim rs As Object
Dim xlApp2 As Excel.Application
Dim xlApp1  As Excel.Application
Dim xlBook As Excel.Workbook
Dim xlSheet As Excel.Worksheet

Dim I As Integer
Dim Fila As Integer

Set xlApp1 = CreateObject("Excel.Application")
xlApp1.Workbooks.Add
Set xlApp2 = xlApp1.Application
Set xlBook = xlApp2.Workbooks(1)
Set xlSheet = xlApp2.Worksheets("HOJA1")

xlSheet.Range("B:B").ColumnWidth = 50
xlSheet.Range("C:F").NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
xlSheet.Cells(1, 1).Value = Trim(Cmbcia.Text)
xlSheet.Cells(1, 1).Font.Bold = True
xlSheet.Cells(3, 1).Value = " SALDOS DE CUENTA CORRIENTE " & Trim(Cmbtipotrabajador.Text & "") & " - " & CmbMes.Text & " " & Txtano.Text
xlSheet.Range("A3:F3").Merge
xlSheet.Cells(3, 1).Font.Bold = True
xlSheet.Cells(3, 1).HorizontalAlignment = xlCenter

Fila = 5
xlSheet.Cells(Fila, 1).Value = "Codigo"
xlSheet.Cells(Fila, 2).Value = "Nombre Trabajador"
xlSheet.Cells(Fila, 3).Value = "Saldo Inicial"
xlSheet.Cells(Fila, 4).Value = "Cargo"
xlSheet.Cells(Fila, 5).Value = "Abono"
xlSheet.Cells(Fila, 6).Value = "Final"
xlSheet.Range(xlSheet.Cells(Fila, 1), xlSheet.Cells(Fila, 6)).Borders.LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(Fila, 1), xlSheet.Cells(Fila, 6)).Font.Bold = True
xlSheet.Range(xlSheet.Cells(Fila, 1), xlSheet.Cells(Fila, 6)).HorizontalAlignment = xlCenter

Sql$ = "Usp_Pla_Cta_Cte_MEs '" & wcia & "'," & CmbMes.ListIndex + 1 & "," & Txtano.Text & ",'" & Mid(Cmbtipotrabajador.Text, 1, 1) & "'"

Fila = 6
Dim msum As Integer
msum = 1
If (fAbrRst(rs, Sql)) Then
    Do While Not rs.EOF
        xlSheet.Cells(Fila, 1).Value = Trim(rs!PlaCod & "")
        xlSheet.Cells(Fila, 2).Value = Trim(rs!nombre & "")
        xlSheet.Cells(Fila, 3).Value = rs!saldoi
        xlSheet.Cells(Fila, 4).Value = rs!prestamo
        xlSheet.Cells(Fila, 5).Value = rs!descuento
        xlSheet.Cells(Fila, 6).Value = rs!Final
        Fila = Fila + 1
        msum = msum + 1
        rs.MoveNext
    Loop
End If

xlSheet.Cells(Fila, 3).Value = "=SUM(R[" & msum * -1 & "]C:R[-1]C)"
xlSheet.Cells(Fila, 4).Value = "=SUM(R[" & msum * -1 & "]C:R[-1]C)"
xlSheet.Cells(Fila, 5).Value = "=SUM(R[" & msum * -1 & "]C:R[-1]C)"
xlSheet.Cells(Fila, 6).Value = "=SUM(R[" & msum * -1 & "]C:R[-1]C)"

xlSheet.Range(xlSheet.Cells(Fila, 3), xlSheet.Cells(Fila, 6)).Borders.LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(Fila, 3), xlSheet.Cells(Fila, 6)).Font.Bold = True


xlApp2.Application.ActiveWindow.DisplayGridLines = False
xlSheet.Range("A1:A1").Select
'xlApp2.ActiveWindow.Zoom = 80
xlApp2.Application.Visible = True

End Sub

Private Sub SSCommand4_Click()
Txtano.Text = Year(Date)
CmbMes.ListIndex = Month(Date) - 1
PnlReporte.Visible = True
End Sub

Private Sub SSCommand5_Click()

If Trim(Cmbtipotrabajador.Text & "") = "" Then
   MsgBox "Seleccione Tipo de Trabajador", vbInformation
   Exit Sub
End If

If Not IsNumeric(Txtano.Text) Then
   MsgBox "Ingrese Año Correctamente", vbInformation
   Exit Sub
End If

If Val(Txtano.Text) < 2013 Then
   MsgBox "Reporte solo esta disponible desde Setiembre del 2013", vbInformation
   Exit Sub
End If

If Me.CmbMes.ListIndex < 0 Then
   MsgBox "Selecciones Mes", vbInformation
   Exit Sub
End If

If Val(Txtano.Text) = 2013 And Me.CmbMes.ListIndex < 9 Then
   MsgBox "Reporte solo esta disponible desde Setiembre del 2013", vbInformation
   Exit Sub
End If

Reporte_Mensual
PnlReporte.Visible = False
End Sub

Private Sub SSCommand6_Click()
PnlReporte.Visible = False
End Sub
