VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form Frmplacte 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Prestamos al Trabajador"
   ClientHeight    =   7140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9690
   Icon            =   "Frmplacte.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7140
   ScaleWidth      =   9690
   Begin VB.Frame Frame3 
      Caption         =   "Prestamo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   0
      TabIndex        =   11
      Top             =   1320
      Width           =   9495
      Begin VB.TextBox TxtSemana 
         Height          =   285
         Left            =   5880
         TabIndex        =   31
         Top             =   1065
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox TxtRefe 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         MaxLength       =   60
         TabIndex        =   24
         Top             =   1080
         Width           =   6375
      End
      Begin VB.OptionButton OptGrati 
         Caption         =   "Gratificación"
         Height          =   195
         Left            =   1560
         TabIndex        =   20
         Top             =   1080
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.OptionButton OptNormal 
         Caption         =   "Normal"
         Height          =   195
         Left            =   3000
         TabIndex        =   18
         Top             =   1080
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox TxtImporte 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3360
         TabIndex        =   15
         Top             =   360
         Width           =   975
      End
      Begin MSComCtl2.DTPicker dtfecha 
         Height          =   315
         Left            =   960
         TabIndex        =   12
         Top             =   330
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   59768833
         CurrentDate     =   39085
      End
      Begin Threed.SSCommand BtnEliminar 
         Height          =   825
         Left            =   8160
         TabIndex        =   38
         ToolTipText     =   "Salir"
         Top             =   360
         Visible         =   0   'False
         Width           =   825
         _Version        =   65536
         _ExtentX        =   1455
         _ExtentY        =   1455
         _StockProps     =   78
         Caption         =   "Eliminar"
         BevelWidth      =   1
         Picture         =   "Frmplacte.frx":030A
      End
      Begin VB.OptionButton OptUtilidad 
         Caption         =   "Utilidad"
         Height          =   195
         Left            =   3960
         TabIndex        =   46
         Top             =   1080
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.OptionButton OptVaca 
         Caption         =   "Vacaciones"
         Height          =   195
         Left            =   240
         TabIndex        =   19
         Top             =   1080
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Referencia"
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   1095
         Width           =   780
      End
      Begin VB.Label Label8 
         Caption         =   "Saldo Actual"
         Height          =   255
         Left            =   5040
         TabIndex        =   40
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label LblSaldo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Height          =   255
         Left            =   5040
         TabIndex        =   39
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label LblSemana 
         AutoSize        =   -1  'True
         Caption         =   "Semana"
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
         Left            =   5040
         TabIndex        =   32
         Top             =   1080
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Boleta en la que Figurara el Prestamo"
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
         TabIndex        =   17
         Top             =   720
         Visible         =   0   'False
         Width           =   3210
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Importe"
         Height          =   195
         Left            =   2760
         TabIndex        =   16
         Top             =   360
         Width           =   525
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Fecha :"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   540
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9495
      Begin VB.ComboBox Cmbcia 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   80
         Width           =   4725
      End
      Begin VB.Label Lblfecha 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
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
         Left            =   5640
         TabIndex        =   4
         Top             =   120
         Width           =   1815
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
         ForeColor       =   &H80000014&
         Height          =   210
         Left            =   240
         TabIndex        =   2
         Top             =   120
         Width           =   825
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Acuerdo de Descuentos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   0
      TabIndex        =   14
      Top             =   3240
      Width           =   9495
      Begin VB.TextBox TxtSubsidio 
         Height          =   285
         Left            =   4560
         TabIndex        =   47
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox TxtUtilidad 
         Height          =   285
         Left            =   3480
         TabIndex        =   44
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox TxtNormal 
         Height          =   285
         Left            =   2400
         TabIndex        =   23
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox TxtGrati 
         Height          =   285
         Left            =   1320
         TabIndex        =   22
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox TxtVaca 
         Height          =   285
         Left            =   240
         TabIndex        =   21
         Top             =   600
         Width           =   975
      End
      Begin MSComCtl2.DTPicker dtFechaPago 
         Height          =   315
         Left            =   5880
         TabIndex        =   29
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   59768833
         CurrentDate     =   39085
      End
      Begin TrueOleDBGrid70.TDBGrid DgrdMovi 
         Height          =   2745
         Left            =   0
         TabIndex        =   41
         Top             =   1080
         Width           =   9465
         _ExtentX        =   16695
         _ExtentY        =   4842
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Desde"
         Columns(0).DataField=   "Fec_Desde"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Normal"
         Columns(1).DataField=   "normal"
         Columns(1).NumberFormat=   "General Number"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Vacaciones"
         Columns(2).DataField=   "vaca"
         Columns(2).NumberFormat=   "General Number"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Gratificación"
         Columns(3).DataField=   "grati"
         Columns(3).NumberFormat=   "General Number"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Utilidades"
         Columns(4).DataField=   "util"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Id"
         Columns(5).DataField=   "Id_dcto"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Subsidio"
         Columns(6).DataField=   "Subsidio"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   7
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).FetchRowStyle=   -1  'True
         Splits(0).DividerColor=   13160660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=7"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=3069"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2990"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=2540"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2461"
         Splits(0)._ColumnProps(8)=   "Column(1)._ColStyle=2"
         Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(10)=   "Column(2).Width=2434"
         Splits(0)._ColumnProps(11)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(12)=   "Column(2)._WidthInPix=2355"
         Splits(0)._ColumnProps(13)=   "Column(2)._ColStyle=2"
         Splits(0)._ColumnProps(14)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(15)=   "Column(3).Width=2487"
         Splits(0)._ColumnProps(16)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(17)=   "Column(3)._WidthInPix=2408"
         Splits(0)._ColumnProps(18)=   "Column(3)._ColStyle=2"
         Splits(0)._ColumnProps(19)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(20)=   "Column(4).Width=2725"
         Splits(0)._ColumnProps(21)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(22)=   "Column(4)._WidthInPix=2646"
         Splits(0)._ColumnProps(23)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(24)=   "Column(5).Width=2725"
         Splits(0)._ColumnProps(25)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(26)=   "Column(5)._WidthInPix=2646"
         Splits(0)._ColumnProps(27)=   "Column(5).Visible=0"
         Splits(0)._ColumnProps(28)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(29)=   "Column(6).Width=2725"
         Splits(0)._ColumnProps(30)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(31)=   "Column(6)._WidthInPix=2646"
         Splits(0)._ColumnProps(32)=   "Column(6).Order=7"
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
         _StyleDefs(26)  =   "Splits(0).CaptionStyle:id=22,.parent=4,.alignment=1,.bgcolor=&HFF8000&"
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
         _StyleDefs(43)  =   "Splits(0).Columns(1).Style:id=28,.parent=13,.alignment=1"
         _StyleDefs(44)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
         _StyleDefs(45)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
         _StyleDefs(46)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17,.alignment=1"
         _StyleDefs(47)  =   "Splits(0).Columns(2).Style:id=50,.parent=13,.alignment=1,.fgcolor=&HFF&"
         _StyleDefs(48)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=14"
         _StyleDefs(49)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=15"
         _StyleDefs(50)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=17"
         _StyleDefs(51)  =   "Splits(0).Columns(3).Style:id=32,.parent=13,.alignment=1,.fgcolor=&HFF0000&"
         _StyleDefs(52)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
         _StyleDefs(53)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
         _StyleDefs(54)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17,.fgcolor=&H80000001&"
         _StyleDefs(55)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
         _StyleDefs(56)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
         _StyleDefs(57)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
         _StyleDefs(58)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
         _StyleDefs(59)  =   "Splits(0).Columns(5).Style:id=62,.parent=13"
         _StyleDefs(60)  =   "Splits(0).Columns(5).HeadingStyle:id=59,.parent=14"
         _StyleDefs(61)  =   "Splits(0).Columns(5).FooterStyle:id=60,.parent=15"
         _StyleDefs(62)  =   "Splits(0).Columns(5).EditorStyle:id=61,.parent=17"
         _StyleDefs(63)  =   "Splits(0).Columns(6).Style:id=58,.parent=13"
         _StyleDefs(64)  =   "Splits(0).Columns(6).HeadingStyle:id=55,.parent=14"
         _StyleDefs(65)  =   "Splits(0).Columns(6).FooterStyle:id=56,.parent=15"
         _StyleDefs(66)  =   "Splits(0).Columns(6).EditorStyle:id=57,.parent=17"
         _StyleDefs(67)  =   "Named:id=33:Normal"
         _StyleDefs(68)  =   ":id=33,.parent=0"
         _StyleDefs(69)  =   "Named:id=34:Heading"
         _StyleDefs(70)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(71)  =   ":id=34,.wraptext=-1"
         _StyleDefs(72)  =   "Named:id=35:Footing"
         _StyleDefs(73)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(74)  =   "Named:id=36:Selected"
         _StyleDefs(75)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(76)  =   "Named:id=37:Caption"
         _StyleDefs(77)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(78)  =   "Named:id=38:HighlightRow"
         _StyleDefs(79)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(80)  =   "Named:id=39:EvenRow"
         _StyleDefs(81)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(82)  =   "Named:id=40:OddRow"
         _StyleDefs(83)  =   ":id=40,.parent=33"
         _StyleDefs(84)  =   "Named:id=41:RecordSelector"
         _StyleDefs(85)  =   ":id=41,.parent=34"
         _StyleDefs(86)  =   "Named:id=42:FilterBar"
         _StyleDefs(87)  =   ":id=42,.parent=33"
      End
      Begin Threed.SSCommand BtnPlan 
         Height          =   345
         Left            =   8160
         TabIndex        =   42
         ToolTipText     =   "Salir"
         Top             =   600
         Visible         =   0   'False
         Width           =   945
         _Version        =   65536
         _ExtentX        =   1667
         _ExtentY        =   609
         _StockProps     =   78
         Caption         =   "Eliminar"
         BevelWidth      =   1
         Picture         =   "Frmplacte.frx":075C
      End
      Begin Threed.SSCommand BtnPlanNew 
         Height          =   345
         Left            =   8160
         TabIndex        =   43
         ToolTipText     =   "Salir"
         Top             =   240
         Visible         =   0   'False
         Width           =   945
         _Version        =   65536
         _ExtentX        =   1667
         _ExtentY        =   609
         _StockProps     =   78
         Caption         =   "Nuevo"
         BevelWidth      =   1
         Picture         =   "Frmplacte.frx":0778
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Descontar Desde"
         Height          =   195
         Left            =   5760
         TabIndex        =   30
         Top             =   360
         Width           =   1605
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Importe Subsidio"
         Height          =   195
         Left            =   4440
         TabIndex        =   48
         Top             =   360
         Width           =   1170
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Utilidad"
         Height          =   195
         Left            =   3480
         TabIndex        =   45
         Top             =   360
         Width           =   525
      End
      Begin VB.Label LblIdDcto 
         Height          =   255
         Left            =   5760
         TabIndex        =   37
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Normal"
         Height          =   195
         Left            =   2400
         TabIndex        =   28
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Gratificación"
         Height          =   195
         Left            =   1320
         TabIndex        =   27
         Top             =   360
         Width           =   885
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Vacaciones"
         Height          =   195
         Left            =   240
         TabIndex        =   26
         Top             =   360
         Width           =   840
      End
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   30
      TabIndex        =   5
      Top             =   450
      Width           =   9495
      Begin VB.OptionButton OptPago 
         Caption         =   "Pago"
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
         Height          =   195
         Left            =   6240
         TabIndex        =   36
         Top             =   440
         Width           =   975
      End
      Begin VB.OptionButton OptPrestamo 
         Caption         =   "Prestamo"
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
         Left            =   6240
         TabIndex        =   35
         Top             =   120
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.TextBox Txtautoriza 
         Height          =   285
         Left            =   1680
         TabIndex        =   9
         Top             =   3240
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox Txtcodpla 
         Height          =   285
         Left            =   960
         TabIndex        =   1
         Top             =   240
         Width           =   855
      End
      Begin VB.Label LblID 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   7320
         TabIndex        =   34
         Top             =   480
         Width           =   45
      End
      Begin VB.Label LblTipoTrab 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   120
         TabIndex        =   33
         Top             =   480
         Width           =   45
      End
      Begin VB.Label Lblautoriza 
         BackColor       =   &H8000000E&
         Height          =   255
         Left            =   3000
         TabIndex        =   10
         Top             =   3240
         Visible         =   0   'False
         Width           =   4455
      End
      Begin VB.Label Lblnombre 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   285
         Left            =   1845
         TabIndex        =   8
         Top             =   240
         Width           =   3930
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Trabajador"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   765
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Autorizado"
         Height          =   195
         Left            =   840
         TabIndex        =   6
         Top             =   3240
         Visible         =   0   'False
         Width           =   750
      End
   End
End
Attribute VB_Name = "Frmplacte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim VTipo As String
Dim VBanco As String
Dim VModo As String
Dim lAccion As Integer
Public proc As String
Public lCod As String

Private Sub BtnEliminar_Click()
lAccion = 3
Graba_Prestamo
End Sub

Private Sub BtnPlan_Click()
If Not IsNumeric(LblIdDcto.Caption) Then Exit Sub
Sql$ = "Update pla_dcto_ctacte Set status='*',user_modi='" & wuser & "',fec_modi=getdate() where Id_dcto=" & LblIdDcto.Caption & ""
cn.Execute Sql$
Carga_Plan (Trim(Txtcodpla.Text))
Nuevo_Plan
End Sub

Private Sub BtnPlanNew_Click()
Nuevo_Plan
End Sub
Private Sub Nuevo_Plan()
LblIdDcto.Caption = "0"
TxtNormal.Text = ""
TxtVaca.Text = ""
TxtGrati.Text = ""
TxtSubsidio.Text = ""
dtFechaPago.Value = Date
BtnPlan.Visible = False
End Sub
Private Sub Cmbcia_Click()
wcia = Trim(Right("00" & CmbCia.ItemData(CmbCia.ListIndex), 2))
End Sub

Private Sub Cmbcta_Click()
If (Right(Cmbcta.Text, 3) <> Mid(Cmbmoneda.Text, 2, 3)) And Cmbcta.ListIndex > -1 Then
  MsgBox "Moneda de Cuenta debe ser igual a moneda de prestamo"
  Cmbcta.ListIndex = -1
End If
TxtDocu.SetFocus
End Sub

Private Sub Cmbmodo_Click()
VModo = Left(Cmbmodo.Text, 1)
End Sub

Private Sub CmbTipo_Click()
VTipo = Left(Cmbtipo.Text, 2)
Label11.Visible = True
Select Case Left(Cmbtipo.Text, 2)
       Case Is = "EF", "FF"
            Cmbbanco.ListIndex = -1
            Cmbcta.ListIndex = -1
            Cmbbanco.Enabled = False
            Cmbcta.Enabled = False
            Txtmotivo.SetFocus
       Case Is = "DR", "CH"
            Cmbbanco.ListIndex = -1
            Cmbcta.ListIndex = -1
            Cmbbanco.Enabled = True
            Cmbcta.Enabled = True
            TxtDocu.Enabled = True
            Cmbbanco.SetFocus
       Case Is = "DB"
            Frame3.Visible = True
            Label11.Visible = False
            
End Select
End Sub

Private Sub DgrdMovi_DblClick()
On Error GoTo CORRIGE
If IsNumeric(DgrdMovi.Columns(5)) Or IsNumeric(DgrdMovi.Columns(6)) Then
   LblIdDcto.Caption = DgrdMovi.Columns(5)
   TxtVaca.Text = Trim(DgrdMovi.Columns(2))
   TxtGrati.Text = Trim(DgrdMovi.Columns(3))
   TxtNormal.Text = Trim(DgrdMovi.Columns(1))
   TxtUtilidad.Text = Trim(DgrdMovi.Columns(5))
   TxtSubsidio.Text = Trim(DgrdMovi.Columns(6))
   dtFechaPago.Value = Trim(DgrdMovi.Columns(0))
   BtnPlan.Visible = True
   BtnPlanNew.Visible = True
End If
CORRIGE:
End Sub

Private Sub Form_Activate()
If Txtcodpla.Enabled = True Then Txtcodpla.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
        If KeyAscii = 27 Then
           End
        End If
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
Me.Height = 7620
Me.Width = 9785
Lblfecha.Caption = Format(Date, "dd/mm/yyyy")
dtfecha.Value = Format(Date, "dd/mm/yyyy")
dtFechaPago.Value = Format(Date, "dd/mm/yyyy")
Call rCarCbo(CmbCia, Carga_Cia, "C", "00")
Call rUbiIndCmbBox(CmbCia, wcia, "00")
proc = 0
Me.KeyPreview = True
End Sub
Private Sub Form_Unload(Cancel As Integer)
If proc = 0 Then Call FrmCtaCte.Carga_Saldos_CtaCte
'If proc = 1 Then Call llena_grid
End Sub

Private Sub OptGrati_Click()
If OptGrati.Value = True Then
   Lblsemana.Visible = False
   Txtsemana.Text = ""
   Txtsemana.Visible = False
End If
End Sub

Private Sub OptNormal_Click()
If OptNormal.Value = True And LblTipoTrab = "02" Then
   'LblSemana.Visible = True
   'TxtSemana.Visible = True
End If
End Sub

Private Sub OptPago_Click()
If OptPago.Value = True Then
   Label3.Visible = False
   OptVaca.Visible = False
   OptGrati.Visible = False
   OptNormal.Visible = False
   Lblsemana.Visible = False
   Txtsemana.Visible = False
   Frame4.Visible = False
End If
End Sub

Private Sub OptPrestamo_Click()
   'Label3.Visible = True
   'OptVaca.Visible = True
   'OptGrati.Visible = True
   'OptNormal.Visible = True
   'If OptNormal.Value = True And Lbltipotrab = "02" Then
   '   Lblsemana.Visible = True
   '   Txtsemana.Visible = True
   'Else
   '   Lblsemana.Visible = False
   '   Txtsemana.Visible = False
   'End If
   Frame4.Visible = True
End Sub

Private Sub OptVaca_Click()
If OptVaca.Value = True Then
   Lblsemana.Visible = False
   Txtsemana.Text = ""
   Txtsemana.Visible = False
End If
End Sub

Private Sub Txtautoriza_GotFocus()
Lblaut = "S"
End Sub

Private Sub Txtautoriza_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Txtcodpla.SetFocus
End Sub

Private Sub Txtautoriza_LostFocus()
cn.CursorLocation = adUseClient
Set rs = New ADODB.Recordset
Sql$ = nombre()
Sql$ = Sql$ & "placod from planillas where status<>'*' " _
     & "and cia='" & wcia & "' AND placod='" & Trim(Txtautoriza.Text) & "'"

cn.CursorLocation = adUseClient
Set rs = New ADODB.Recordset
Set rs = cn.Execute(Sql$)
If rs.RecordCount > 0 Then
   Lblautoriza.Caption = rs!nombre
ElseIf Trim(Txtcodpla.Text) <> "" Then
   MsgBox "Codigo de Planilla No Existe ", vbExclamation, "Codigo N° => " & Txtcodpla.Text
   Lblautoriza.Caption = ""
   Txtautoriza.SetFocus
End If

End Sub

Private Sub Txtcodpla_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 Then
    Frmgrdpla.Show
    Frmgrdpla.ZOrder 0
End If
End Sub

Private Sub Txtcodpla_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Txtimporte.SetFocus
End Sub

Private Sub Txtcodpla_LostFocus()
Busca_Trabajador
End Sub
Private Sub Busca_Trabajador()
LblSaldo.Caption = "0.00"
Set DgrdMovi.DataSource = Nothing
cn.CursorLocation = adUseClient
Set rs = New ADODB.Recordset
Sql$ = nombre()
Sql$ = Sql$ & "placod,tipotrabajador from planillas where status<>'*' " _
     & "and cia='" & wcia & "' AND placod='" & Trim(Txtcodpla.Text) & "'"

cn.CursorLocation = adUseClient
Set rs = New ADODB.Recordset
Set rs = cn.Execute(Sql$)
Lblsemana.Visible = False: Txtsemana.Visible = False
If rs.RecordCount > 0 Then
   Lblnombre.Caption = Space(5) & Trim(rs!nombre)
   LblTipoTrab.Caption = Trim(rs!TipoTrabajador & "")
   
   Dim rssaldo As New ADODB.Recordset
   Sql$ = "Usp_Pla_Carga_Saldos_Ctacte '" & wcia & "','" & LblTipoTrab.Caption & "','" & Trim(Txtcodpla.Text) & "','N'"
   cn.CursorLocation = adUseClient
   Set rssaldo = New ADODB.Recordset
   Set rssaldo = cn.Execute(Sql$, 64)
   If rssaldo.RecordCount > 0 Then LblSaldo.Caption = rssaldo!saldo
   rssaldo.Close: Set rssaldo = Nothing
   
   Carga_Plan (Trim(Txtcodpla.Text))
   
   'If Trim(rs!TipoTrabajador & "") = "02" And OptPrestamo.Value And OptNormal.Value Then LblSemana.Visible = True: TxtSemana.Visible = True
   Txtimporte.SetFocus
ElseIf Trim(Txtcodpla.Text) <> "" Then
   MsgBox "Codigo de Planilla No Existe ", vbExclamation, "Codigo N° => " & Txtcodpla.Text
   Lblnombre.Caption = ""
   'Lblcte.Caption = "0.00"
   Txtcodpla.SetFocus
End If
End Sub
Public Sub Carga_Plan(lCod As String)
On Error GoTo CORRIGE
If Trim(Txtcodpla.Text & "") <> "" Then
   Dim rsplan As New ADODB.Recordset
   Sql$ = "Usp_Pla_Plan_Dcto '" & wcia & "','" & lCod & "'"
   cn.CursorLocation = adUseClient
   Set rsplan = New ADODB.Recordset
   Set rsplan = cn.Execute(Sql$, 64)
   Set DgrdMovi.DataSource = rsplan
End If
CORRIGE:
End Sub


Private Sub Txtcuotas_KeyPress(KeyAscii As Integer)
Txtcuotas.Text = Txtcuotas.Text + fc_ValNumeros(KeyAscii)
If KeyAscii = 13 Then Cmbmodo.SetFocus
End Sub

Private Sub TxtDocu_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Txtmotivo.SetFocus
End Sub

Private Sub Txtfecha_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then Txtautoriza.SetFocus
End Sub

Private Sub Txtimporte_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
       Case 8, 13, 27, 32, 42, 44, 45, 46, 48 To 57
            KeyAscii = KeyAscii
       Case Else
            KeyAscii = 0
            Exit Sub
End Select
If KeyAscii = 13 Then
   If Txtsemana.Visible = True Then Txtsemana.SetFocus Else Txtrefe.SetFocus
End If
End Sub

Private Sub Txtimporte_LostFocus()
Txtimporte = Format(Txtimporte, "###,###.00")
End Sub
Public Function Graba_Prestamo()
Dim mano As Integer
Dim mmes As Integer
Dim mdia As Integer
Dim NroTrans As Integer
Dim lSaldo As Currency
lSaldo = 0
If IsNumeric(LblSaldo.Caption) Then lSaldo = LblSaldo.Caption
Txtcodpla.Text = UCase(Txtcodpla.Text)
'On Error GoTo CORRIGE
NroTrans = 0
If Not IsNumeric(Txtimporte.Text) Then Txtimporte.Text = "0"
If Trim(Txtimporte) = "" Then Txtimporte = "0.00"
If Trim(Txtcodpla.Text) = "" Then MsgBox "Debe Indicar el Trabajador", vbCritical, TitMsg: Txtcodpla.SetFocus: Exit Function
If CCur(Txtimporte.Text) = 0 Then MsgBox "Debe Indicar Importe", vbCritical, TitMsg: Txtimporte.SetFocus: Exit Function

If CCur(Txtimporte.Text) < 0 Then
   MsgBox "El Importe no Puede ser Negativo", vbCritical, TitMsg: Txtimporte.SetFocus: Exit Function
End If
If Txtsemana.Visible = True Then
   If Not Verifica_Semana Then
      Exit Function
   End If
End If
If lAccion = 3 Then
   Mgrab = MsgBox("Seguro de Eliminar Movimiento", vbYesNo + vbQuestion, TitMsg)
Else
   Mgrab = MsgBox("Seguro de Grabar Movimiento", vbYesNo + vbQuestion, TitMsg)
End If
If Mgrab <> 6 Then Exit Function
Screen.MousePointer = vbArrowHourglass

cn.BeginTrans
NroTrans = 1

Dim lMovi As String
Dim lProceso As String
Dim lSemana As String

lProceso = "": lSemana = ""
If OptPago Then
   lMovi = "S"
Else
   lMovi = "I"
   If OptGrati.Value Then lProceso = "03"
   If OptVaca.Value Then lProceso = "02"
   If OptUtilidad.Value Then lProceso = "11"
   If OptNormal.Value Then
      lProceso = "01"
      If Txtsemana.Visible = True Then lSemana = Trim(Txtsemana.Text & "")
   End If
End If

Dim lVaca As Double
Dim lGrati As Double
Dim lNormal As Double
Dim lUtil As Double
Dim lSubsidio As Double
lVaca = 0: lGrati = 0: lNormal = 0: lUtil = 0: lSubsidio = 0

If OptPrestamo.Value = True Then
   If Not IsNumeric(TxtVaca.Text) Then TxtVaca.Text = "0"
   If Not IsNumeric(TxtGrati.Text) Then TxtGrati.Text = "0"
   If Not IsNumeric(TxtNormal.Text) Then TxtNormal.Text = "0"
    If Not IsNumeric(TxtUtilidad.Text) Then TxtUtilidad.Text = "0"
    If Not IsNumeric(TxtSubsidio.Text) Then TxtSubsidio.Text = "0"
    lVaca = CCur(TxtVaca.Text)
    lGrati = CCur(TxtGrati.Text)
    lNormal = CCur(TxtNormal.Text)
    lUtil = CCur(TxtUtilidad.Text)
    lSubsidio = CCur(TxtSubsidio.Text)
End If
If Trim(LblId.Caption & "") = "" Then LblId.Caption = "0"
Sql$ = "Usp_Pla_movi_ctacte_Reg " & LblId.Caption & ",'" & wcia & "','" & Txtcodpla.Text & "','C','" & dtfecha.Value & "','" & lMovi & "'," & CCur(Txtimporte.Text) & ",'" & lProceso & "','" & lSemana & "',0,'" & wuser & "','" & Trim(Txtrefe.Text) & "'," & lVaca & "," & lGrati & "," & lNormal & "," & lUtil & ",'" & dtFechaPago.Value & "'," & LblIdDcto.Caption & "," & lAccion & "," & lSaldo & "," & lSubsidio & ""
cn.Execute Sql$

cn.CommitTrans

Frame2.Enabled = False
Screen.MousePointer = vbDefault
MsgBox "Datos Grabados Correctamente", vbInformation + vbOKOnly, Me.Caption
'If proc = 1 Then Unload Me
Unload Me
'If FrmCtaCte.FrameKardex.Visible Then
'   FrmCtaCte.Carga_Kardex (FrmCtaCte.LblCodTrab.Caption)
'Else
'   FrmCtaCte.Carga_Saldos_CtaCte
'End If
FrmCtaCte.FrameKardex.Visible = False
FrmCtaCte.Carga_Saldos_CtaCte
Exit Function
CORRIGE:
  If NroTrans = 1 Then
    cn.RollbackTrans
  End If
  MsgBox "Error : " & Err.Description, vbCritical, Me.Caption
End Function

Public Sub Nuevo_Prestamo()
Frame2.Enabled = True
If Trim(lCod & "") <> "" Then Txtcodpla.Text = lCod
LIMPIA
OptNormal.Value = True
LblIdDcto.Caption = 0
BtnEliminar.Visible = False
End Sub
Public Sub Carga_Prestamo(lId As Integer)
LIMPIA
LblId.Caption = lId
Sql$ = "Select * from pla_movi_ctacte where Id_CtaCte=" & LblId.Caption & ""
cn.CursorLocation = adUseClient
Set rs = New ADODB.Recordset
Set rs = cn.Execute(Sql$, 64)
rs.MoveFirst
Txtcodpla.Text = Trim(rs!PlaCod & "")
dtfecha.Value = rs!fecha
Txtimporte.Text = rs!importe
Txtrefe.Text = Trim(rs!referencia & "")
Txtsemana.Text = Trim(rs!semana & "")
If Trim(rs!Proceso & "") = "01" Then OptNormal.Value = True
If Trim(rs!Proceso & "") = "02" Then OptVaca.Value = True
If Trim(rs!Proceso & "") = "03" Then OptGrati.Value = True
If Trim(rs!Proceso & "") = "11" Then OptUtilidad.Value = True
If rs!tipo = "C" And rs!movi = "I" Then
'   OptPrestamo.Value = True
'
'   Dim rsmovi As New ADODB.Recordset
'   Sql$ = "select * from pla_dcto_ctacte where cia='" & wcia & "' and placod='" & Txtcodpla.Text & "' and status<>'*'"
'   cn.CursorLocation = adUseClient
'   Set rsmovi = New ADODB.Recordset
'   Set rsmovi = cn.Execute(Sql$, 64)
'   If rsmovi.RecordCount > 0 Then
'      TxtVaca.Text = rsmovi!vaca
'      TxtGrati.Text = rsmovi!grati
'      TxtNormal.Text = rsmovi!Normal
'      dtFechaPago.Value = rsmovi!fec_desde
'      LblIdDcto.Caption = rsmovi!Id_dcto
'   End If
'   rsmovi.Close: Set rsmovi = Nothing
Else
   OptPago.Value = True
End If
Txtimporte.SetFocus
Txtcodpla.Enabled = False
OptPago.Enabled = False
OptPrestamo.Enabled = False
Busca_Trabajador
BtnEliminar.Visible = True
If rs.State = 1 Then rs.Close
End Sub

Private Sub Txtmotivo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Txtimporte.SetFocus
End Sub

Function DEVUELVE_FECHA() As String
'On Error GoTo CORRIGE
Dim cadfecha As String
Dim con As String

If Val(Txtsemana) = 0 Then
   MsgBox "Ingrese La Semana", vbInformation, Me.Caption
   Exit Function
End If

If Trim(Txtcodpla) = "" Then
   MsgBox "Ingrese El Codigo del Trabajador", vbInformation, Me.Caption
   Exit Function
End If

   cadfecha = Str(Format(DTPFIN, "YYYY")) & "-" & Trim(Str(Format(Txtsemana, "00")))
   DEVUELVE_FECHA = cadfecha
End Function

Private Sub UpDown1_DownClick()
If Trim(Txtsemana.Text) = "" Then Txtsemana.Text = "0"
If Txtsemana.Text > 0 Then Txtsemana = Txtsemana - 1
End Sub

Private Sub UpDown1_UpClick()
If Trim(Txtsemana.Text) = "" Then Txtsemana.Text = "0"
Txtsemana = Txtsemana + 1
End Sub
Sub LIMPIA()
Txtcodpla.Text = ""
Lblnombre.Caption = ""
Txtimporte = ""
Txtrefe.Text = ""
Txtsemana.Text = ""
LblTipoTrab.Caption = ""
LblIdDcto.Caption = "0"
OptNormal.Value = True

Txtcodpla.Enabled = True
OptPago.Enabled = True
OptPrestamo.Enabled = True
TxtGrati.Text = ""
TxtVaca.Text = ""
TxtNormal.Text = ""
lAccion = 1
End Sub

Sub llena_grid()

     If Frmboleta.rsdesadic.RecordCount > 0 Then
        Frmboleta.rsdesadic.MoveFirst
        Do While Not Frmboleta.rsdesadic.EOF
           If Frmboleta.rsdesadic.Fields("codigo") = "07" Then
              Frmboleta.rsdesadic.Fields("monto") = Val(txtdescuento.Text)
           End If
           Frmboleta.rsdesadic.MoveNext
        Loop

     End If


End Sub
Private Function Verifica_Semana() As Boolean
Verifica_Semana = True
Txtsemana.Text = Format(Txtsemana.Text, "00")
Sql$ = "select * from plasemanas where cia='" & wcia & "' and ano='" & Format(dtfecha, "YYYY") & "' and semana='" & Format(Trim(Txtsemana.Text), "00") & "'  and status<>'*'"
Set RX = cn.Execute(Sql)
If RX.RecordCount <= 0 Then
   MsgBox "Semana Incorrecta", vbInformation
   Verifica_Semana = False
End If
If RX.State = adStateOpen Then RX.Close
End Function
