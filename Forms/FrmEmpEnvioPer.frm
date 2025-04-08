VERSION 5.00
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmEmpEnvioPer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7425
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   7425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdAceptar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7080
      Width           =   1695
   End
   Begin VB.CommandButton CmdCancelar 
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
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
      Height          =   375
      Index           =   1
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7080
      Width           =   1695
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   3375
      Index           =   2
      Left            =   120
      TabIndex        =   11
      Top             =   3600
      Width           =   7215
      Begin VB.CommandButton CmdNewEst 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Nuevo"
         Height          =   375
         Index           =   0
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2760
         Width           =   1695
      End
      Begin VB.CommandButton CmdUpdEst 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Modificar"
         Height          =   375
         Index           =   1
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2760
         Width           =   1695
      End
      Begin VB.CommandButton CmdDelEst 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Eliminar"
         Height          =   375
         Index           =   2
         Left            =   5400
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2760
         Width           =   1695
      End
      Begin TrueOleDBGrid70.TDBGrid DgrdEstablecimientos 
         Height          =   2415
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   4260
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Cod.Estab."
         Columns(0).DataField=   "codest"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Denominación"
         Columns(1).DataField=   "nomest"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   2
         Splits(0)._UserFlags=   0
         Splits(0).MarqueeStyle=   3
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   13160660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=2"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1561"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1482"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=10107"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=10028"
         Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(1).AutoDropDown=1"
         Splits(0)._ColumnProps(12)=   "Column(1).AutoCompletion=1"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowUpdate     =   0   'False
         DefColWidth     =   0
         HeadLines       =   2
         FootLines       =   1
         Caption         =   "Detalle de Establecimientos"
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
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37,.bgcolor=&H800000&"
         _StyleDefs(10)  =   ":id=4,.fgcolor=&H80000014&,.bold=-1,.fontsize=825,.italic=0,.underline=0"
         _StyleDefs(11)  =   ":id=4,.strikethrough=0,.charset=0"
         _StyleDefs(12)  =   ":id=4,.fontname=MS Sans Serif"
         _StyleDefs(13)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bgcolor=&HEFEFEF&,.appearance=0"
         _StyleDefs(14)  =   ":id=2,.bold=0,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(15)  =   ":id=2,.fontname=MS Sans Serif"
         _StyleDefs(16)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(17)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(18)  =   ":id=3,.fontname=MS Sans Serif"
         _StyleDefs(19)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(20)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(21)  =   "EditorStyle:id=7,.parent=1,.bgcolor=&HFFECD9&"
         _StyleDefs(22)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(23)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(24)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40,.bgcolor=&HFFEBD7&"
         _StyleDefs(25)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41,.bgcolor=&HEFEFEF&"
         _StyleDefs(26)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(27)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(28)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(29)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(30)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(31)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(32)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(33)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(34)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(35)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(36)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(37)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(38)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(39)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(40)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(41)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(42)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(43)  =   "Splits(0).Columns(1).Style:id=50,.parent=13"
         _StyleDefs(44)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=14"
         _StyleDefs(45)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=15"
         _StyleDefs(46)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=17"
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
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2295
      Index           =   1
      Left            =   120
      TabIndex        =   10
      Top             =   1320
      Width           =   7215
      Begin TrueOleDBGrid70.TDBGrid DgrdServ 
         Height          =   1935
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   3413
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   49
         Columns(0)._MaxComboItems=   20
         Columns(0).Caption=   "Actividad de la empresa"
         Columns(0).DataField=   "cod_serv"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   1
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   13160660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=1"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=11562"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=11483"
         Splits(0)._ColumnProps(4)=   "Column(0).Button=1"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(0).AutoDropDown=1"
         Splits(0)._ColumnProps(7)=   "Column(0).AutoCompletion=1"
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
         DefColWidth     =   0
         HeadLines       =   2
         FootLines       =   1
         Caption         =   "Servicios prestados a la Empresa"
         MultipleLines   =   0
         CellTipsWidth   =   0
         DeadAreaBackColor=   13160660
         RowDividerColor =   13160660
         RowSubDividerColor=   13160660
         DirectionAfterEnter=   2
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
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HFFFFFF&,.bold=0,.fontsize=825"
         _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37,.bgcolor=&H800000&"
         _StyleDefs(10)  =   ":id=4,.fgcolor=&H80000014&,.bold=-1,.fontsize=825,.italic=0,.underline=0"
         _StyleDefs(11)  =   ":id=4,.strikethrough=0,.charset=0"
         _StyleDefs(12)  =   ":id=4,.fontname=MS Sans Serif"
         _StyleDefs(13)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bgcolor=&HEFEFEF&,.appearance=0"
         _StyleDefs(14)  =   ":id=2,.bold=0,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(15)  =   ":id=2,.fontname=MS Sans Serif"
         _StyleDefs(16)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(17)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(18)  =   ":id=3,.fontname=MS Sans Serif"
         _StyleDefs(19)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(20)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(21)  =   "EditorStyle:id=7,.parent=1,.bgcolor=&HFFECD9&"
         _StyleDefs(22)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(23)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(24)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40,.bgcolor=&HFFEBD7&"
         _StyleDefs(25)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41,.bgcolor=&HEFEFEF&"
         _StyleDefs(26)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(27)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(28)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(29)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(30)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(31)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(32)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(33)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(34)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(35)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(36)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(37)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(38)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(39)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(40)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(41)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(42)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(43)  =   "Named:id=33:Normal"
         _StyleDefs(44)  =   ":id=33,.parent=0"
         _StyleDefs(45)  =   "Named:id=34:Heading"
         _StyleDefs(46)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(47)  =   ":id=34,.wraptext=-1"
         _StyleDefs(48)  =   "Named:id=35:Footing"
         _StyleDefs(49)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(50)  =   "Named:id=36:Selected"
         _StyleDefs(51)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(52)  =   "Named:id=37:Caption"
         _StyleDefs(53)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(54)  =   "Named:id=38:HighlightRow"
         _StyleDefs(55)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(56)  =   "Named:id=39:EvenRow"
         _StyleDefs(57)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(58)  =   "Named:id=40:OddRow"
         _StyleDefs(59)  =   ":id=40,.parent=33"
         _StyleDefs(60)  =   "Named:id=41:RecordSelector"
         _StyleDefs(61)  =   ":id=41,.parent=34"
         _StyleDefs(62)  =   "Named:id=42:FilterBar"
         _StyleDefs(63)  =   ":id=42,.parent=33"
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Declarante"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1215
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   7215
      Begin VB.TextBox TxtRuc 
         Height          =   315
         Left            =   1440
         MaxLength       =   11
         TabIndex        =   0
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox TxtRazsoc 
         Height          =   315
         Left            =   1440
         TabIndex        =   1
         Top             =   720
         Width           =   5535
      End
      Begin VB.Label Label1 
         Caption         =   "Ruc"
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
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Razon Social"
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
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   750
         Width           =   1215
      End
   End
End
Attribute VB_Name = "FrmEmpEnvioPer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public RsServ As New ADODB.Recordset
Public RsEstablecimientos As New ADODB.Recordset
Public RsEstTasa As New ADODB.Recordset
Dim vAccion As Integer
Dim Sql As String





Private Sub CmdAceptar_Click(Index As Integer)
If Trim(TxtRuc.Text) = "" Then
    MsgBox "Ingrese Nro de RUC", vbExclamation, Me.Caption
    TxtRuc.SetFocus
    Exit Sub
ElseIf Not IsNumeric(TxtRuc.Text) Then
    MsgBox "Ingrese Nro de RUC,correctamente solo dígitos", vbExclamation, Me.Caption
    TxtRuc.SetFocus
    Exit Sub
ElseIf Len(Trim(TxtRuc.Text)) < 11 Then
    MsgBox "Ingrese Nro de RUC correctamnete 11 dígitos", vbExclamation, Me.Caption
    TxtRuc.SetFocus
    Exit Sub
ElseIf Trim(Frmcia.TxtRuc.Text) = Trim(TxtRuc.Text) Then
    MsgBox "El Nro de Ruc no puede ser igual al de la compañia declarante", vbExclamation, Me.Caption
    TxtRuc.SetFocus
    Exit Sub
ElseIf Trim(TxtRazsoc.Text) = "" Then
    MsgBox "Ingrese Razón Social de la Empresa", vbExclamation, Me.Caption
    TxtRazsoc.SetFocus
    Exit Sub
ElseIf RsServ.RecordCount = 0 Then
    MsgBox "Ingrese Tipo de Actividad(es) de la empresa", vbExclamation, Me.Caption
    DgrdServ.SetFocus
    Exit Sub
'ElseIf RsEstablecimientos.RecordCount = 0 Then
'    MsgBox "Ingrese establecimientos de la empresa", vbExclamation, Me.Caption
'    DgrdEstablecimientos.SetFocus
'    Exit Sub
End If

With RsServ
    If .RecordCount > 0 Then
        .MoveFirst
        Do While Not .EOF
                If ExisteDuplicado(RsServ, "cod_serv", Trim(!cod_serv)) Then
                        MsgBox "La Actividad " & DgrdServ.Columns(0) & " ya existe", vbExclamation, Me.Caption
                   Exit Sub
                End If
            .MoveNext
        Loop
    End If
End With
Select Case vAccion
    Case 1 'nuevo
        If Frmcia.RsEmpEnvioPer.RecordCount > 0 Then
            Frmcia.RsEmpEnvioPer.MoveFirst
            Frmcia.RsEmpEnvioPer.FIND "ruc='" & Trim(TxtRuc.Text) & "'"
            If Not Frmcia.RsEmpEnvioPer.EOF Then
                MsgBox "El Número de Ruc ya fue ingresado", vbExclamation, Me.Caption
                TxtRuc.SetFocus
                Exit Sub
            End If
        End If
        Frmcia.RsEmpEnvioPer.AddNew
End Select
Frmcia.RsEmpEnvioPer!RUC = Trim(TxtRuc.Text)
Frmcia.RsEmpEnvioPer!razsoc = Trim(TxtRazsoc.Text)

'actividad y/o servicio
With RsServ
    
        If Frmcia.RsEmpEnvioPer_Actividad.RecordCount > 0 Then Frmcia.RsEmpEnvioPer_Actividad.MoveFirst
        Do While Not Frmcia.RsEmpEnvioPer_Actividad.EOF
            If Trim(Frmcia.RsEmpEnvioPer_Actividad!RUC) = Trim(Me.TxtRuc.Text) Then
                Frmcia.RsEmpEnvioPer_Actividad.Delete
            End If
           Frmcia.RsEmpEnvioPer_Actividad.MoveNext
        Loop
    If .RecordCount > 0 Then
        .MoveFirst
        Do While Not .EOF
                Frmcia.RsEmpEnvioPer_Actividad.AddNew
                Frmcia.RsEmpEnvioPer_Actividad!RUC = Trim(Me.TxtRuc.Text)
                Frmcia.RsEmpEnvioPer_Actividad!cod_serv = Trim(!cod_serv)
            .MoveNext
        Loop
    End If
End With
        
        
With RsEstablecimientos
    
        If Frmcia.RsEmpEnvioPer_Establecimientos.RecordCount > 0 Then Frmcia.RsEmpEnvioPer_Establecimientos.MoveFirst
        Do While Not Frmcia.RsEmpEnvioPer_Establecimientos.EOF
            If Trim(Frmcia.RsEmpEnvioPer_Establecimientos!RUC) = Trim(Me.TxtRuc.Text) Then
                Frmcia.RsEmpEnvioPer_Establecimientos.Delete
            End If
           Frmcia.RsEmpEnvioPer_Establecimientos.MoveNext
        Loop
    If .RecordCount > 0 Then
        .MoveFirst
        Do While Not .EOF
                Frmcia.RsEmpEnvioPer_Establecimientos.AddNew
                Frmcia.RsEmpEnvioPer_Establecimientos!RUC = Trim(TxtRuc.Text)
                Frmcia.RsEmpEnvioPer_Establecimientos!codest = Trim(!codest & "")
                Frmcia.RsEmpEnvioPer_Establecimientos!tipest = Trim(!tipest & "")
                Frmcia.RsEmpEnvioPer_Establecimientos!nomtipest = Trim(!tipest & "")
                Frmcia.RsEmpEnvioPer_Establecimientos!nomest = Trim(!nomest & "")
                Frmcia.RsEmpEnvioPer_Establecimientos!centro_riesgo = !centro_riesgo
                
            .MoveNext
        Loop
    End If
End With

With RsEstTasa
    
        If Frmcia.RsEmpEnvioPer_Establecimientos_Tasa.RecordCount > 0 Then Frmcia.RsEmpEnvioPer_Establecimientos_Tasa.MoveFirst
        Do While Not Frmcia.RsEmpEnvioPer_Establecimientos_Tasa.EOF
            If Trim(Frmcia.RsEmpEnvioPer_Establecimientos_Tasa!RUC) = Trim(Me.TxtRuc.Text) Then
                Frmcia.RsEmpEnvioPer_Establecimientos_Tasa.Delete
            End If
           Frmcia.RsEmpEnvioPer_Establecimientos_Tasa.MoveNext
        Loop
    If .RecordCount > 0 Then
        .MoveFirst
        Do While Not .EOF
                Frmcia.RsEmpEnvioPer_Establecimientos_Tasa.AddNew
                Frmcia.RsEmpEnvioPer_Establecimientos_Tasa!RUC = Trim(TxtRuc.Text)
                Frmcia.RsEmpEnvioPer_Establecimientos_Tasa!codest = Trim(!codest & "")
                Frmcia.RsEmpEnvioPer_Establecimientos_Tasa!porc = !porc
            .MoveNext
        Loop
    End If
End With



Select Case vAccion
    Case 1 'nuevo
        Frmcia.DgrdEnvPer.Refresh
        Limpiar
    Case 2 'modificar
        
        Unload Me
End Select

End Sub

Private Sub CmdCancelar_Click(Index As Integer)
Unload Me
End Sub

Private Sub CmdDelEst_Click(Index As Integer)
If Me.RsEstablecimientos.RecordCount = 0 Then Exit Sub
FrmEstab_EmpEnvioPer.MantAccion = 3 'eliminar
If MsgBox("Seguro de Eliminar el registro elegido? ", vbDefaultButton2 + vbYesNo + vbQuestion, "Eliminar") = vbNo Then Exit Sub
Dim xCod As String
xCod = Trim(RsEstablecimientos!codest)
If Me.RsEstablecimientos.RecordCount > 0 Then
    
    If Not RsEstablecimientos.BOF And Not RsEstablecimientos.EOF Then
        With RsEstTasa
            If RsEstTasa.RecordCount > 0 Then
                .MoveFirst
                Do While Not .EOF
                        If Trim(!codest) = Trim(xCod) Then
                            .Delete
                            .Update
                        End If
                    .MoveNext
                Loop
            End If
        End With
        RsEstablecimientos.Delete
        If Not RsEstablecimientos.BOF Then RsEstablecimientos.MovePrevious
    End If
End If

End Sub

Private Sub CmdNewEst_Click(Index As Integer)
FrmEstab_EmpEnvioPer.MantAccion = 1 'nuevo
FrmEstab_EmpEnvioPer.LblDecRuc(0).Caption = Trim(TxtRuc.Text)
FrmEstab_EmpEnvioPer.LblDecRazsoc(1).Caption = Trim(TxtRazsoc.Text)
FrmEstab_EmpEnvioPer.Show vbModal
End Sub

Private Sub CmdUpdEst_Click(Index As Integer)
If Me.RsEstablecimientos.RecordCount = 0 Then Exit Sub
    
FrmEstab_EmpEnvioPer.MantAccion = 2 'MODIFICAR
FrmEstab_EmpEnvioPer.LblDecRuc(0).Caption = Trim(TxtRuc.Text)
FrmEstab_EmpEnvioPer.LblDecRazsoc(1).Caption = Trim(TxtRazsoc.Text)

If RsEstablecimientos.RecordCount > 0 Then
    FrmEstab_EmpEnvioPer.TxtCodigo.Text = Trim(RsEstablecimientos!codest)
    FrmEstab_EmpEnvioPer.TxtDenominacion.Text = Trim(RsEstablecimientos!nomest)
    Call rUbiIndCmbBox(FrmEstab_EmpEnvioPer.CmbTipEst, RsEstablecimientos!tipest, "00")
    'FrmEstab_EmpEnvioPer.LblIndex.Caption = Me.RsEstablecimientos.AbsolutePosition
    If RsEstablecimientos!centro_riesgo = True Then
        FrmEstab_EmpEnvioPer.OptSi(0).Value = True
    Else
        FrmEstab_EmpEnvioPer.OptNo(1).Value = True
    End If
    If FrmEstab_EmpEnvioPer.RsTasa.RecordCount > 0 Then FrmEstab_EmpEnvioPer.RsTasa.MoveFirst
    Do While Not FrmEstab_EmpEnvioPer.RsTasa.EOF
       FrmEstab_EmpEnvioPer.RsTasa.Delete
       FrmEstab_EmpEnvioPer.RsTasa.MoveNext
    Loop
    With RsEstTasa
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                    If Trim(!codest) = Trim(RsEstablecimientos!codest) Then
                        FrmEstab_EmpEnvioPer.RsTasa.AddNew
                        FrmEstab_EmpEnvioPer.RsTasa!porc = RsEstTasa!porc
                    End If
                .MoveNext
            Loop
        End If
    End With
End If
FrmEstab_EmpEnvioPer.Show vbModal
End Sub

Private Sub DgrdEstablecimientos_DblClick()
If Me.RsEstablecimientos.RecordCount > 0 Then CmdUpdEst_Click (1)
End Sub

Private Sub Form_Load()
Crea_Rs

Sql = "select * from sunat_actividad_empresa order by nom_actividad"
TrueDbgrid_CargarCombo DgrdServ, 0, Sql, -1

Select Case vAccion
Case 1 'NUEVO
    Limpiar
    TxtRuc.BackColor = vbWhite
    TxtRuc.Enabled = True
    
Case 2 'MODIFICA
    TxtRuc.Enabled = False
    TxtRuc.BackColor = &HE0E0E0

End Select

End Sub



Public Property Get MantAccion() As Variant
    MantAccion = vAccion
End Property

Public Property Let MantAccion(ByVal vNewValue As Variant)
    vAccion = vNewValue
End Property

Private Sub Crea_Rs()
    
    'Servicios
    If RsServ.State = 1 Then RsServ.Close
    RsServ.Fields.Append "cod_serv", adVarChar, 50, adFldIsNullable
    RsServ.Open
    Set DgrdServ.DataSource = RsServ
    
   'Establecimientos
    If RsEstablecimientos.State = 1 Then RsEstablecimientos.Close
    RsEstablecimientos.Fields.Append "codest", adChar, 4, adFldIsNullable
    RsEstablecimientos.Fields.Append "tipest", adChar, 2, adFldIsNullable
    RsEstablecimientos.Fields.Append "nomtipest", adChar, 40, adFldIsNullable
    RsEstablecimientos.Fields.Append "nomest", adChar, 40, adFldIsNullable
    RsEstablecimientos.Fields.Append "centro_riesgo", adBoolean, , adFldIsNullable
    RsEstablecimientos.Open
    Set DgrdEstablecimientos.DataSource = RsEstablecimientos
       
    
    If RsEstTasa.State = 1 Then RsEstTasa.Close
    RsEstTasa.Fields.Append "codest", adChar, 4, adFldIsNullable
    RsEstTasa.Fields.Append "porc", adChar, 6, adFldIsNullable
    RsEstTasa.Open
    
End Sub

Public Sub Limpiar()
Me.TxtRuc.Text = ""
Me.TxtRazsoc.Text = ""

    If RsServ.RecordCount > 0 Then RsServ.MoveFirst
    Do While Not RsServ.EOF
       RsServ.Delete
       RsServ.MoveNext
    Loop
    
    If RsEstablecimientos.RecordCount > 0 Then RsEstablecimientos.MoveFirst
    Do While Not RsEstablecimientos.EOF
       RsEstablecimientos.Delete
       RsEstablecimientos.MoveNext
    Loop
    
End Sub

Public Sub TrueDbgrid_CargarCombo(ByRef Tdbgrd As TrueOleDBGrid70.TDBGrid, ByVal Col As Integer, ByVal Strsql As String, ByVal default As Integer)
    Dim vItem As New TrueOleDBGrid70.ValueItem
    Dim vItems As TrueOleDBGrid70.ValueItems
    Dim AdoRs As New ADODB.Recordset
    
    Set vItems = Tdbgrd.Columns(Col).ValueItems
    AdoRs.CursorLocation = adUseClient
    AdoRs.Open Strsql, cn, adOpenForwardOnly, adLockReadOnly
    Set AdoRs.ActiveConnection = Nothing
    
    If Not AdoRs.EOF Then
        Do While Not AdoRs.EOF
           vItem.Value = CStr(Trim(AdoRs(0)))
           vItem.DisplayValue = CStr(Trim(AdoRs(1)))
           vItems.Add vItem
           AdoRs.MoveNext
        Loop
        AdoRs.Close
    End If
    Set AdoRs = Nothing
    If default <> -1 Then vItems.DefaultItem = default
End Sub

Private Sub TxtRazsoc_GotFocus()
ResaltarTexto TxtRazsoc
End Sub

Private Sub TxtRazsoc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub TxtRazsoc_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TxtRuc_GotFocus()
ResaltarTexto TxtRuc
End Sub

Private Sub txtruc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub TxtRuc_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))

End Sub


Private Function ExisteDuplicado(ByVal pControl As ADODB.Recordset, ByVal pCampo As String, ByVal pDato As String) As Boolean
Dim Rc As ADODB.Recordset
Set Rc = pControl.Clone
Dim i As Integer
i = 0
With Rc
    If .RecordCount > 0 Then
        .MoveFirst
        Do While Not .EOF
            If Trim(UCase(.Fields(pCampo))) = Trim(UCase(pDato)) Then
                i = i + 1
            End If
            .MoveNext
        Loop
        .MoveFirst
    End If
End With
If i > 1 Then
    ExisteDuplicado = True
Else
    ExisteDuplicado = False
End If
Termina:
Rc.Close
Set Rc = Nothing

End Function

Public Function Busca_Razsoc(ByVal pRuc As String)
If Trim(TxtRuc.Text) <> "" Then
    Sql = "select top 1 razsoc from plaCiaEmpEnvioPer where cod_cia='" & wcia & "' and ruc='" & Trim(pRuc) & "' and status<>'*'"
    Dim Rq As ADODB.Recordset
    If fAbrRst(Rq, Sql) Then
        TxtRazsoc.Text = Trim(Rq!razsoc)
    End If
    Rq.Close
    Set Rq = Nothing
End If
End Function

Private Sub TxtRuc_LostFocus()
If Trim(TxtRuc.Text) <> "" Then Busca_Razsoc Trim(TxtRuc.Text)
End Sub
