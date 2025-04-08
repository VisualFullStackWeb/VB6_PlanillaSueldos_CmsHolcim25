VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{92F3EE28-0DF4-4CB0-AAA5-E6578B12929B}#1.0#0"; "CboFacil.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form Frmcia 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "» Mantenimiento de Compañia «"
   ClientHeight    =   7560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11805
   Icon            =   "Frmcia.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   11805
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   7515
      Left            =   60
      TabIndex        =   31
      Top             =   0
      Width           =   11700
      _ExtentX        =   20638
      _ExtentY        =   13256
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Datos Generales"
      TabPicture(0)   =   "Frmcia.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "SSFrame6"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "SSFrame7"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "ssfradir"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "ssfradat"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "ssfrarep"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "FrameTelf"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "E&stablecimientos de la compañia"
      TabPicture(1)   =   "Frmcia.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Empleadores a quienes destacó o desplazo personal"
      TabPicture(2)   =   "Frmcia.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame4"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Empleadores que me destacan o desplazan personal"
      TabPicture(3)   =   "Frmcia.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame5"
      Tab(3).ControlCount=   1
      Begin VB.Frame Frame5 
         Caption         =   "Detalle"
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
         Height          =   5895
         Left            =   -74880
         TabIndex        =   83
         Top             =   720
         Width           =   11175
         Begin VB.CommandButton CmdDelTraePer 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Eliminar"
            Height          =   375
            Index           =   1
            Left            =   7920
            Style           =   1  'Graphical
            TabIndex        =   86
            Top             =   5400
            Width           =   1695
         End
         Begin VB.CommandButton CmdUpdTraeper 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Modificar"
            Height          =   375
            Index           =   1
            Left            =   6120
            Style           =   1  'Graphical
            TabIndex        =   85
            Top             =   5400
            Width           =   1695
         End
         Begin VB.CommandButton CmdNewTraePer 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Nuevo"
            Height          =   375
            Index           =   0
            Left            =   4320
            Style           =   1  'Graphical
            TabIndex        =   84
            Top             =   5400
            Width           =   1695
         End
         Begin TrueOleDBGrid70.TDBGrid DgrdTraePer 
            Height          =   5055
            Left            =   120
            TabIndex        =   87
            Top             =   240
            Width           =   10935
            _ExtentX        =   19288
            _ExtentY        =   8916
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "RUC"
            Columns(0).DataField=   "ruc"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "RAZON SOCIAL"
            Columns(1).DataField=   "razsoc"
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
            Splits(0)._ColumnProps(1)=   "Column(0).Width=3254"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3175"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(6)=   "Column(1).Width=15399"
            Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=15319"
            Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            AllowUpdate     =   0   'False
            Appearance      =   2
            DefColWidth     =   0
            HeadLines       =   2
            FootLines       =   1
            Caption         =   "Empleadores que me destacan o desplazan personal"
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
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37,.bgcolor=&H808000&"
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
      Begin VB.Frame Frame4 
         Caption         =   "Detalle"
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
         Height          =   5895
         Left            =   -74760
         TabIndex        =   78
         Top             =   840
         Width           =   11175
         Begin VB.CommandButton CmdNewEnvPer 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Nuevo"
            Height          =   375
            Index           =   1
            Left            =   4320
            Style           =   1  'Graphical
            TabIndex        =   81
            Top             =   5400
            Width           =   1695
         End
         Begin VB.CommandButton CmdUpdEnvper 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Modificar"
            Height          =   375
            Index           =   0
            Left            =   6120
            Style           =   1  'Graphical
            TabIndex        =   80
            Top             =   5400
            Width           =   1695
         End
         Begin VB.CommandButton CmdDelEnvPer 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Eliminar"
            Height          =   375
            Index           =   0
            Left            =   7920
            Style           =   1  'Graphical
            TabIndex        =   79
            Top             =   5400
            Width           =   1695
         End
         Begin TrueOleDBGrid70.TDBGrid DgrdEnvPer 
            Height          =   5055
            Left            =   120
            TabIndex        =   82
            Top             =   255
            Width           =   10935
            _ExtentX        =   19288
            _ExtentY        =   8916
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "RUC"
            Columns(0).DataField=   "ruc"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "RAZON SOCIAL"
            Columns(1).DataField=   "razsoc"
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
            Splits(0)._ColumnProps(1)=   "Column(0).Width=3254"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3175"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(6)=   "Column(1).Width=15399"
            Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=15319"
            Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            AllowUpdate     =   0   'False
            Appearance      =   2
            DefColWidth     =   0
            HeadLines       =   2
            FootLines       =   1
            Caption         =   "Detalle Empleadores a quienes destacó o desplazo personal"
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
      Begin VB.Frame Frame1 
         Caption         =   "Establecimientos"
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
         Height          =   5895
         Left            =   -74760
         TabIndex        =   73
         Top             =   840
         Width           =   11055
         Begin VB.CommandButton CmdDelEst 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Eliminar"
            Height          =   375
            Index           =   2
            Left            =   7920
            Style           =   1  'Graphical
            TabIndex        =   77
            Top             =   5400
            Width           =   1695
         End
         Begin VB.CommandButton CmdUpdEst 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Modificar"
            Height          =   375
            Index           =   1
            Left            =   6120
            Style           =   1  'Graphical
            TabIndex        =   76
            Top             =   5400
            Width           =   1695
         End
         Begin VB.CommandButton CmdNewEst 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Nuevo"
            Height          =   375
            Index           =   0
            Left            =   4320
            Style           =   1  'Graphical
            TabIndex        =   75
            Top             =   5400
            Width           =   1695
         End
         Begin TrueOleDBGrid70.TDBGrid DgrdEstablecimientos 
            Height          =   5055
            Left            =   120
            TabIndex        =   74
            Top             =   240
            Width           =   10695
            _ExtentX        =   18865
            _ExtentY        =   8916
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
            Columns(1).Caption=   "Tipo de Establecimiento"
            Columns(1).DataField=   "nomtipest"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   49
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Denominación"
            Columns(2).DataField=   "nomest"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "codest"
            Columns(3).DataField=   "tipest"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   4
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   13160660
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=4"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2355"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2275"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(6)=   "Column(1).Width=8652"
            Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=8573"
            Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(11)=   "Column(2).Width=7223"
            Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=7144"
            Splits(0)._ColumnProps(14)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(15)=   "Column(2).Button=1"
            Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(17)=   "Column(2).AutoDropDown=1"
            Splits(0)._ColumnProps(18)=   "Column(2).AutoCompletion=1"
            Splits(0)._ColumnProps(19)=   "Column(3).Width=2725"
            Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=2646"
            Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(23)=   "Column(3).Order=4"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            AllowUpdate     =   0   'False
            Appearance      =   2
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
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37,.bgcolor=&HFF0000&"
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
            _StyleDefs(43)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
            _StyleDefs(44)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
            _StyleDefs(45)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(46)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(47)  =   "Splits(0).Columns(2).Style:id=50,.parent=13"
            _StyleDefs(48)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=14"
            _StyleDefs(49)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=15"
            _StyleDefs(50)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=17"
            _StyleDefs(51)  =   "Splits(0).Columns(3).Style:id=46,.parent=13"
            _StyleDefs(52)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
            _StyleDefs(53)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
            _StyleDefs(54)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
            _StyleDefs(55)  =   "Named:id=33:Normal"
            _StyleDefs(56)  =   ":id=33,.parent=0"
            _StyleDefs(57)  =   "Named:id=34:Heading"
            _StyleDefs(58)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(59)  =   ":id=34,.wraptext=-1"
            _StyleDefs(60)  =   "Named:id=35:Footing"
            _StyleDefs(61)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(62)  =   "Named:id=36:Selected"
            _StyleDefs(63)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(64)  =   "Named:id=37:Caption"
            _StyleDefs(65)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(66)  =   "Named:id=38:HighlightRow"
            _StyleDefs(67)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(68)  =   "Named:id=39:EvenRow"
            _StyleDefs(69)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(70)  =   "Named:id=40:OddRow"
            _StyleDefs(71)  =   ":id=40,.parent=33"
            _StyleDefs(72)  =   "Named:id=41:RecordSelector"
            _StyleDefs(73)  =   ":id=41,.parent=34"
            _StyleDefs(74)  =   "Named:id=42:FilterBar"
            _StyleDefs(75)  =   ":id=42,.parent=33"
         End
      End
      Begin VB.Frame FrameTelf 
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         Caption         =   "Frame19"
         Height          =   2655
         Left            =   3255
         TabIndex        =   61
         Top             =   2385
         Visible         =   0   'False
         Width           =   5295
         Begin Threed.SSCommand SSCommand5 
            Height          =   255
            Left            =   120
            TabIndex        =   62
            Top             =   2280
            Width           =   5055
            _Version        =   65536
            _ExtentX        =   8916
            _ExtentY        =   450
            _StockProps     =   78
            Caption         =   "Salir"
            ForeColor       =   -2147483635
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   1
         End
         Begin MSDataGridLib.DataGrid DgrdTelf 
            Height          =   1815
            Left            =   120
            TabIndex        =   63
            Top             =   480
            Width           =   5055
            _ExtentX        =   8916
            _ExtentY        =   3201
            _Version        =   393216
            AllowUpdate     =   -1  'True
            Appearance      =   0
            HeadLines       =   1
            RowHeight       =   15
            FormatLocked    =   -1  'True
            AllowAddNew     =   -1  'True
            AllowDelete     =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   3
            BeginProperty Column00 
               DataField       =   "telefono"
               Caption         =   "Telefono"
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
               DataField       =   "fax"
               Caption         =   "Fax"
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
            BeginProperty Column02 
               DataField       =   "referencia"
               Caption         =   "Referencia"
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
               EndProperty
               BeginProperty Column01 
               EndProperty
               BeginProperty Column02 
               EndProperty
            EndProperty
         End
         Begin VB.Label Label56 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "Lista de Teléfonos"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   375
            Left            =   120
            TabIndex        =   64
            Top             =   135
            Width           =   5055
         End
      End
      Begin VB.Frame Frame2 
         ForeColor       =   &H00800000&
         Height          =   1485
         Left            =   9405
         TabIndex        =   32
         Top             =   1845
         Width           =   1770
         Begin VB.OptionButton Opceps 
            Caption         =   "EPS"
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
            Left            =   75
            TabIndex        =   28
            Top             =   345
            Width           =   735
         End
         Begin VB.OptionButton Opcessalud 
            Caption         =   "ESSALUD"
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
            Left            =   90
            TabIndex        =   29
            Top             =   585
            Width           =   1095
         End
         Begin VB.TextBox Txtaccidente 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   180
            TabIndex        =   30
            Top             =   1125
            Width           =   1125
         End
         Begin MSForms.CheckBox Chkscrt 
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   0
            Width           =   855
            BackColor       =   -2147483633
            ForeColor       =   0
            DisplayStyle    =   4
            Size            =   "1508;450"
            Value           =   "0"
            Caption         =   "SCRT"
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Acc. de Trabajo"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   180
            TabIndex        =   33
            Top             =   900
            Width           =   1140
         End
      End
      Begin Threed.SSFrame ssfrarep 
         Height          =   1125
         Left            =   180
         TabIndex        =   34
         Top             =   5050
         Width           =   10995
         _Version        =   65536
         _ExtentX        =   19394
         _ExtentY        =   1984
         _StockProps     =   14
         Caption         =   "&Representante Legal"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShadowStyle     =   1
         Begin VB.TextBox TxtCargo_Rep_Legal 
            Height          =   330
            Left            =   8160
            MaxLength       =   50
            TabIndex        =   91
            Top             =   675
            Width           =   2655
         End
         Begin VB.ComboBox Cmbdia 
            Height          =   315
            ItemData        =   "Frmcia.frx":037A
            Left            =   9450
            List            =   "Frmcia.frx":0393
            Style           =   2  'Dropdown List
            TabIndex        =   89
            Top             =   240
            Visible         =   0   'False
            Width           =   1425
         End
         Begin CboFacil.cbo_facil cbonacion 
            Height          =   315
            Left            =   1320
            TabIndex        =   21
            Top             =   675
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   556
            NameTab         =   ""
            NameCod         =   ""
            NameDesc        =   ""
            Filtro          =   ""
            OrderBy         =   ""
            SetIndex        =   ""
            NameSistema     =   ""
            Mensaje         =   0   'False
            ToolTip         =   0   'False
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
         Begin VB.TextBox txtrlegal 
            Height          =   330
            Left            =   1320
            MaxLength       =   80
            TabIndex        =   20
            Top             =   270
            Width           =   6240
         End
         Begin VB.TextBox txtdoc 
            Height          =   330
            Left            =   6000
            MaxLength       =   11
            TabIndex        =   23
            Top             =   675
            Width           =   1455
         End
         Begin CboFacil.cbo_facil cbotipdoc 
            Height          =   315
            Left            =   3840
            TabIndex        =   22
            Top             =   675
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            NameTab         =   ""
            NameCod         =   ""
            NameDesc        =   ""
            Filtro          =   ""
            OrderBy         =   ""
            SetIndex        =   ""
            NameSistema     =   ""
            Mensaje         =   0   'False
            ToolTip         =   0   'False
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
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Cargo :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   7560
            TabIndex        =   92
            Top             =   720
            Width           =   540
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Inicio de Semana"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   7800
            TabIndex        =   90
            Top             =   300
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "T. Doc :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3240
            TabIndex        =   67
            Top             =   720
            Width           =   570
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Nombre R.Legal"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   135
            TabIndex        =   37
            Top             =   300
            Width           =   1140
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "N° Doc :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   5400
            TabIndex        =   36
            Top             =   720
            Width           =   600
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Nacionalidad"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   135
            TabIndex        =   35
            Top             =   720
            Width           =   900
         End
      End
      Begin Threed.SSFrame ssfradat 
         Height          =   2580
         Left            =   240
         TabIndex        =   38
         Top             =   700
         Width           =   9165
         _Version        =   65536
         _ExtentX        =   16166
         _ExtentY        =   4551
         _StockProps     =   14
         Caption         =   "&Datos Generales"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShadowStyle     =   1
         Begin VB.CheckBox Chk_EPS 
            Caption         =   "EPS"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5310
            TabIndex        =   88
            Top             =   2130
            Width           =   630
         End
         Begin CboFacil.cbo_facil cboestablecimiento 
            Height          =   315
            Left            =   1080
            TabIndex        =   4
            Top             =   1125
            Width           =   2310
            _ExtentX        =   4075
            _ExtentY        =   556
            NameTab         =   ""
            NameCod         =   ""
            NameDesc        =   ""
            Filtro          =   ""
            OrderBy         =   ""
            SetIndex        =   ""
            NameSistema     =   ""
            Mensaje         =   0   'False
            ToolTip         =   0   'False
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
         Begin VB.TextBox txtgiro 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4905
            MaxLength       =   60
            TabIndex        =   5
            Top             =   1125
            Width           =   3795
         End
         Begin VB.CheckBox Chksenati 
            Caption         =   "Senati"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4380
            TabIndex        =   10
            Top             =   2130
            Width           =   810
         End
         Begin VB.CommandButton cmdtelef 
            Height          =   450
            Left            =   8580
            Picture         =   "Frmcia.frx":03D3
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   1935
            Width           =   510
         End
         Begin VB.Frame Frame3 
            Caption         =   "Domicilio"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   90
            TabIndex        =   66
            Top             =   1935
            Width           =   4065
            Begin VB.OptionButton optnodomfiscal 
               Caption         =   "No Dom. Fiscal"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   2385
               TabIndex        =   9
               Top             =   255
               Value           =   -1  'True
               Width           =   1545
            End
            Begin VB.OptionButton optdomfiscal 
               Caption         =   "Dom. Fiscal"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   405
               TabIndex        =   8
               Top             =   255
               Width           =   1455
            End
         End
         Begin VB.TextBox txtruc 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   645
            MaxLength       =   11
            TabIndex        =   1
            Top             =   705
            Width           =   1545
         End
         Begin VB.TextBox txtcob 
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
            Height          =   315
            Left            =   7395
            MaxLength       =   6
            TabIndex        =   3
            Top             =   690
            Width           =   1665
         End
         Begin VB.TextBox txtmail 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   660
            MaxLength       =   60
            TabIndex        =   6
            Top             =   1545
            Width           =   4110
         End
         Begin VB.TextBox txtrazsoc 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   3585
            TabIndex        =   0
            Top             =   240
            Width           =   5475
         End
         Begin VB.TextBox txtreg 
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
            Height          =   315
            Left            =   4035
            MaxLength       =   10
            TabIndex        =   2
            Top             =   705
            Width           =   1305
         End
         Begin VB.TextBox txtcia 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00EBFEFC&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   945
            TabIndex        =   39
            Top             =   300
            Width           =   990
         End
         Begin CboFacil.cbo_facil cbocondicion 
            Height          =   315
            Left            =   6300
            TabIndex        =   7
            Top             =   1530
            Width           =   2805
            _ExtentX        =   4948
            _ExtentY        =   556
            NameTab         =   ""
            NameCod         =   ""
            NameDesc        =   ""
            Filtro          =   ""
            OrderBy         =   ""
            SetIndex        =   ""
            NameSistema     =   ""
            Mensaje         =   0   'False
            ToolTip         =   0   'False
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
         Begin VB.Label Label17 
            Appearance      =   0  'Flat
            BackColor       =   &H00EBFEFE&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " ..."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   8715
            TabIndex        =   72
            Top             =   1125
            Width           =   375
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Condicion :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   5310
            TabIndex        =   69
            Top             =   1575
            Width           =   795
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Teléfonos - Faxes"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   7080
            TabIndex        =   68
            Top             =   1935
            Width           =   1350
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Esta. :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   135
            TabIndex        =   65
            Top             =   1170
            Width           =   825
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Reg. Patronal :"
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
            Height          =   195
            Left            =   2760
            TabIndex        =   46
            Top             =   750
            Width           =   1095
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Cod. Cobertura :"
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
            Height          =   195
            Left            =   5940
            TabIndex        =   45
            Top             =   765
            Width           =   1230
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Giro de Negocio"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3555
            TabIndex        =   44
            Top             =   1170
            Width           =   1125
         End
         Begin VB.Label lblemail 
            AutoSize        =   -1  'True
            Caption         =   "E-Mail"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   90
            TabIndex        =   43
            Top             =   1620
            Width           =   420
         End
         Begin VB.Label Label49 
            AutoSize        =   -1  'True
            Caption         =   "Cod. Cia :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   135
            TabIndex        =   42
            Top             =   375
            Width           =   720
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Razón Social :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2325
            TabIndex        =   41
            Top             =   330
            Width           =   1005
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "RUC."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   135
            TabIndex        =   40
            Top             =   750
            Width           =   375
         End
      End
      Begin Threed.SSFrame ssfradir 
         Height          =   1755
         Left            =   180
         TabIndex        =   47
         Top             =   3300
         Width           =   10995
         _Version        =   65536
         _ExtentX        =   19394
         _ExtentY        =   3096
         _StockProps     =   14
         Caption         =   "D&atos de Dirección Legal"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShadowStyle     =   1
         Begin VB.TextBox txtzona 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4500
            MaxLength       =   20
            TabIndex        =   17
            Top             =   765
            Width           =   5145
         End
         Begin VB.CommandButton cmdubi 
            Appearance      =   0  'Flat
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Garamond"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   10455
            TabIndex        =   48
            Top             =   1320
            Width           =   345
         End
         Begin VB.TextBox txtdir 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3690
            MaxLength       =   60
            TabIndex        =   13
            Top             =   360
            Width           =   2805
         End
         Begin VB.TextBox txtref 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            MaxLength       =   60
            TabIndex        =   18
            Top             =   1335
            Width           =   5385
         End
         Begin VB.TextBox txtdpto 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   9360
            MaxLength       =   10
            TabIndex        =   15
            Top             =   315
            Width           =   825
         End
         Begin VB.TextBox txtnro 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   7155
            MaxLength       =   15
            TabIndex        =   14
            Top             =   360
            Width           =   960
         End
         Begin CboFacil.cbo_facil cbovia 
            Height          =   315
            Left            =   540
            TabIndex        =   12
            Top             =   360
            Width           =   2130
            _ExtentX        =   3757
            _ExtentY        =   556
            NameTab         =   "maestros_2"
            NameCod         =   "cod_maestro2"
            NameDesc        =   "descrip"
            Filtro          =   "ciamaestro='01068' and status!='*'"
            OrderBy         =   "descrip"
            SetIndex        =   ""
            NameSistema     =   "maestros_2"
            Mensaje         =   0   'False
            ToolTip         =   0   'False
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
         Begin CboFacil.cbo_facil cbozona 
            Height          =   315
            Left            =   1170
            TabIndex        =   16
            Top             =   810
            Width           =   2130
            _ExtentX        =   3757
            _ExtentY        =   556
            NameTab         =   ""
            NameCod         =   ""
            NameDesc        =   ""
            Filtro          =   ""
            OrderBy         =   ""
            SetIndex        =   ""
            NameSistema     =   ""
            Mensaje         =   0   'False
            ToolTip         =   0   'False
            Enabled         =   -1  'True
            Ninguno         =   0   'False
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
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Nom Zona :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3510
            TabIndex        =   71
            Top             =   810
            Width           =   825
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Via :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   135
            TabIndex        =   70
            Top             =   405
            Width           =   315
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "Lugar (País-Ciudad)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   5730
            TabIndex        =   54
            Top             =   1140
            Width           =   1410
         End
         Begin VB.Label lbllugar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
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
            Height          =   315
            Left            =   5700
            TabIndex        =   19
            Top             =   1350
            Width           =   4695
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            Caption         =   "Referencia"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   53
            Top             =   1125
            Width           =   780
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            Caption         =   "Int./Dpto."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   8415
            TabIndex        =   52
            Top             =   405
            Width           =   735
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Zona :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   135
            TabIndex        =   51
            Top             =   855
            Width           =   810
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            Caption         =   "N° :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   6810
            TabIndex        =   50
            Top             =   420
            Width           =   285
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            Caption         =   "Nom Via :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2835
            TabIndex        =   49
            Top             =   405
            Width           =   675
         End
      End
      Begin Threed.SSFrame SSFrame7 
         Height          =   1170
         Left            =   180
         TabIndex        =   55
         Top             =   6150
         Width           =   10995
         _Version        =   65536
         _ExtentX        =   19394
         _ExtentY        =   2064
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShadowStyle     =   1
         Begin MSForms.CheckBox ChkAportePatronal 
            Height          =   615
            Left            =   6240
            TabIndex        =   94
            Top             =   480
            Width           =   4695
            BackColor       =   -2147483633
            ForeColor       =   255
            DisplayStyle    =   4
            Size            =   "8281;1085"
            Value           =   "0"
            Caption         =   "Empresa Requiere Aporte Patronal en Provisión de Vacaciones"
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   195
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
         Begin MSForms.CheckBox ChkTrabSunat 
            Height          =   375
            Left            =   6240
            TabIndex        =   93
            Top             =   150
            Width           =   4695
            BackColor       =   -2147483633
            ForeColor       =   255
            DisplayStyle    =   4
            Size            =   "8281;661"
            Value           =   "0"
            Caption         =   "Empresa Requiere Tipo de Trabajador SUNAT"
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   195
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
         Begin VB.Label Label47 
            AutoSize        =   -1  'True
            Caption         =   "Fec. Ing :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   59
            Top             =   225
            Width           =   705
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Fec Ult. Modi :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3000
            TabIndex        =   58
            Top             =   255
            Width           =   1035
         End
         Begin VB.Label lblfecmodi 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
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
            Height          =   315
            Left            =   4080
            TabIndex        =   57
            Top             =   180
            Width           =   1965
         End
         Begin VB.Label lblfecing 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
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
            Height          =   315
            Left            =   840
            TabIndex        =   56
            Top             =   180
            Width           =   2010
         End
      End
      Begin Threed.SSFrame SSFrame6 
         Height          =   1095
         Left            =   9435
         TabIndex        =   60
         Top             =   735
         Width           =   1740
         _Version        =   65536
         _ExtentX        =   3069
         _ExtentY        =   1931
         _StockProps     =   14
         Caption         =   "&Situación"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShadowStyle     =   1
         Begin VB.OptionButton opcinactivo 
            Caption         =   "Inactivo"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   270
            TabIndex        =   25
            Top             =   540
            Width           =   885
         End
         Begin VB.OptionButton opcactivo 
            Caption         =   "Activo"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   270
            TabIndex        =   24
            Top             =   270
            Value           =   -1  'True
            Width           =   915
         End
         Begin VB.OptionButton opcanula 
            Caption         =   "Anulado"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   270
            TabIndex        =   26
            Top             =   810
            Width           =   1125
         End
      End
   End
End
Attribute VB_Name = "Frmcia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim VCia As String
Dim VNgiro As String
Dim VNUrb As String
Public Ciacodubi$, cod_rlegal$
Dim rstlf As New Recordset
Dim Sql As String
Public RsEstablecimientos As New ADODB.Recordset
Public RsEstTasa As New ADODB.Recordset

Public RsEmpEnvioPer As New ADODB.Recordset
Public RsEmpEnvioPer_Actividad As New ADODB.Recordset
Public RsEmpEnvioPer_Establecimientos As New ADODB.Recordset
Public RsEmpEnvioPer_Establecimientos_Tasa As New ADODB.Recordset

Public RsEmpTraePer As New ADODB.Recordset
Public RsEmpTraePer_Actividad As New ADODB.Recordset

Private Sub Chkscrt_Click()
If Chkscrt.Value = False Then
   Opceps.Value = False
   Opcessalud.Value = False
End If
End Sub

Private Sub CmdDelEnvPer_Click(index As Integer)
If Me.RsEmpEnvioPer.RecordCount = 0 Then Exit Sub
FrmEmpEnvioPer.MantAccion = 3 'eliminar
If MsgBox("Seguro de Eliminar el registro elegido? ", vbDefaultButton2 + vbYesNo + vbQuestion, "Eliminar") = vbNo Then Exit Sub
If Me.RsEmpEnvioPer.RecordCount > 0 Then
    Dim xCod As String
    xCod = Trim(RsEmpEnvioPer!RUC)
    If Not RsEmpEnvioPer.BOF And Not RsEmpEnvioPer.EOF Then

        Eliminar_Reg RsEmpEnvioPer_Actividad, xCod
        Eliminar_Reg RsEmpEnvioPer_Establecimientos, xCod
        Eliminar_Reg RsEmpEnvioPer_Establecimientos_Tasa, xCod
        RsEmpEnvioPer.Delete
        If Not RsEmpEnvioPer.BOF Then RsEmpEnvioPer.MovePrevious
    End If
End If
End Sub

Private Sub CmdDelEst_Click(index As Integer)
If Me.RsEstablecimientos.RecordCount = 0 Then Exit Sub
FrmEstab.MantAccion = 3 'eliminar
If MsgBox("Seguro de Eliminar el registro elegido? ", vbDefaultButton2 + vbYesNo + vbQuestion, "Eliminar") = vbNo Then Exit Sub
If Me.RsEstablecimientos.RecordCount > 0 Then
    Dim xCod As String
    xCod = Trim(RsEstablecimientos!codest)
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

Private Sub CmdDelTraePer_Click(index As Integer)
If Me.RsEmpTraePer.RecordCount = 0 Then Exit Sub
FrmEmpTraePer.MantAccion = 3 'eliminar
If MsgBox("Seguro de Eliminar el registro elegido? ", vbDefaultButton2 + vbYesNo + vbQuestion, "Eliminar") = vbNo Then Exit Sub
If Me.RsEmpTraePer.RecordCount > 0 Then
    Dim xCod As String
    xCod = Trim(RsEmpTraePer!RUC)
    If Not RsEmpTraePer.BOF And Not RsEmpTraePer.EOF Then
        Eliminar_Reg RsEmpTraePer_Actividad, xCod
        RsEmpTraePer.Delete
        If Not RsEmpTraePer.BOF Then RsEmpTraePer.MovePrevious
    End If
End If
End Sub

Private Sub CmdNewEnvPer_Click(index As Integer)
FrmEmpEnvioPer.MantAccion = 1 'nuevo
FrmEmpEnvioPer.Caption = Trim(txtruc.Text) & " - " & Trim(txtrazsoc.Text)
FrmEmpEnvioPer.Show vbModal
End Sub

Private Sub CmdNewEst_Click(index As Integer)
FrmEstab.MantAccion = 1 'nuevo
FrmEstab.LblDecRuc(0).Caption = Trim(txtruc.Text)
FrmEstab.LblDecRazsoc(1).Caption = Trim(txtrazsoc.Text)
FrmEstab.Show vbModal

End Sub

Private Sub CmdNewTraePer_Click(index As Integer)
FrmEmpTraePer.MantAccion = 1 'nuevo
FrmEmpTraePer.Caption = Trim(txtruc.Text) & " - " & Trim(txtrazsoc.Text)
FrmEmpTraePer.Show vbModal
End Sub

Private Sub cmdtelef_Click()
FrameTelf.Visible = True
FrameTelf.ZOrder 0
cmdtelef.Enabled = False
End Sub

Private Sub cmdubi_Click()
'Load FrmUbigeo
'FrmUbigeo.Show
'FrmUbigeo.ZOrder 0

Load FrmUbiSunat
FrmUbiSunat.Show
FrmUbiSunat.ZOrder 0

End Sub

Private Sub CmdUpdEnvper_Click(index As Integer)
If Me.RsEmpEnvioPer.RecordCount = 0 Then Exit Sub
    
FrmEmpEnvioPer.MantAccion = 2 'MODIFICAR
FrmEmpEnvioPer.Caption = Trim(txtruc.Text) & " - " & Trim(txtrazsoc.Text)

If RsEmpEnvioPer.RecordCount > 0 Then
    FrmEmpEnvioPer.txtruc.Text = Trim(RsEmpEnvioPer!RUC)
    FrmEmpEnvioPer.txtrazsoc.Text = Trim(RsEmpEnvioPer!razsoc)

    If FrmEmpEnvioPer.RsServ.RecordCount > 0 Then FrmEmpEnvioPer.RsServ.MoveFirst
    Do While Not FrmEmpEnvioPer.RsServ.EOF
       FrmEmpEnvioPer.RsServ.Delete
       FrmEmpEnvioPer.RsServ.MoveNext
    Loop
    
    With RsEmpEnvioPer_Actividad
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                    If Trim(!RUC) = Trim(RsEmpEnvioPer!RUC) Then
                        FrmEmpEnvioPer.RsServ.AddNew
                        FrmEmpEnvioPer.RsServ!cod_serv = Trim(!cod_serv)
                    End If
                .MoveNext
            Loop
        End If
    End With
End If

If Me.RsEmpEnvioPer_Establecimientos.RecordCount > 0 Then
    If FrmEmpEnvioPer.RsEstablecimientos.RecordCount > 0 Then FrmEmpEnvioPer.RsEstablecimientos.MoveFirst
    Do While Not FrmEmpEnvioPer.RsEstablecimientos.EOF
       FrmEmpEnvioPer.RsEstablecimientos.Delete
       FrmEmpEnvioPer.RsEstablecimientos.MoveNext
    Loop
    With Me.RsEmpEnvioPer_Establecimientos
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                    If Trim(!RUC) = Trim(RsEmpEnvioPer!RUC) Then
                        FrmEmpEnvioPer.RsEstablecimientos.AddNew
                        FrmEmpEnvioPer.RsEstablecimientos!codest = Trim(!codest)
                        FrmEmpEnvioPer.RsEstablecimientos!tipest = Trim(!tipest)
                        FrmEmpEnvioPer.RsEstablecimientos!nomtipest = Trim(!nomtipest)
                        FrmEmpEnvioPer.RsEstablecimientos!nomest = Trim(!nomest)
                        FrmEmpEnvioPer.RsEstablecimientos!centro_riesgo = !centro_riesgo
                        
                    End If
                .MoveNext
            Loop
        End If
    End With
End If


If RsEmpEnvioPer_Establecimientos_Tasa.RecordCount > 0 Then
    If FrmEmpEnvioPer.RsEstTasa.RecordCount > 0 Then FrmEmpEnvioPer.RsEstTasa.MoveFirst
    Do While Not FrmEmpEnvioPer.RsEstTasa.EOF
       FrmEmpEnvioPer.RsEstTasa.Delete
       FrmEmpEnvioPer.RsEstTasa.MoveNext
    Loop
    With Me.RsEmpEnvioPer_Establecimientos_Tasa
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                    If Trim(!RUC) = Trim(RsEmpEnvioPer!RUC) Then
                        FrmEmpEnvioPer.RsEstTasa.AddNew
                        FrmEmpEnvioPer.RsEstTasa!codest = Trim(!codest)
                        FrmEmpEnvioPer.RsEstTasa!PORC = Trim(!PORC)
                        
                    End If
                .MoveNext
            Loop
        End If
    End With
End If

FrmEmpEnvioPer.Show vbModal
End Sub

Private Sub CmdUpdEst_Click(index As Integer)
If Me.RsEstablecimientos.RecordCount = 0 Then Exit Sub
    
FrmEstab.MantAccion = 2 'MODIFICAR
FrmEstab.LblDecRuc(0).Caption = Trim(txtruc.Text)
FrmEstab.LblDecRazsoc(1).Caption = Trim(txtrazsoc.Text)

If RsEstablecimientos.RecordCount > 0 Then
    FrmEstab.TxtCodigo.Text = Trim(RsEstablecimientos!codest)
    FrmEstab.TxtDenominacion.Text = Trim(RsEstablecimientos!nomest)
    Call rUbiIndCmbBox(FrmEstab.CmbTipEst, RsEstablecimientos!tipest, "00")
    'FrmEstab.LblIndex.Caption = Me.RsEstablecimientos.AbsolutePosition
    If RsEstablecimientos!centro_riesgo = True Then
        FrmEstab.OptSi(0).Value = True
    Else
        FrmEstab.OptNo(1).Value = True
    End If
    If FrmEstab.RsTasa.RecordCount > 0 Then FrmEstab.RsTasa.MoveFirst
    Do While Not FrmEstab.RsTasa.EOF
       FrmEstab.RsTasa.Delete
       FrmEstab.RsTasa.MoveNext
    Loop
    With RsEstTasa
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                    If Trim(!codest) = Trim(RsEstablecimientos!codest) Then
                        FrmEstab.RsTasa.AddNew
                        FrmEstab.RsTasa!PORC = RsEstTasa!PORC
                    End If
                .MoveNext
            Loop
        End If
    End With
End If
FrmEstab.Show vbModal

End Sub

Private Sub CmdUpdTraeper_Click(index As Integer)
If Me.RsEmpEnvioPer.RecordCount = 0 Then Exit Sub
    
FrmEmpTraePer.MantAccion = 2 'MODIFICAR
FrmEmpTraePer.Caption = Trim(txtruc.Text) & " - " & Trim(txtrazsoc.Text)

If RsEmpTraePer.RecordCount > 0 Then
    FrmEmpTraePer.txtruc.Text = Trim(RsEmpTraePer!RUC)
    FrmEmpTraePer.txtrazsoc.Text = Trim(RsEmpTraePer!razsoc)

    If FrmEmpTraePer.RsServ.RecordCount > 0 Then FrmEmpTraePer.RsServ.MoveFirst
    Do While Not FrmEmpTraePer.RsServ.EOF
       FrmEmpTraePer.RsServ.Delete
       FrmEmpTraePer.RsServ.MoveNext
    Loop
    
    
    With RsEmpTraePer_Actividad
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                    If Trim(!RUC) = Trim(RsEmpTraePer!RUC) Then
                        FrmEmpTraePer.RsServ.AddNew
                        FrmEmpTraePer.RsServ!cod_serv = Trim(!cod_serv)
                    End If
                .MoveNext
            Loop
        End If
    End With
End If

'If Me.RsEmpEnvioPer_Establecimientos.RecordCount > 0 Then
'    If FrmEmpEnvioPer.RsEstablecimientos.RecordCount > 0 Then FrmEmpEnvioPer.RsEstablecimientos.MoveFirst
'    Do While Not FrmEmpEnvioPer.RsEstablecimientos.EOF
'       FrmEmpEnvioPer.RsEstablecimientos.Delete
'       FrmEmpEnvioPer.RsEstablecimientos.MoveNext
'    Loop
'    With Me.RsEmpEnvioPer_Establecimientos
'        If .RecordCount > 0 Then
'            .MoveFirst
'            Do While Not .EOF
'                    If Trim(!RUC) = Trim(RsEmpEnvioPer!RUC) Then
'                        FrmEmpEnvioPer.RsEstablecimientos.AddNew
'                        FrmEmpEnvioPer.RsEstablecimientos!codest = Trim(!codest)
'                        FrmEmpEnvioPer.RsEstablecimientos!tipest = Trim(!tipest)
'                        FrmEmpEnvioPer.RsEstablecimientos!nomtipest = Trim(!nomtipest)
'                        FrmEmpEnvioPer.RsEstablecimientos!nomest = Trim(!nomest)
'                        FrmEmpEnvioPer.RsEstablecimientos!centro_riesgo = !centro_riesgo
'
'                    End If
'                .MoveNext
'            Loop
'        End If
'    End With
'End If


'If RsEmpEnvioPer_Establecimientos_Tasa.RecordCount > 0 Then
'    If FrmEmpEnvioPer.RsEstTasa.RecordCount > 0 Then FrmEmpEnvioPer.RsEstTasa.MoveFirst
'    Do While Not FrmEmpEnvioPer.RsEstTasa.EOF
'       FrmEmpEnvioPer.RsEstTasa.Delete
'       FrmEmpEnvioPer.RsEstTasa.MoveNext
'    Loop
'    With Me.RsEmpEnvioPer_Establecimientos_Tasa
'        If .RecordCount > 0 Then
'            .MoveFirst
'            Do While Not .EOF
'                    If Trim(!RUC) = Trim(RsEmpEnvioPer!RUC) Then
'                        FrmEmpEnvioPer.RsEstTasa.AddNew
'                        FrmEmpEnvioPer.RsEstTasa!codest = Trim(!codest)
'                        FrmEmpEnvioPer.RsEstTasa!porc = Trim(!porc)
'
'                    End If
'                .MoveNext
'            Loop
'        End If
'    End With
'End If

FrmEmpTraePer.Show vbModal

End Sub

Private Sub DgrdEnvPer_DblClick()
If Me.RsEstablecimientos.RecordCount > 0 Then CmdUpdEnvper_Click (1)
End Sub

Private Sub DgrdEstablecimientos_DblClick()
If Me.RsEstablecimientos.RecordCount > 0 Then CmdUpdEst_Click (1)

End Sub

Private Sub DgrdTraePer_DblClick()
If Me.RsEstablecimientos.RecordCount > 0 Then CmdUpdTraeper_Click (1)
End Sub

Private Sub Label17_Click()
frmgironegocio.OpcionCarga = GIROS_NEGOCIO
frmgironegocio.Show vbModal
End Sub

Private Sub Form_Load()
'Me.Top = 0
'Me.Left = 0
'Me.Width = 11175
'Me.Height = 6915

Me.Top = 0
Me.Left = 0
Me.Height = 8040
Me.Width = 11895

Crea_Rs

'With cnsgiro
'    .BackColorBoton = &HEFF7F7
'    .BackColorClick = &HFF8080
'    .BackColorAntesDelClick = &HFFC0C0
'    .Visible = True
'End With

'TIPOS ESTABLECIMIENTOS
With cboestablecimiento
    .NameCod = "cod_maestro2"
    .NameDesc = "descrip"
    .NameTab = "maestros_2"
    .Filtro = " right(ciamaestro,3)='138' and status!='*'"
    .conexion = cn
    .Execute
End With
'****************************
'TIPOS DE VIAS
With cbovia
    .NameCod = "cod_maestro2"
    .NameDesc = "descrip"
    .NameTab = "maestros_2"
    .Filtro = " right(ciamaestro,3)='036' and flag6 IS NOT NULL AND status!='*'"
    .conexion = cn
    .Execute
End With
'****************************
'NACIONES
With cbonacion
    .NameCod = "codigo"
    .NameDesc = "descripcion"
    .NameTab = "tnacionalidad"
    .conexion = cn
    .Execute
End With
'****************************
'TIPOS DE DOCUMENTOS
Sql = "select cod_maestro2, descrip from maestros_2 where ciamaestro='01032' and status <> '*'"
CargaCombo2 cbotipdoc, , Sql, False, True

'With cbotipdoc
'    .NameCod = "cod_maestro2"
'    .NameDesc = "descrip"
'    .NameTab = "maestros_2"
'    .Filtro = " ciamaestro='01032' AND flag6 IS NOT NULL  AND status!='*'"
'    .Conexion = cn
'    .Execute
'End With
'****************************
'CONDICION DE ESTABLECIMIENTO
With cbocondicion
    .NameCod = "cod_maestro2"
    .NameDesc = "descrip"
    .NameTab = "maestros_2"
    .Filtro = " right(ciamaestro,3)='139' and status!='*'"
    .conexion = cn
    .Execute
End With
'******************************
'TIPOS ZONAS
With cbozona
    .NameCod = "cod_maestro2"
    .NameDesc = "descrip"
    .NameTab = "maestros_2"
    .Filtro = " right(ciamaestro,3)='035' and status!='*'"
    .conexion = cn
    .Execute
End With

'Call fc_Descrip_Maestros2("01010", "", cmbgiro)
cn.CursorLocation = adUseClient
Set rs = New ADODB.Recordset
Sql$ = "select * from tablas where cod_tipo='01'  and status<>'*' order by descrip desc"
Set rs = cn.Execute(Sql$)
'Cmburb.Clear
'Do While Not rs.EOF
'   Cmburb.AddItem rs!descrip
'   Cmburb.ItemData(Cmburb.NewIndex) = Trim(rs!cod_maestro2)
'   rs.MoveNext
'Loop
'Telefonos
If rstlf.State = 1 Then rstlf.Close
rstlf.Fields.Append "telefono", adChar, 15, adFldIsNullable
rstlf.Fields.Append "fax", adChar, 15, adFldIsNullable
rstlf.Fields.Append "referencia", adChar, 25, adFldIsNullable
rstlf.Open
Set DgrdTelf.DataSource = rstlf
SSTab1.Tab = 0
End Sub
Public Sub Carga_Cia(CODCIA As String, razon As String)
wcia = CODCIA
VCia = CODCIA
txtcia.Text = CODCIA

Sql$ = " select c.*, " & _
      " (select nombre from sunat_departamento where id_dpto=left(c.cod_ubi,2)) as dp, " & _
      " (select nombre from sunat_provincia where id_prov=left(c.cod_ubi,4)) as prov, " & _
      "  u.nombre  as dist, 'PERU' as pais from cia c, sunat_ubigeo u " & _
      "  WHERE cod_cia='" & CODCIA & "' AND u.id_ubigeo=c.cod_ubi"

cn.CursorLocation = adUseClient
Set rs = New ADODB.Recordset
Set rs = cn.Execute(Sql$, 64)
'rs.MoveFirst
If IsNull(rs!acctrab) Then Txtaccidente.Text = "0.00" Else Txtaccidente.Text = Format(rs!acctrab, "###,###.00")
If rs!sctreps = "S" Then
   Opceps.Value = True
   Chkscrt.Value = True
End If
If rs!sctressalud = "S" Then
   Opcessalud.Value = True
   Chkscrt.Value = True
End If
If rs!senati = "S" Then Chksenati.Value = 1
'========================================================================
Chk_EPS.Value = IIf(IsNull(rs!Aportacion_EPS), 0, rs!Aportacion_EPS)
'========================================================================
If rs!status = "*" Then
  opcanula.Value = True
ElseIf rs!status = "I" Then
   opcinactivo.Value = True
Else: opcactivo.Value = True
End If
           
'cbosemana.ListIndex = rs!semana_pago - 1
           
If Not IsNull(rs!condicion_esta) Then cbocondicion.SetIndice rs!condicion_esta
If Not IsNull(rs!tipo_esta) Then cboestablecimiento.SetIndice rs!tipo_esta
If Not IsNull(rs!reP_nac) Then cbonacion.SetIndice rs!reP_nac
If Not IsNull(rs!tipo_doc) Then cbotipdoc.SetIndice rs!tipo_doc
If Not IsNull(rs!cod_via) Then cbovia.SetIndice rs!cod_via
If Not IsNull(rs!cod_zona) Then cbozona.SetIndice rs!cod_zona
If Trim(rs!TipoTrabSunat & "") = "S" Then ChkTrabSunat.Value = True Else ChkTrabSunat.Value = False
If Trim(rs!AportePatronal & "") = "S" Then ChkAportePatronal.Value = True Else ChkAportePatronal.Value = False
txtcia = rs!cod_cia
txtrazsoc = rs!razsoc
txtruc = rs!RUC
txtcob = IIf(IsNull(rs!cod_cober), "", rs!cod_cober)
txtreg = IIf(IsNull(rs!reg_pat), "", rs!reg_pat)
txtdir = rs!direcc
txtdpto = rs!DPTO
txtnro = rs!NRO
urb$ = rs!urb
lbllugar = rs!pais & " - " & rs!DP & " - " & rs!PROV & " - " & rs!DIST
txtref = rs!ref
            
txtmsg = Trim(rs!ctemsg)
'cmbgiro.ListIndex = RS!giro_emp - 1
'Call rUbiIndCmbBox(cmbgiro, rs!giro_emp, "00")
'Call rUbiIndCmbBox(Cmburb, rs!urb, "000")
lbllugar.Tag = rs!cod_ubi
txtmail = rs!Email
cod_rlegal = rs!rep_legal
txtrlegal.Text = rs!rep_nom
TxtCargo_Rep_Legal.Text = Trim(rs!cargo_rep_legal & "")
'lblrnac = rs!rep_nac
txtdoc.Text = IIf(IsNull(rs!num_doc), "", Trim(rs!num_doc))
'txtrle = rs!rep_le
'txtrruc = rs!rep_ruc
'txtrpas = rs!rep_pasap
lblfecing = rs!fec_crea
If Not IsNull(rs!fec_modi) Then lblfecmodi = rs!fec_modi

If rs.State = 1 Then rs.Close

'telefonos
Sql$ = "SELECT telef,fax,referencia FROM telef_cia WHERE cod_cia='" & Trim(wcia) & "'"
cn.CursorLocation = adUseClient
Set rs = New ADODB.Recordset
Set rs = cn.Execute(Sql$, 64)
If (fAbrRst(rs, Sql$)) Then
   rs.MoveFirst
   Do While Not rs.EOF
      rstlf.AddNew
      rstlf!telefono = rs!telef
      rstlf!fax = rs!fax
      rstlf!referencia = rs!referencia
      rs.MoveNext
  Loop
End If
If rs.State = 1 Then rs.Close

If RsEstablecimientos.RecordCount > 0 Then RsEstablecimientos.MoveFirst
Do While Not RsEstablecimientos.EOF
   RsEstablecimientos.Delete
   RsEstablecimientos.MoveNext
Loop
If RsEstTasa.RecordCount > 0 Then RsEstTasa.MoveFirst
Do While Not RsEstTasa.EOF
   RsEstTasa.Delete
   RsEstTasa.MoveNext
Loop

Sql = "select * "
Sql = Sql & " ,(select descrip from maestros_2 where right(ciamaestro,3)='138' and cod_maestro2=placiaestablecimientos.tipo_establecimiento and status<>'*') as nomtip_establecimiento"
Sql = Sql & " from placiaestablecimientos where cod_cia='" & wcia & "' and status<>'*' order by cod_establecimiento"
Dim Re As ADODB.Recordset
If fAbrRst(Re, Sql) Then
    Do While Not Re.EOF
            RsEstablecimientos.AddNew
            RsEstablecimientos!codest = Trim(Re!cod_establecimiento & "")
            RsEstablecimientos!tipest = Trim(Re!tipo_establecimiento & "")
            RsEstablecimientos!nomtipest = Trim(Re!nomtip_establecimiento & "")
            RsEstablecimientos!nomest = Trim(Re!nom_establecimiento & "")
            RsEstablecimientos!centro_riesgo = IIf(Re!indicador_centro_riesgo = True, True, False)
            
        Re.MoveNext
    Loop
End If

Sql = "select * "
Sql = Sql & " from placiaestablecimientostasa where cod_cia='" & wcia & "' and status<>'*' "
If fAbrRst(Re, Sql) Then
    Do While Not Re.EOF
            RsEstTasa.AddNew
            RsEstTasa!codest = Trim(Re!cod_establecimiento & "")
            RsEstTasa!PORC = Format(Re!Tasa & "", "##0.00")
            
        Re.MoveNext
    Loop
End If


'//**** Empleadores a quienes destacó o desplazo personal ****/////

EliminarRst RsEmpEnvioPer
EliminarRst RsEmpEnvioPer_Actividad

'If RsEmpEnvioPer.RecordCount > 0 Then RsEmpEnvioPer.MoveFirst
'Do While Not RsEmpEnvioPer.EOF
'   RsEmpEnvioPer.Delete
'   RsEmpEnvioPer.MoveNext
'Loop
'
'If RsEmpEnvioPer_Actividad.RecordCount > 0 Then RsEmpEnvioPer_Actividad.MoveFirst
'Do While Not RsEmpEnvioPer_Actividad.EOF
'   RsEmpEnvioPer_Actividad.Delete
'   RsEmpEnvioPer_Actividad.MoveNext
'Loop

Sql = "select * "
Sql = Sql & " from plaCiaEmpEnvioPer where cod_cia='" & wcia & "' and status<>'*' "
If fAbrRst(Re, Sql) Then
    Do While Not Re.EOF
            RsEmpEnvioPer.AddNew
            RsEmpEnvioPer!RUC = Trim(Re!RUC & "")
            RsEmpEnvioPer!razsoc = Trim(Re!razsoc & "")
        Re.MoveNext
    Loop
End If

Sql = "select * "
Sql = Sql & " from plaCiaEmpEnvioPer_Actividad where cod_cia='" & wcia & "' and status<>'*' "
If fAbrRst(Re, Sql) Then
    Do While Not Re.EOF
            RsEmpEnvioPer_Actividad.AddNew
            RsEmpEnvioPer_Actividad!RUC = Trim(Re!RUC & "")
            RsEmpEnvioPer_Actividad!cod_serv = Trim(Re!cod_serv & "")
        Re.MoveNext
    Loop
End If

EliminarRst RsEmpEnvioPer_Establecimientos

Sql = "select * "
Sql = Sql & " ,(select descrip from maestros_2 where right(ciamaestro,3)='138' and cod_maestro2=plaCiaEmpEnvioPer_Establecimiento.tipo_establecimiento and status<>'*') as nomtip_establecimiento"
Sql = Sql & " from plaCiaEmpEnvioPer_Establecimiento where cod_cia='" & wcia & "' and status<>'*' order by cod_establecimiento"

If fAbrRst(Re, Sql) Then
    Do While Not Re.EOF
            
            RsEmpEnvioPer_Establecimientos.AddNew
            RsEmpEnvioPer_Establecimientos!RUC = Trim(Re!RUC & "")
            RsEmpEnvioPer_Establecimientos!codest = Trim(Re!cod_establecimiento & "")
            RsEmpEnvioPer_Establecimientos!tipest = Trim(Re!tipo_establecimiento & "")
            RsEmpEnvioPer_Establecimientos!nomtipest = Trim(Re!nomtip_establecimiento & "")
            RsEmpEnvioPer_Establecimientos!nomest = Trim(Re!nom_establecimiento & "")
            RsEmpEnvioPer_Establecimientos!centro_riesgo = IIf(Re!indicador_centro_riesgo = True, True, False)
            
        Re.MoveNext
    Loop
End If
EliminarRst RsEmpEnvioPer_Establecimientos_Tasa
Sql = "select * "
Sql = Sql & " from plaCiaEmpEnvioPer_Establecimiento_Tasa where cod_cia='" & wcia & "' and status<>'*' "
If fAbrRst(Re, Sql) Then
    Do While Not Re.EOF
            
            RsEmpEnvioPer_Establecimientos_Tasa.AddNew
            RsEmpEnvioPer_Establecimientos_Tasa!RUC = Trim(Re!RUC & "")
            RsEmpEnvioPer_Establecimientos_Tasa!codest = Trim(Re!cod_establecimiento & "")
            RsEmpEnvioPer_Establecimientos_Tasa!PORC = Format(Re!Tasa & "", "##0.00")
            
        Re.MoveNext
    Loop
End If



'//**** Empleadores que me destacan o desplazan personal ******/////

EliminarRst RsEmpTraePer
EliminarRst RsEmpTraePer_Actividad

Sql = "select * "
Sql = Sql & " from plaCiaEmpTraePer where cod_cia='" & wcia & "' and status<>'*' "
If fAbrRst(Re, Sql) Then
    Do While Not Re.EOF
            RsEmpTraePer.AddNew
            RsEmpTraePer!RUC = Trim(Re!RUC & "")
            RsEmpTraePer!razsoc = Trim(Re!razsoc & "")
        Re.MoveNext
    Loop
End If

Sql = "select * "
Sql = Sql & " from plaCiaEmptraePer_Actividad where cod_cia='" & wcia & "' and status<>'*' "
If fAbrRst(Re, Sql) Then
    Do While Not Re.EOF
            RsEmpTraePer_Actividad.AddNew
            RsEmpTraePer_Actividad!RUC = Trim(Re!RUC & "")
            RsEmpTraePer_Actividad!cod_serv = Trim(Re!cod_serv & "")
        Re.MoveNext
    Loop
End If

Re.Close
Set Re = Nothing

End Sub
Public Function Graba_Cia()
Dim Meps As String
Dim Messalud As String
Dim Msenati As String
Dim MTipoTrabSunat As String
Dim MAportePatronal As String
Meps = ""
Messalud = ""
Msenati = ""
If txtrazsoc.Text = "" Then MsgBox "Debe Ingresar Razon Social", vbInformation, "Mantenimiento de Cia.": txtrazsoc.SetFocus: Exit Function
If txtruc.Text = "" Then MsgBox "Debe Ingresar Ruc", vbInformation, "Mantenimiento de Cia.": txtruc.SetFocus: Exit Function
'If txtreg.Text = "" Then MsgBox "Debe Ingresar Registro Patronal", vbInformation, "Mantenimiento de Cia.": txtreg.SetFocus: Exit Function
'If cmbgiro.Text = "" Then MsgBox "Debe Seleccionar Giro", vbInformation, "Mantenimiento de Cia.": cmbgiro.SetFocus: Exit Function
'If txtcob.Text = "" Then MsgBox "Debe Ingresar Codigo de Cobertura", vbInformation, "Mantenimiento de Cia.": txtcob.SetFocus: Exit Function
If txtdir.Text = "" Then MsgBox "Debe Ingresar Direccion", vbInformation, "Mantenimiento de Cia.": txtdir.SetFocus: Exit Function
If txtnro.Text = "" Then MsgBox "Debe Ingresar Numero", vbInformation, "Mantenimiento de Cia.": txtnro.SetFocus: Exit Function
If lbllugar.Caption = "" Then MsgBox "Debe Indicar Ubigeo", vbInformation, "Mantenimiento de Cia.": cmdubi.SetFocus: Exit Function
If txtrlegal.Text = "" Then MsgBox "Debe Ingresar Representante Legal", vbInformation, "Mantenimiento de Cia.": txtrlegal.SetFocus: Exit Function

Dim situ$
If opcactivo = True Then
   situ$ = ""
Else
   If opcanula = True Then situ$ = "*" Else situ$ = "I"
End If

'COMPROBAR SI SE REPITE RUC
Sql$ = "SELECT ruc FROM CIA WHERE cod_cia <> '" & Trim(wcia) & "' AND ruc = '" & Trim(txtruc) & "'"
Set rs = cn.Execute(Sql$)
If (fAbrRst(rs, Sql$)) Then
   MsgBox "El RUC está duplicado en Otra Compañía", vbExclamation
   If rs.State = 1 Then rs.Close
   Exit Function
End If
If rs.State = 1 Then rs.Close

Mgrab = MsgBox("Seguro de Grabar Compañia", vbYesNo + vbQuestion, TitMsg)
If Mgrab <> 6 Then Exit Function
If Txtaccidente.Text = "" Then Txtaccidente.Text = "0.00"
If Opceps.Value = True Then Meps = "S"
If Opcessalud.Value = True Then Messalud = "S"
If Chksenati.Value = 1 Then Msenati = "S": wSenati = "S" Else wSenati = ""
If ChkTrabSunat.Value = True Then MTipoTrabSunat = "S" Else MTipoTrabSunat = ""
If Me.ChkAportePatronal.Value = True Then MAportePatronal = "S" Else MAportePatronal = ""

Screen.MousePointer = vbArrowHourglass
Sql$ = wInicioTrans
cn.Execute Sql$
If Trim(txtcia.Text) = "" Then
   Sql$ = "select max(cod_cia) from cia"
   Set rs = cn.Execute(Sql$)
   If (fAbrRst(rs, Sql$)) Then
      wcia = Format(rs(0) + 1, "00")
   Else
      wcia = "01"
   End If
   If rs.State = 1 Then rs.Close
'   Sql = "INSERT INTO cia(cod_cia,razsoc,direcc,nro,dpto,urb,ref,cod_ubi,ruc, giro_emp, email," _
'        & "cod_cober, rep_legal,rep_nom, rep_nac, " _
'        & "reg_pat,ctedialib,ctemoros,ctefacvol,cteigv,ctereg,ctebajovol,status,user_crea,cteptovta,ctemsg,ctemonedac,fec_crea,acctrab,sctreps,sctressalud,afpresponsable,afparearesp,afpresptlf,afpnrocta,afpbanco,afptipocta,senati,cod_zona,tipo_doc,num_doc,tipo_esta,condicion_esta,domicilio_fiscal,cod_via) " _
'        & " VALUES('" & Trim(wcia) & "','" & Apostrofe(Trim(TxtRazsoc)) & "','" & Trim(Txtdir) & "','" & Trim(TxtNro) & "'," _
'        & "'" & Trim(txtdpto) & "','" & Trim(VNUrb) & "','" & Trim(txtref) & "','" & lbllugar.Tag & "','" & Trim(TxtRuc) & "'," _
'        & "'" & txtgiro.Tag & "','" & Trim(txtmail) & "','" & Trim(txtcob) & "'," _
'        & "'" & cod_rlegal & "','" & Trim(txtrlegal) & "','" & Trim(cbonacion.ReturnCodigo) & "','" & Trim(txtreg) & "',0,0,0,0,'',' ', " _
'        & "'" & situ & "','" & wuser & "','','" & Trim(txtmsg) & "','S/.'," & FechaSys & "," & CCur(Txtaccidente.Text) & ",'" & Meps & "','" & Messalud & "', " _
'        & "'','','','','','','" & Msenati & "'," & IIf(cbozona.ReturnCodigo = 0, "NULL", "'" & Format(cbozona.ReturnCodigo, "00") & "'") & "," & IIf(cbotipdoc.ReturnCodigo = 0, "NULL", "'" & Format(cbotipdoc.ReturnCodigo, "00") & "'") & "," & _
'        IIf(Len(Trim(txtdoc.Text)) = 0, "NULL", "'" & Trim(txtdoc.Text) & "'") & "," & IIf(cboestablecimiento.ReturnCodigo = 0, "NULL", "'" & Format(cboestablecimiento.ReturnCodigo, "00") & "'") & "," & IIf(cbocondicion.ReturnCodigo = 0 Or cbocondicion.ReturnCodigo = -1, "NULL", "'" & Format(cbocondicion.ReturnCodigo, "00") & "'") & _
'        "," & IIf(optdomfiscal.Value = False, 1, 0) & "," & IIf(cbovia.ReturnCodigo = 0, "NULL", "'" & Format(cbovia.ReturnCodigo, "00") & "'") & ")"
Sql = "INSERT INTO cia(cod_cia,razsoc,direcc,nro,dpto,urb,ref,cod_ubi,ruc, giro_emp, email," _
        & " rep_legal,rep_nom, rep_nac, " _
        & "ctedialib,ctemoros,ctefacvol,cteigv,ctereg,ctebajovol,status,user_crea,cteptovta,ctemsg,ctemonedac,fec_crea,acctrab,sctreps,sctressalud,afpresponsable,afparearesp,afpresptlf,afpnrocta,afpbanco,afptipocta,senati,cod_zona,tipo_doc,num_doc,tipo_esta,condicion_esta,domicilio_fiscal,cod_via,Aportacion_EPS,cargo_rep_legal,TipoTrabSunat,AportePatronal,SysContable) " _
        & " VALUES('" & Trim(wcia) & "','" & Apostrofe(Trim(txtrazsoc)) & "','" & Trim(txtdir) & "','" & Trim(txtnro) & "'," _
        & "'" & Trim(txtdpto) & "','" & Trim(VNUrb) & "','" & Trim(txtref) & "','" & lbllugar.Tag & "','" & Trim(txtruc) & "'," _
        & "'" & txtgiro.Tag & "','" & Trim(txtmail) & "'," _
        & "'" & cod_rlegal & "','" & Trim(txtrlegal) & "','" & Trim(cbonacion.ReturnCodigo) & "',0,0,0,0,'',' ', " _
        & "'" & situ & "','" & wuser & "','','" & Trim(txtmsg) & "','S/.'," & FechaSys & "," & CCur(Txtaccidente.Text) & ",'" & Meps & "','" & Messalud & "', " _
        & "'','','','','','','" & Msenati & "'," & IIf(cbozona.ReturnCodigo = 0, "NULL", "'" & Format(cbozona.ReturnCodigo, "00") & "'") & "," & IIf(cbotipdoc.ReturnCodigo = 0, "NULL", "'" & Format(cbotipdoc.ReturnCodigo, "00") & "'") & "," & _
        IIf(Len(Trim(txtdoc.Text)) = 0, "NULL", "'" & Trim(txtdoc.Text) & "'") & "," & IIf(cboestablecimiento.ReturnCodigo = 0, "NULL", "'" & Format(cboestablecimiento.ReturnCodigo, "00") & "'") & "," & IIf(cbocondicion.ReturnCodigo = 0 Or cbocondicion.ReturnCodigo = -1, "NULL", "'" & Format(cbocondicion.ReturnCodigo, "00") & "'") & _
        "," & IIf(optdomfiscal.Value = False, 1, 0) & "," & IIf(cbovia.ReturnCodigo = 0, "NULL", "'" & Format(cbovia.ReturnCodigo, "00") & "'") & ", " & IIf(Chk_EPS.Value = Checked, 1, 0) & ",'" & Trim(TxtCargo_Rep_Legal.Text) & "','" & MTipoTrabSunat & "','" & MAportePatronal & "','01')"
Else
'   Sql = "UPDATE cia SET razsoc ='" & Apostrofe(Trim(txtrazsoc)) & "', direcc ='" & Trim(txtdir) & "', nro ='" & Trim(txtnro) & "', dpto ='" & Trim(txtdpto) & "', urb ='" & Trim(VNUrb) & "'," _
'       & "ref ='" & Trim(txtref) & "', cod_ubi ='" & lbllugar.Tag & "', ruc ='" & Trim(txtruc) & "'," _
'       & "giro_emp ='" & Format("", "00") & "', email ='" & Trim(txtmail) & "', cod_cober ='" & Trim(txtcob) & "'," _
'       & "rep_legal ='" & Trim(cod_rlegal) & "',rep_nac='" & IIf(cbonacion.ReturnCodigo = 0 Or cbonacion.ReturnCodigo = -1, "NULL", cbonacion.ReturnCodigo) & "', rep_nom ='" & Trim(txtrlegal) & "', reg_pat ='" & Trim(txtreg) & "'," _
'       & "ctemsg='" & Trim(txtmsg) & "',cod_zona=" & IIf(cbozona.ReturnCodigo = 0 Or cbozona.ReturnCodigo = -1, "NULL", "'" & Format(cbozona.ReturnCodigo, "00") & "'") & _
'       ",tipo_doc=" & IIf(cbotipdoc.ReturnCodigo = 0 Or cbotipdoc.ReturnCodigo = -1, "NULL", "'" & Format(cbotipdoc.ReturnCodigo, "00") & "'") & _
'       ",num_doc=" & IIf(Len(Trim(txtdoc.Text)) = 0, "NULL", "'" & Trim(txtdoc.Text) & "'") & ",tipo_esta=" & IIf(cboestablecimiento.ReturnCodigo = 0 Or cboestablecimiento.ReturnCodigo = -1, "NULL", "'" & Format(cboestablecimiento.ReturnCodigo, "00") & "'") & _
'       ",condicion_esta=" & IIf(cbocondicion.ReturnCodigo = 0 Or cbocondicion.ReturnCodigo = -1, "NULL", "'" & Format(cbocondicion.ReturnCodigo, "00") & "'") & ",domicilio_fiscal=" & IIf(optdomfiscal.Value = False, 1, 0) & ",cod_via=" & IIf(cbovia.ReturnCodigo = 0 Or cbovia.ReturnCodigo = -1, "NULL", "'" & Format(cbovia.ReturnCodigo, "00") & "'") _
'       & ",status='" & situ & "',user_modi='" & wuser & "',fec_modi=" & FechaSys & ", " _
'       & "acctrab=" & CCur(Txtaccidente.Text) & ",sctreps='" & Meps & "',sctressalud='" & Messalud & "', " _
'       & "senati='" & Msenati & "' " _
'       & "WHERE cod_cia='" & Trim(wcia) & "'"
   Sql = "UPDATE cia SET razsoc ='" & Apostrofe(Trim(txtrazsoc)) & "', direcc ='" & Trim(txtdir) & "', nro ='" & Trim(txtnro) & "', dpto ='" & Trim(txtdpto) & "', urb ='" & Trim(VNUrb) & "'," _
       & "ref ='" & Trim(txtref) & "', cod_ubi ='" & lbllugar.Tag & "', ruc ='" & Trim(txtruc) & "'," _
       & "giro_emp ='" & Format("", "00") & "', email ='" & Trim(txtmail) & "'," _
       & "rep_legal ='" & Trim(cod_rlegal) & "',rep_nac='" & IIf(cbonacion.ReturnCodigo = 0 Or cbonacion.ReturnCodigo = -1, "NULL", cbonacion.ReturnCodigo) & "', rep_nom ='" & Trim(txtrlegal) & "'," _
       & "ctemsg='" & Trim(txtmsg) & "',cod_zona=" & IIf(cbozona.ReturnCodigo = 0 Or cbozona.ReturnCodigo = -1, "NULL", "'" & Format(cbozona.ReturnCodigo, "00") & "'") & _
       ",tipo_doc=" & IIf(cbotipdoc.ReturnCodigo = 0 Or cbotipdoc.ReturnCodigo = -1, "NULL", "'" & Format(cbotipdoc.ReturnCodigo, "00") & "'") & _
       ",num_doc=" & IIf(Len(Trim(txtdoc.Text)) = 0, "NULL", "'" & Trim(txtdoc.Text) & "'") & ",tipo_esta=" & IIf(cboestablecimiento.ReturnCodigo = 0 Or cboestablecimiento.ReturnCodigo = -1, "NULL", "'" & Format(cboestablecimiento.ReturnCodigo, "00") & "'") & _
       ",condicion_esta=" & IIf(cbocondicion.ReturnCodigo = 0 Or cbocondicion.ReturnCodigo = -1, "NULL", "'" & Format(cbocondicion.ReturnCodigo, "00") & "'") & ",domicilio_fiscal=" & IIf(optdomfiscal.Value = False, 1, 0) & ",cod_via=" & IIf(cbovia.ReturnCodigo = 0 Or cbovia.ReturnCodigo = -1, "NULL", "'" & Format(cbovia.ReturnCodigo, "00") & "'") _
       & ",status='" & situ & "',user_modi='" & wuser & "',fec_modi=" & FechaSys & ", " _
       & "acctrab=" & CCur(Txtaccidente.Text) & ",sctreps='" & Meps & "',sctressalud='" & Messalud & "', " _
       & "senati='" & Msenati & "', Aportacion_EPS = " & IIf(Chk_EPS.Value = Checked, 1, 0) & ",cargo_rep_legal='" & Trim(TxtCargo_Rep_Legal.Text) & "',TipoTrabSunat='" & MTipoTrabSunat & "', AportePatronal='" & MAportePatronal & "'" _
       & "WHERE cod_cia='" & Trim(wcia) & "'"
End If

cn.Execute Sql$

Sql$ = "delete from telef_cia where cod_cia='" & VCia & "'"
cn.Execute Sql$
If rstlf.RecordCount > 0 Then rstlf.MoveFirst
Do While Not rstlf.EOF
   Sql$ = "INSERT INTO telef_cia values('" & wcia & "','" & rstlf!telefono & "','" & rstlf!fax & "','" & wuser & "'," & FechaSys & ",'" & rstlf!referencia & "')"
   cn.Execute Sql$
   rstlf.MoveNext
Loop

Sql = "update placiaestablecimientos set status='*',user_modi='" & Trim(wuser) & "',fec_modi=getdate() where cod_cia='" & wcia & "' and status<>'*'"
cn.Execute Sql, 64
Sql = "update placiaestablecimientostasa set status='*',user_modi='" & Trim(wuser) & "',fec_modi=getdate() where cod_cia='" & wcia & "' and status<>'*'"
cn.Execute Sql, 64
With Me.RsEstablecimientos
    If .RecordCount > 0 Then
        .MoveFirst
        Do While Not .EOF
                Sql = "insert into placiaestablecimientos values('" & wcia & "','" & Trim(!codest) & "','" & Trim(!tipest) & "','" & Apostrofe(Trim(!nomest)) & "'," & IIf(!centro_riesgo = True, "1", "0") & ",'','" & Trim(wuser) & "',getdate(),null,null)"
                cn.Execute Sql, 64
            .MoveNext
        Loop
    End If
End With

With Me.RsEstTasa
    If .RecordCount > 0 Then
        .MoveFirst
        Do While Not .EOF
                Sql = "insert into placiaestablecimientostasa values('" & wcia & "','" & Trim(!codest) & "'," & CCur(!PORC) & ",'','" & Trim(wuser) & "',getdate(),null,null)"
                cn.Execute Sql, 64
            .MoveNext
        Loop
    End If
End With

'//**** Empleadores a quienes destacó o desplazo personal ****/////

Sql = "update plaCiaEmpEnvioPer set status='*',user_modi='" & Trim(wuser) & "',fec_modi=getdate() where cod_cia='" & wcia & "' and status<>'*'"
cn.Execute Sql, 64
Sql = "update plaCiaEmpEnvioPer_Actividad set status='*',user_modi='" & Trim(wuser) & "',fec_modi=getdate() where cod_cia='" & wcia & "' and status<>'*'"
cn.Execute Sql, 64
With Me.RsEmpEnvioPer
    If .RecordCount > 0 Then
        .MoveFirst
        Do While Not .EOF
                Sql = "insert into plaCiaEmpEnvioPer values('" & wcia & "','" & Trim(!RUC) & "','" & Apostrofe(Trim(!razsoc)) & "','','" & Trim(wuser) & "',getdate(),null,null)"
                cn.Execute Sql, 64
            .MoveNext
        Loop
    End If
End With

With Me.RsEmpEnvioPer_Actividad
    If .RecordCount > 0 Then
        .MoveFirst
        Do While Not .EOF
                Sql = "insert into plaCiaEmpEnvioPer_Actividad values('" & wcia & "','" & Trim(!RUC) & "'," & Trim(!cod_serv) & ",'','" & Trim(wuser) & "',getdate(),null,null)"
                cn.Execute Sql, 64
            .MoveNext
        Loop
    End If
End With


Sql = "update plaCiaEmpEnvioPer_Establecimiento set status='*',user_modi='" & Trim(wuser) & "',fec_modi=getdate() where cod_cia='" & wcia & "' and status<>'*'"
cn.Execute Sql, 64
Sql = "update plaCiaEmpEnvioPer_Establecimiento_Tasa set status='*',user_modi='" & Trim(wuser) & "',fec_modi=getdate() where cod_cia='" & wcia & "' and status<>'*'"
cn.Execute Sql, 64
With Me.RsEmpEnvioPer_Establecimientos
    If .RecordCount > 0 Then
        .MoveFirst
        Do While Not .EOF
                Sql = "insert into plaCiaEmpEnvioPer_Establecimiento values('" & wcia & "','" & Trim(!RUC) & "','" & Trim(!codest) & "','" & Trim(!tipest) & "','" & Apostrofe(Trim(!nomest)) & "'," & IIf(!centro_riesgo = True, "1", "0") & ",'','" & Trim(wuser) & "',getdate(),null,null)"
                cn.Execute Sql, 64
            .MoveNext
        Loop
    End If
End With

With Me.RsEmpEnvioPer_Establecimientos_Tasa
    If .RecordCount > 0 Then
        .MoveFirst
        Do While Not .EOF
                Sql = "insert into plaCiaEmpEnvioPer_Establecimiento_Tasa values('" & wcia & "','" & Trim(!RUC) & "','" & Trim(!codest) & "'," & CCur(!PORC) & ",'','" & Trim(wuser) & "',getdate(),null,null)"
                cn.Execute Sql, 64
            .MoveNext
        Loop
    End If
End With


'//**** Empleadores que me destacan o desplazan personal ******/////

Sql = "update plaCiaEmpTraePer set status='*',user_modi='" & Trim(wuser) & "',fec_modi=getdate() where cod_cia='" & wcia & "' and status<>'*'"
cn.Execute Sql, 64
Sql = "update plaCiaEmpTraePer_Actividad set status='*',user_modi='" & Trim(wuser) & "',fec_modi=getdate() where cod_cia='" & wcia & "' and status<>'*'"
cn.Execute Sql, 64
With Me.RsEmpTraePer
    If .RecordCount > 0 Then
        .MoveFirst
        Do While Not .EOF
                Sql = "insert into plaCiaEmpTraePer values('" & wcia & "','" & Trim(!RUC) & "','" & Apostrofe(Trim(!razsoc)) & "','','" & Trim(wuser) & "',getdate(),null,null)"
                cn.Execute Sql, 64
            .MoveNext
        Loop
    End If
End With

With Me.RsEmpTraePer_Actividad
    If .RecordCount > 0 Then
        .MoveFirst
        Do While Not .EOF
                Sql = "insert into plaCiaEmpTraePer_Actividad values('" & wcia & "','" & Trim(!RUC) & "'," & Trim(!cod_serv) & ",'','" & Trim(wuser) & "',getdate(),null,null)"
                cn.Execute Sql, 64
            .MoveNext
        Loop
    End If
End With

Sql$ = wFinTrans
cn.Execute Sql$
MDIplared.Activa_Menu
Screen.MousePointer = vbDefault
Unload Me
End Function

Private Sub Form_Unload(Cancel As Integer)
Frmgrdcia.Procesa_Cia
End Sub

Private Sub Opceps_Click()
If Opceps.Value = True Then Chkscrt.Value = True
If Opceps.Value = False And Opcessalud.Value = False Then Chkscrt.Value = False
End Sub

Private Sub Opcessalud_Click()
If Opcessalud.Value = True Then Chkscrt.Value = True
If Opceps.Value = False And Opcessalud.Value = False Then Chkscrt.Value = False
End Sub

Private Sub SSCommand5_Click()
FrameTelf.Visible = False
cmdtelef.Enabled = True
End Sub

Private Sub Txtaccidente_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
       Case 8, 13, 27, 32, 42, 44, 46, 48 To 57
            KeyAscii = KeyAscii
       Case Else
            KeyAscii = 0
            Exit Sub
End Select
If KeyAscii = 13 Then Txtaccidente.Text = Format(Txtaccidente.Text, "###,###.00")
End Sub
Private Sub Txtuit_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
       Case 8, 13, 27, 32, 42, 44, 46, 48 To 57
            KeyAscii = KeyAscii
       Case Else
            KeyAscii = 0
            Exit Sub
End Select
If KeyAscii = 13 Then Txtuit.Text = Format(Txtuit.Text, "###,###.00"): Txtaccidente.SetFocus
End Sub
Public Sub Nueva_cia()
txtrazsoc.Text = ""
txtruc.Text = ""
txtreg.Text = ""
'cmbgiro.ListIndex = -1
txtcob.Text = ""
txtmail.Text = ""
txtdir.Text = ""
txtnro.Text = ""
txtdpto.Text = ""
'cmburb.ListIndex = -1
txtref.Text = ""
lbllugar.Caption = ""
txtrlegal.Text = ""
'lblnac.Caption = ""
'txtrruc.Text = ""
'txtrle.Text = ""
'txtrpas.Text = ""
lblfecing.Caption = ""
lblfecmodi.Caption = ""
Chk_EPS.Value = 0
ChkAportePatronal.Value = False


'If RsEstablecimientos.RecordCount > 0 Then RsEstablecimientos.MoveFirst
'Do While Not RsEstablecimientos.EOF
'   RsEstablecimientos.Delete
'   RsEstablecimientos.MoveNext
'Loop
'
'If RsEstTasa.RecordCount > 0 Then RsEstTasa.MoveFirst
'Do While Not RsEstTasa.EOF
'   RsEstTasa.Delete
'   RsEstTasa.MoveNext
'Loop
ChkTrabSunat.Value = False

EliminarRst RsEstablecimientos
EliminarRst RsEstTasa

EliminarRst RsEmpEnvioPer
EliminarRst RsEmpEnvioPer_Actividad
EliminarRst RsEmpEnvioPer_Establecimientos
EliminarRst RsEmpEnvioPer_Establecimientos_Tasa

EliminarRst RsEmpTraePer
EliminarRst RsEmpTraePer_Actividad
End Sub

Private Function getFilter() As String

    'Creates the SQL statement in adodc1.recordset.filter

    'and only filters text currently. It must be modified to
    'filter other data types.
    Dim tmp As String

    Dim n As Integer

    For Each Col In Cols

        If Trim(Col.FilterText) <> "" Then

            n = n + 1

            If n > 1 Then

                tmp = tmp & " AND "

            End If

            tmp = tmp & Col.DataField & " LIKE '*" & Col.FilterText & "*'"

        End If

    Next Col
    getFilter = tmp
End Function

Private Sub Crea_Rs()
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
    
    '//****** Empresa a la destaco personal **********////
    'empresas
    If RsEmpEnvioPer.State = 1 Then RsEmpEnvioPer.Close
    RsEmpEnvioPer.Fields.Append "ruc", adChar, 11, adFldIsNullable
    RsEmpEnvioPer.Fields.Append "razsoc", adVarChar, 100, adFldIsNullable
    RsEmpEnvioPer.Open
    Set DgrdEnvPer.DataSource = RsEmpEnvioPer
        
    'actividad
    If RsEmpEnvioPer_Actividad.State = 1 Then RsEmpEnvioPer_Actividad.Close
    RsEmpEnvioPer_Actividad.Fields.Append "ruc", adChar, 11, adFldIsNullable
    RsEmpEnvioPer_Actividad.Fields.Append "cod_serv", adChar, 6, adFldIsNullable
    RsEmpEnvioPer_Actividad.Open
    
    'Establecimientos
    If RsEmpEnvioPer_Establecimientos.State = 1 Then RsEmpEnvioPer_Establecimientos.Close
    RsEmpEnvioPer_Establecimientos.Fields.Append "ruc", adChar, 11, adFldIsNullable
    RsEmpEnvioPer_Establecimientos.Fields.Append "codest", adChar, 4, adFldIsNullable
    RsEmpEnvioPer_Establecimientos.Fields.Append "tipest", adChar, 2, adFldIsNullable
    RsEmpEnvioPer_Establecimientos.Fields.Append "nomtipest", adChar, 40, adFldIsNullable
    RsEmpEnvioPer_Establecimientos.Fields.Append "nomest", adChar, 40, adFldIsNullable
    RsEmpEnvioPer_Establecimientos.Fields.Append "centro_riesgo", adBoolean, , adFldIsNullable
    RsEmpEnvioPer_Establecimientos.Open
    
    'tasa
    If RsEmpEnvioPer_Establecimientos_Tasa.State = 1 Then RsEmpEnvioPer_Establecimientos_Tasa.Close
    RsEmpEnvioPer_Establecimientos_Tasa.Fields.Append "ruc", adChar, 11, adFldIsNullable
    RsEmpEnvioPer_Establecimientos_Tasa.Fields.Append "codest", adChar, 4, adFldIsNullable
    RsEmpEnvioPer_Establecimientos_Tasa.Fields.Append "porc", adChar, 6, adFldIsNullable
    RsEmpEnvioPer_Establecimientos_Tasa.Open
    
    
     '//****** Empleadores que me destacan o desplazan personal **********////
    'empresas
    If RsEmpTraePer.State = 1 Then RsEmpTraePer.Close
    RsEmpTraePer.Fields.Append "ruc", adChar, 11, adFldIsNullable
    RsEmpTraePer.Fields.Append "razsoc", adVarChar, 100, adFldIsNullable
    RsEmpTraePer.Open
    Set DgrdTraePer.DataSource = RsEmpTraePer
        
    'actividad
    If RsEmpTraePer_Actividad.State = 1 Then RsEmpTraePer_Actividad.Close
    RsEmpTraePer_Actividad.Fields.Append "ruc", adChar, 11, adFldIsNullable
    RsEmpTraePer_Actividad.Fields.Append "cod_serv", adChar, 6, adFldIsNullable
    RsEmpTraePer_Actividad.Open
         
End Sub

Public Sub Eliminar_Reg(ByRef pRs As ADODB.Recordset, ByRef pId As String)
    With pRs
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                    If Trim(!RUC) = Trim(pId) Then
                        .Delete
                        .Update
                    End If
                .MoveNext
            Loop
        End If
    End With
End Sub

Public Sub EliminarRst(ByRef pRs As ADODB.Recordset)
    With pRs
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                .Delete
                .MoveNext
            Loop
        End If
    End With
End Sub



