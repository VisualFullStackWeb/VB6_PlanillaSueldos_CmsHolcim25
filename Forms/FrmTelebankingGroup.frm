VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmTelebankingGroup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "*** Archivos de Depositos al Banco ***"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   14400
   Begin VB.Frame Frame2 
      Caption         =   "Datos Generales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   14175
      Begin VB.CheckBox chkFechaPago 
         Caption         =   "Fecha Pago:"
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
         TabIndex        =   16
         Top             =   720
         Width           =   1455
      End
      Begin VB.ComboBox Cmbtipotrabajador 
         Height          =   315
         ItemData        =   "FrmTelebankingGroup.frx":0000
         Left            =   6120
         List            =   "FrmTelebankingGroup.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   720
         Width           =   2055
      End
      Begin VB.Frame Frame1 
         Caption         =   "Estado"
         Height          =   975
         Left            =   11760
         TabIndex        =   9
         Top             =   120
         Width           =   2295
         Begin VB.OptionButton optTodos 
            Caption         =   "Todos"
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
            Left            =   480
            TabIndex        =   12
            Top             =   680
            Value           =   -1  'True
            Width           =   1455
         End
         Begin VB.OptionButton optCancelado 
            Caption         =   "Cancelado"
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
            Left            =   480
            TabIndex        =   11
            Top             =   440
            Width           =   1455
         End
         Begin VB.OptionButton optPendiente 
            Caption         =   "Pendiente"
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
            Left            =   480
            TabIndex        =   10
            Top             =   200
            Width           =   1335
         End
      End
      Begin VB.ComboBox CmbBcoPago 
         Height          =   315
         ItemData        =   "FrmTelebankingGroup.frx":0004
         Left            =   840
         List            =   "FrmTelebankingGroup.frx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   240
         Width           =   4335
      End
      Begin VB.ComboBox CmbMonPago 
         Height          =   315
         ItemData        =   "FrmTelebankingGroup.frx":0008
         Left            =   6120
         List            =   "FrmTelebankingGroup.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
      Begin VB.ComboBox cboCuenta 
         Height          =   315
         ItemData        =   "FrmTelebankingGroup.frx":000C
         Left            =   7800
         List            =   "FrmTelebankingGroup.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   3615
      End
      Begin MSComCtl2.DTPicker dtpFecha1 
         Height          =   285
         Left            =   1560
         TabIndex        =   4
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Format          =   63504385
         CurrentDate     =   37265
      End
      Begin MSComCtl2.DTPicker dtpFecha2 
         Height          =   285
         Left            =   3000
         TabIndex        =   13
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Format          =   63504385
         CurrentDate     =   37265
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Tipo Trabajador:"
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
         Left            =   4620
         TabIndex        =   15
         Top             =   720
         Width           =   1425
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Banco"
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
         TabIndex        =   7
         Top             =   360
         Width           =   555
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Moneda:"
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
         Left            =   5280
         TabIndex        =   6
         Top             =   360
         Width           =   750
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta:"
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
         Left            =   7080
         TabIndex        =   5
         Top             =   360
         Width           =   675
      End
   End
   Begin TrueOleDBGrid70.TDBGrid GrdSolicitud 
      Height          =   2775
      Left            =   120
      TabIndex        =   8
      Top             =   1440
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   4895
      _LayoutType     =   4
      _RowHeight      =   17
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "IdTelebanking"
      Columns(0).DataField=   "IdTelebanking"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Fecha Pago"
      Columns(1).DataField=   "FechaProceso"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Planilla"
      Columns(2).DataField=   "Planilla"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "CodBanco"
      Columns(3).DataField=   "CodBanco"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Banco"
      Columns(4).DataField=   "Banco"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Moneda"
      Columns(5).DataField=   "Moneda"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Importe"
      Columns(6).DataField=   "Importe"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "NroCta"
      Columns(7).DataField=   "NroCta"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "CodTipoTrab"
      Columns(8).DataField=   "CodTipoTrab"
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "TipoTrab"
      Columns(9).DataField=   "TipoTrab"
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "TipoRegistro"
      Columns(10).DataField=   "TipoRegistro"
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "NroCtaTrabCia"
      Columns(11).DataField=   "NroCtaTrabCia"
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "Referencia"
      Columns(12).DataField=   "Referencia"
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).Caption=   "Estado"
      Columns(13).DataField=   "Estado"
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(14)._VlistStyle=   0
      Columns(14)._MaxComboItems=   5
      Columns(14).Caption=   "status"
      Columns(14).DataField=   "status"
      Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   15
      Splits(0)._UserFlags=   0
      Splits(0).MarqueeStyle=   3
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).FetchRowStyle=   -1  'True
      Splits(0).DividerColor=   13160660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=15"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2117"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2037"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=8193"
      Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(8)=   "Column(1).Width=2117"
      Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=2037"
      Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=8196"
      Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(14)=   "Column(2).Width=2725"
      Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=2646"
      Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=8196"
      Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(20)=   "Column(3).Width=2725"
      Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=2646"
      Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=8196"
      Splits(0)._ColumnProps(25)=   "Column(3).Visible=0"
      Splits(0)._ColumnProps(26)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(27)=   "Column(4).Width=5292"
      Splits(0)._ColumnProps(28)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(29)=   "Column(4)._WidthInPix=5212"
      Splits(0)._ColumnProps(30)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(31)=   "Column(4)._ColStyle=8196"
      Splits(0)._ColumnProps(32)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(33)=   "Column(5).Width=1244"
      Splits(0)._ColumnProps(34)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(35)=   "Column(5)._WidthInPix=1164"
      Splits(0)._ColumnProps(36)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(37)=   "Column(5)._ColStyle=8196"
      Splits(0)._ColumnProps(38)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(39)=   "Column(6).Width=2725"
      Splits(0)._ColumnProps(40)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(41)=   "Column(6)._WidthInPix=2646"
      Splits(0)._ColumnProps(42)=   "Column(6)._EditAlways=0"
      Splits(0)._ColumnProps(43)=   "Column(6)._ColStyle=258"
      Splits(0)._ColumnProps(44)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(45)=   "Column(7).Width=3360"
      Splits(0)._ColumnProps(46)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(47)=   "Column(7)._WidthInPix=3281"
      Splits(0)._ColumnProps(48)=   "Column(7)._EditAlways=0"
      Splits(0)._ColumnProps(49)=   "Column(7)._ColStyle=8196"
      Splits(0)._ColumnProps(50)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(51)=   "Column(8).Width=2725"
      Splits(0)._ColumnProps(52)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(53)=   "Column(8)._WidthInPix=2646"
      Splits(0)._ColumnProps(54)=   "Column(8)._EditAlways=0"
      Splits(0)._ColumnProps(55)=   "Column(8)._ColStyle=8196"
      Splits(0)._ColumnProps(56)=   "Column(8).Visible=0"
      Splits(0)._ColumnProps(57)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(58)=   "Column(9).Width=3519"
      Splits(0)._ColumnProps(59)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(60)=   "Column(9)._WidthInPix=3440"
      Splits(0)._ColumnProps(61)=   "Column(9)._EditAlways=0"
      Splits(0)._ColumnProps(62)=   "Column(9)._ColStyle=8196"
      Splits(0)._ColumnProps(63)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(64)=   "Column(10).Width=2725"
      Splits(0)._ColumnProps(65)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(66)=   "Column(10)._WidthInPix=2646"
      Splits(0)._ColumnProps(67)=   "Column(10)._EditAlways=0"
      Splits(0)._ColumnProps(68)=   "Column(10)._ColStyle=8196"
      Splits(0)._ColumnProps(69)=   "Column(10).Visible=0"
      Splits(0)._ColumnProps(70)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(71)=   "Column(11).Width=2725"
      Splits(0)._ColumnProps(72)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(73)=   "Column(11)._WidthInPix=2646"
      Splits(0)._ColumnProps(74)=   "Column(11)._EditAlways=0"
      Splits(0)._ColumnProps(75)=   "Column(11)._ColStyle=8196"
      Splits(0)._ColumnProps(76)=   "Column(11).Visible=0"
      Splits(0)._ColumnProps(77)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(78)=   "Column(12).Width=2725"
      Splits(0)._ColumnProps(79)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(80)=   "Column(12)._WidthInPix=2646"
      Splits(0)._ColumnProps(81)=   "Column(12)._EditAlways=0"
      Splits(0)._ColumnProps(82)=   "Column(12)._ColStyle=8196"
      Splits(0)._ColumnProps(83)=   "Column(12).Visible=0"
      Splits(0)._ColumnProps(84)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(85)=   "Column(13).Width=2725"
      Splits(0)._ColumnProps(86)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(87)=   "Column(13)._WidthInPix=2646"
      Splits(0)._ColumnProps(88)=   "Column(13)._EditAlways=0"
      Splits(0)._ColumnProps(89)=   "Column(13)._ColStyle=8196"
      Splits(0)._ColumnProps(90)=   "Column(13).Order=14"
      Splits(0)._ColumnProps(91)=   "Column(14).Width=2725"
      Splits(0)._ColumnProps(92)=   "Column(14).DividerColor=0"
      Splits(0)._ColumnProps(93)=   "Column(14)._WidthInPix=2646"
      Splits(0)._ColumnProps(94)=   "Column(14)._EditAlways=0"
      Splits(0)._ColumnProps(95)=   "Column(14)._ColStyle=8196"
      Splits(0)._ColumnProps(96)=   "Column(14).Visible=0"
      Splits(0)._ColumnProps(97)=   "Column(14).Order=15"
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
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37,.bgcolor=&HC0C0C0&"
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
      _StyleDefs(39)  =   "Splits(0).Columns(0).Style:id=50,.parent=13,.alignment=2,.locked=-1"
      _StyleDefs(40)  =   "Splits(0).Columns(0).HeadingStyle:id=47,.parent=14"
      _StyleDefs(41)  =   "Splits(0).Columns(0).FooterStyle:id=48,.parent=15"
      _StyleDefs(42)  =   "Splits(0).Columns(0).EditorStyle:id=49,.parent=17"
      _StyleDefs(43)  =   "Splits(0).Columns(1).Style:id=28,.parent=13,.locked=-1"
      _StyleDefs(44)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
      _StyleDefs(45)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
      _StyleDefs(46)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
      _StyleDefs(47)  =   "Splits(0).Columns(2).Style:id=32,.parent=13,.locked=-1"
      _StyleDefs(48)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
      _StyleDefs(49)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
      _StyleDefs(50)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
      _StyleDefs(51)  =   "Splits(0).Columns(3).Style:id=46,.parent=13,.locked=-1"
      _StyleDefs(52)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
      _StyleDefs(53)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
      _StyleDefs(54)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
      _StyleDefs(55)  =   "Splits(0).Columns(4).Style:id=54,.parent=13,.locked=-1"
      _StyleDefs(56)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
      _StyleDefs(57)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
      _StyleDefs(58)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
      _StyleDefs(59)  =   "Splits(0).Columns(5).Style:id=58,.parent=13,.locked=-1"
      _StyleDefs(60)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
      _StyleDefs(61)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
      _StyleDefs(62)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
      _StyleDefs(63)  =   "Splits(0).Columns(6).Style:id=94,.parent=13,.alignment=1"
      _StyleDefs(64)  =   "Splits(0).Columns(6).HeadingStyle:id=91,.parent=14,.alignment=0"
      _StyleDefs(65)  =   "Splits(0).Columns(6).FooterStyle:id=92,.parent=15"
      _StyleDefs(66)  =   "Splits(0).Columns(6).EditorStyle:id=93,.parent=17"
      _StyleDefs(67)  =   "Splits(0).Columns(7).Style:id=62,.parent=13,.locked=-1"
      _StyleDefs(68)  =   "Splits(0).Columns(7).HeadingStyle:id=59,.parent=14"
      _StyleDefs(69)  =   "Splits(0).Columns(7).FooterStyle:id=60,.parent=15"
      _StyleDefs(70)  =   "Splits(0).Columns(7).EditorStyle:id=61,.parent=17"
      _StyleDefs(71)  =   "Splits(0).Columns(8).Style:id=66,.parent=13,.locked=-1"
      _StyleDefs(72)  =   "Splits(0).Columns(8).HeadingStyle:id=63,.parent=14"
      _StyleDefs(73)  =   "Splits(0).Columns(8).FooterStyle:id=64,.parent=15"
      _StyleDefs(74)  =   "Splits(0).Columns(8).EditorStyle:id=65,.parent=17"
      _StyleDefs(75)  =   "Splits(0).Columns(9).Style:id=70,.parent=13,.locked=-1"
      _StyleDefs(76)  =   "Splits(0).Columns(9).HeadingStyle:id=67,.parent=14"
      _StyleDefs(77)  =   "Splits(0).Columns(9).FooterStyle:id=68,.parent=15"
      _StyleDefs(78)  =   "Splits(0).Columns(9).EditorStyle:id=69,.parent=17"
      _StyleDefs(79)  =   "Splits(0).Columns(10).Style:id=74,.parent=13,.locked=-1"
      _StyleDefs(80)  =   "Splits(0).Columns(10).HeadingStyle:id=71,.parent=14"
      _StyleDefs(81)  =   "Splits(0).Columns(10).FooterStyle:id=72,.parent=15"
      _StyleDefs(82)  =   "Splits(0).Columns(10).EditorStyle:id=73,.parent=17"
      _StyleDefs(83)  =   "Splits(0).Columns(11).Style:id=78,.parent=13,.locked=-1"
      _StyleDefs(84)  =   "Splits(0).Columns(11).HeadingStyle:id=75,.parent=14"
      _StyleDefs(85)  =   "Splits(0).Columns(11).FooterStyle:id=76,.parent=15"
      _StyleDefs(86)  =   "Splits(0).Columns(11).EditorStyle:id=77,.parent=17"
      _StyleDefs(87)  =   "Splits(0).Columns(12).Style:id=82,.parent=13,.locked=-1"
      _StyleDefs(88)  =   "Splits(0).Columns(12).HeadingStyle:id=79,.parent=14"
      _StyleDefs(89)  =   "Splits(0).Columns(12).FooterStyle:id=80,.parent=15"
      _StyleDefs(90)  =   "Splits(0).Columns(12).EditorStyle:id=81,.parent=17"
      _StyleDefs(91)  =   "Splits(0).Columns(13).Style:id=86,.parent=13,.locked=-1"
      _StyleDefs(92)  =   "Splits(0).Columns(13).HeadingStyle:id=83,.parent=14"
      _StyleDefs(93)  =   "Splits(0).Columns(13).FooterStyle:id=84,.parent=15"
      _StyleDefs(94)  =   "Splits(0).Columns(13).EditorStyle:id=85,.parent=17"
      _StyleDefs(95)  =   "Splits(0).Columns(14).Style:id=90,.parent=13,.locked=-1"
      _StyleDefs(96)  =   "Splits(0).Columns(14).HeadingStyle:id=87,.parent=14"
      _StyleDefs(97)  =   "Splits(0).Columns(14).FooterStyle:id=88,.parent=15"
      _StyleDefs(98)  =   "Splits(0).Columns(14).EditorStyle:id=89,.parent=17"
      _StyleDefs(99)  =   "Named:id=33:Normal"
      _StyleDefs(100) =   ":id=33,.parent=0"
      _StyleDefs(101) =   "Named:id=34:Heading"
      _StyleDefs(102) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(103) =   ":id=34,.wraptext=-1"
      _StyleDefs(104) =   "Named:id=35:Footing"
      _StyleDefs(105) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(106) =   "Named:id=36:Selected"
      _StyleDefs(107) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(108) =   "Named:id=37:Caption"
      _StyleDefs(109) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(110) =   "Named:id=38:HighlightRow"
      _StyleDefs(111) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(112) =   "Named:id=39:EvenRow"
      _StyleDefs(113) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(114) =   "Named:id=40:OddRow"
      _StyleDefs(115) =   ":id=40,.parent=33"
      _StyleDefs(116) =   "Named:id=41:RecordSelector"
      _StyleDefs(117) =   ":id=41,.parent=34"
      _StyleDefs(118) =   "Named:id=42:FilterBar"
      _StyleDefs(119) =   ":id=42,.parent=33"
   End
End
Attribute VB_Name = "FrmTelebankingGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmbBcoPago_Click()
Call Cargar_Cuenta_Banco
End Sub

Private Sub CmbMonPago_Click()
Call Cargar_Cuenta_Banco
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0

Call fc_Descrip_Maestros2("01007", "", CmbBcoPago, False, True)
Call fc_Descrip_Maestros2_Mon("01006", "", CmbMonPago)
CmbBcoPago.AddItem "<<TODOS>>"
CmbBcoPago.ItemData(CmbBcoPago.NewIndex) = "999"
CmbBcoPago.ListIndex = CmbBcoPago.NewIndex

CmbMonPago.AddItem "<<TODOS>>"
CmbMonPago.ItemData(CmbMonPago.NewIndex) = "999"
CmbMonPago.ListIndex = CmbMonPago.NewIndex

Call fc_Descrip_Maestros2("01055", "", Cmbtipotrabajador)

Cmbtipotrabajador.AddItem "<<TODOS>>"
Cmbtipotrabajador.ItemData(Cmbtipotrabajador.NewIndex) = "999"
Cmbtipotrabajador.ListIndex = Cmbtipotrabajador.NewIndex

Me.dtpFecha1.Value = Format(Now, "dd/mm/yyyy")
Me.dtpFecha2.Value = Format(Now, "dd/mm/yyyy")
End Sub

Private Sub Cargar_Cuenta_Banco()
VBcoPago = fc_CodigoComboBox(CmbBcoPago, 2)
VMoneda = Mid(Trim(Me.CmbMonPago.Text), 2, 3)
cboCuenta.Clear
Sql$ = "usp_pla_PlaBcoCta 0,'" & wcia & "','" & VBcoPago & "','" & Trim(VMoneda) & "','','','','" & wuser & "',5"
If fAbrRst(rs, Sql$) Then
    rs.MoveFirst
    Do While Not rs.EOF
        cboCuenta.AddItem Trim(rs!Sucursal) & Trim(rs!cuentabco)
        cboCuenta.ItemData(cboCuenta.NewIndex) = Val(rs!IdBcoCta)
        rs.MoveNext
    Loop
End If
rs.Close
Set rs = Nothing

cboCuenta.AddItem "<<TODOS>>"
cboCuenta.ItemData(cboCuenta.NewIndex) = "999"
cboCuenta.ListIndex = cboCuenta.NewIndex


End Sub

Private Sub Label8_Click()

End Sub
