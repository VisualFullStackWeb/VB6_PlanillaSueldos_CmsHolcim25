VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmConDerechoHab 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "» Consulta DerechoHabientes «"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11775
   Icon            =   "FrmConDerechoHab.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   11775
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Listado Trabajadores Activos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   360
      Picture         =   "FrmConDerechoHab.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   5160
      Width           =   1440
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00404040&
      Caption         =   "Activo / Cesado"
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
      Height          =   495
      Left            =   0
      TabIndex        =   26
      Top             =   120
      Width           =   3735
      Begin VB.OptionButton OptTodos 
         BackColor       =   &H00404040&
         Caption         =   "Todos"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   2520
         TabIndex        =   29
         Top             =   240
         Width           =   1115
      End
      Begin VB.OptionButton OptEstado 
         BackColor       =   &H00404040&
         Caption         =   "Activo"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton OptInactivo 
         BackColor       =   &H00404040&
         Caption         =   "Inactivo"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   1320
         TabIndex        =   27
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command 
      Caption         =   "Actualizar direcciones"
      Height          =   255
      Left            =   150
      TabIndex        =   24
      Top             =   6405
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton CmdFind 
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
      Left            =   3840
      Picture         =   "FrmConDerechoHab.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   1080
      Width           =   400
   End
   Begin VB.CommandButton CmdRep 
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
      Height          =   1005
      Left            =   600
      Picture         =   "FrmConDerechoHab.frx":0B9E
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   3960
      Width           =   960
   End
   Begin VB.CommandButton CmdClear 
      Caption         =   "Limpiar Filtros"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2280
      TabIndex        =   21
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CheckBox chkFechas 
      BackColor       =   &H00404040&
      Caption         =   "Check1"
      Height          =   255
      Left            =   930
      TabIndex        =   20
      Top             =   1995
      Width           =   255
   End
   Begin VB.CommandButton CmdConsultar 
      Caption         =   "Consultar"
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
      Left            =   10440
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc AdoCon 
      Height          =   330
      Left            =   120
      Top             =   6360
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
   Begin VB.ComboBox CmbCia 
      Enabled         =   0   'False
      Height          =   315
      Left            =   2280
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   720
      Width           =   9375
   End
   Begin VB.ComboBox CmbVinculo 
      Height          =   315
      Left            =   4320
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   360
      Width           =   4935
   End
   Begin VB.Frame Frafechas 
      BackColor       =   &H00404040&
      Caption         =   "Fecha"
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
      ForeColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   60
      TabIndex        =   9
      Top             =   2025
      Width           =   2055
      Begin MSComCtl2.DTPicker FecIni 
         Height          =   315
         Index           =   0
         Left            =   480
         TabIndex        =   12
         Top             =   840
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         CalendarTitleBackColor=   4194304
         CalendarTitleForeColor=   -2147483639
         Format          =   61341697
         CurrentDate     =   39517
         MinDate         =   39448
      End
      Begin VB.OptionButton OptBaja 
         BackColor       =   &H00404040&
         Caption         =   "Baja"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   11
         Top             =   480
         Width           =   1335
      End
      Begin VB.OptionButton OptActivo 
         BackColor       =   &H00404040&
         Caption         =   "Alta"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker FecFin 
         Height          =   315
         Index           =   1
         Left            =   480
         TabIndex        =   13
         Top             =   1200
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         CalendarTitleBackColor=   4194304
         CalendarTitleForeColor=   16777215
         Format          =   61341697
         CurrentDate     =   39517
      End
      Begin VB.Label Label5 
         BackColor       =   &H00404040&
         Caption         =   "al"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label Label5 
         BackColor       =   &H00404040&
         Caption         =   "Del"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   840
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Caption         =   "Consultar Trabajador x"
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
      Height          =   855
      Left            =   60
      TabIndex        =   6
      Top             =   1065
      Width           =   2055
      Begin VB.OptionButton OptCodigo 
         BackColor       =   &H00404040&
         Caption         =   "Código"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   8
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton OptDNI 
         BackColor       =   &H00404040&
         Caption         =   "DNI"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.TextBox TxtNro 
      Height          =   315
      Left            =   2280
      TabIndex        =   0
      Top             =   1080
      Width           =   1575
   End
   Begin TrueOleDBGrid70.TDBGrid GrdCon 
      Bindings        =   "FrmConDerechoHab.frx":0EA8
      Height          =   5295
      Left            =   2280
      TabIndex        =   2
      Top             =   1440
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   9340
      _LayoutType     =   4
      _RowHeight      =   16
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Sexo"
      Columns(0).DataField=   "sexo"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Apellido Paterno Derechohabiente"
      Columns(1).DataField=   "ap_pat"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Apellido Materno Derechohabiente"
      Columns(2).DataField=   "ap_mat"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Nombres Derechohabiente"
      Columns(3).DataField=   "nombres"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Vinculo"
      Columns(4).DataField=   "vinculodh"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Fecha de Nacimiento"
      Columns(5).DataField=   "fec_nacdh"
      Columns(5).NumberFormat=   "Short Date"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Fecha Situacion Activo"
      Columns(6).DataField=   "fecha_alta"
      Columns(6).NumberFormat=   "dd/mm/yyyy"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Codigo Trabajador"
      Columns(7).DataField=   "placod"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "Nombre Trabajador"
      Columns(8).DataField=   "nom_trab"
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "idinterno"
      Columns(9).DataField=   "idinterno"
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   10
      Splits(0)._UserFlags=   0
      Splits(0).Locked=   -1  'True
      Splits(0).RecordSelectorWidth=   688
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).FetchRowStyle=   -1  'True
      Splits(0).DividerColor=   14215660
      Splits(0).FilterBar=   -1  'True
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=10"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=820"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=741"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=8209"
      Splits(0)._ColumnProps(5)=   "Column(0).WrapText=1"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=2725"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2646"
      Splits(0)._ColumnProps(14)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(15)=   "Column(3).Width=2725"
      Splits(0)._ColumnProps(16)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(17)=   "Column(3)._WidthInPix=2646"
      Splits(0)._ColumnProps(18)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(19)=   "Column(4).Width=2196"
      Splits(0)._ColumnProps(20)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(21)=   "Column(4)._WidthInPix=2117"
      Splits(0)._ColumnProps(22)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(23)=   "Column(5).Width=2037"
      Splits(0)._ColumnProps(24)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(25)=   "Column(5)._WidthInPix=1958"
      Splits(0)._ColumnProps(26)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(27)=   "Column(6).Width=2090"
      Splits(0)._ColumnProps(28)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(29)=   "Column(6)._WidthInPix=2011"
      Splits(0)._ColumnProps(30)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(31)=   "Column(7).Width=1429"
      Splits(0)._ColumnProps(32)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(33)=   "Column(7)._WidthInPix=1349"
      Splits(0)._ColumnProps(34)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(35)=   "Column(8).Width=7752"
      Splits(0)._ColumnProps(36)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(37)=   "Column(8)._WidthInPix=7673"
      Splits(0)._ColumnProps(38)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(39)=   "Column(9).Width=2725"
      Splits(0)._ColumnProps(40)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(41)=   "Column(9)._WidthInPix=2646"
      Splits(0)._ColumnProps(42)=   "Column(9).Order=10"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
      Appearance      =   0
      DefColWidth     =   0
      HeadLines       =   2
      FootLines       =   1
      Caption         =   "Detalle DerechoHabientes"
      MultipleLines   =   0
      CellTipsWidth   =   0
      MultiSelect     =   2
      DeadAreaBackColor=   -2147483633
      RowDividerColor =   14215660
      RowSubDividerColor=   14215660
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HFFFFFF&,.fgcolor=&H80000006&"
      _StyleDefs(7)   =   ":id=1,.bold=-1,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=Arial"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37,.bgcolor=&HFF&,.fgcolor=&H80000014&"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=-1,.fontsize=975,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(12)  =   ":id=2,.fontname=Arial"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=-1,.fontsize=975,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(15)  =   ":id=3,.fontname=Arial"
      _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36,.bgcolor=&H808080&"
      _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1,.bold=0,.fontsize=825,.italic=0,.underline=0"
      _StyleDefs(25)  =   ":id=13,.strikethrough=0,.charset=0"
      _StyleDefs(26)  =   ":id=13,.fontname=MS Sans Serif"
      _StyleDefs(27)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(28)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(29)  =   ":id=14,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(30)  =   ":id=14,.fontname=MS Sans Serif"
      _StyleDefs(31)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(32)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(33)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(34)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(35)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(36)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(37)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(38)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(39)  =   "Splits(0).FilterBarStyle:id=24,.parent=12,.bold=-1,.fontsize=825,.italic=0"
      _StyleDefs(40)  =   ":id=24,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(41)  =   ":id=24,.fontname=MS Sans Serif"
      _StyleDefs(42)  =   "Splits(0).Columns(0).Style:id=32,.parent=13,.alignment=2,.valignment=2"
      _StyleDefs(43)  =   ":id=32,.wraptext=-1,.locked=-1"
      _StyleDefs(44)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
      _StyleDefs(45)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
      _StyleDefs(46)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
      _StyleDefs(47)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
      _StyleDefs(48)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
      _StyleDefs(49)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
      _StyleDefs(50)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
      _StyleDefs(51)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
      _StyleDefs(52)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
      _StyleDefs(53)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
      _StyleDefs(54)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
      _StyleDefs(55)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
      _StyleDefs(56)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
      _StyleDefs(57)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
      _StyleDefs(58)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
      _StyleDefs(59)  =   "Splits(0).Columns(4).Style:id=66,.parent=13"
      _StyleDefs(60)  =   "Splits(0).Columns(4).HeadingStyle:id=63,.parent=14"
      _StyleDefs(61)  =   "Splits(0).Columns(4).FooterStyle:id=64,.parent=15"
      _StyleDefs(62)  =   "Splits(0).Columns(4).EditorStyle:id=65,.parent=17"
      _StyleDefs(63)  =   "Splits(0).Columns(5).Style:id=54,.parent=13"
      _StyleDefs(64)  =   "Splits(0).Columns(5).HeadingStyle:id=51,.parent=14"
      _StyleDefs(65)  =   "Splits(0).Columns(5).FooterStyle:id=52,.parent=15"
      _StyleDefs(66)  =   "Splits(0).Columns(5).EditorStyle:id=53,.parent=17"
      _StyleDefs(67)  =   "Splits(0).Columns(6).Style:id=58,.parent=13"
      _StyleDefs(68)  =   "Splits(0).Columns(6).HeadingStyle:id=55,.parent=14"
      _StyleDefs(69)  =   "Splits(0).Columns(6).FooterStyle:id=56,.parent=15"
      _StyleDefs(70)  =   "Splits(0).Columns(6).EditorStyle:id=57,.parent=17"
      _StyleDefs(71)  =   "Splits(0).Columns(7).Style:id=62,.parent=13"
      _StyleDefs(72)  =   "Splits(0).Columns(7).HeadingStyle:id=59,.parent=14"
      _StyleDefs(73)  =   "Splits(0).Columns(7).FooterStyle:id=60,.parent=15"
      _StyleDefs(74)  =   "Splits(0).Columns(7).EditorStyle:id=61,.parent=17"
      _StyleDefs(75)  =   "Splits(0).Columns(8).Style:id=70,.parent=13"
      _StyleDefs(76)  =   "Splits(0).Columns(8).HeadingStyle:id=67,.parent=14"
      _StyleDefs(77)  =   "Splits(0).Columns(8).FooterStyle:id=68,.parent=15"
      _StyleDefs(78)  =   "Splits(0).Columns(8).EditorStyle:id=69,.parent=17"
      _StyleDefs(79)  =   "Splits(0).Columns(9).Style:id=74,.parent=13"
      _StyleDefs(80)  =   "Splits(0).Columns(9).HeadingStyle:id=71,.parent=14"
      _StyleDefs(81)  =   "Splits(0).Columns(9).FooterStyle:id=72,.parent=15"
      _StyleDefs(82)  =   "Splits(0).Columns(9).EditorStyle:id=73,.parent=17"
      _StyleDefs(83)  =   "Named:id=33:Normal"
      _StyleDefs(84)  =   ":id=33,.parent=0"
      _StyleDefs(85)  =   "Named:id=34:Heading"
      _StyleDefs(86)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(87)  =   ":id=34,.wraptext=-1"
      _StyleDefs(88)  =   "Named:id=35:Footing"
      _StyleDefs(89)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(90)  =   "Named:id=36:Selected"
      _StyleDefs(91)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(92)  =   "Named:id=37:Caption"
      _StyleDefs(93)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(94)  =   "Named:id=38:HighlightRow"
      _StyleDefs(95)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(96)  =   "Named:id=39:EvenRow"
      _StyleDefs(97)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(98)  =   "Named:id=40:OddRow"
      _StyleDefs(99)  =   ":id=40,.parent=33"
      _StyleDefs(100) =   "Named:id=41:RecordSelector"
      _StyleDefs(101) =   ":id=41,.parent=34"
      _StyleDefs(102) =   "Named:id=42:FilterBar"
      _StyleDefs(103) =   ":id=42,.parent=33"
   End
   Begin VB.Label LblPlacod 
      Height          =   255
      Left            =   10440
      TabIndex        =   25
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   2
      Left            =   720
      TabIndex        =   18
      Top             =   720
      Width           =   825
   End
   Begin VB.Label Label4 
      BackColor       =   &H00404040&
      Caption         =   "Filtros"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   3840
      TabIndex        =   16
      Top             =   390
      Width           =   1815
   End
   Begin VB.Label LblNomTrab 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   4320
      TabIndex        =   5
      Top             =   1080
      Width           =   6015
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   6165
      Left            =   0
      TabIndex        =   4
      Top             =   615
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00404040&
      Caption         =   "Consulta DerechoHabientes"
      DragMode        =   1  'Automatic
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   720
      Index           =   0
      Left            =   45
      TabIndex        =   3
      Top             =   0
      Width           =   11685
   End
End
Attribute VB_Name = "FrmConDerechoHab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Col As TrueOleDBGrid70.Column
Dim Cols As TrueOleDBGrid70.Columns
Dim Rq As ADODB.Recordset
Dim XEstado As Integer

Private Sub chkFechas_Click()
If chkFechas.Value = 1 Then
    Frafechas.Enabled = True
    Me.FecIni(0).SetFocus
Else
    Frafechas.Enabled = False
End If
End Sub

Private Sub Cmbcia_Click()
Procesar
End Sub

Private Sub CmbVinculo_KeyDown(KeyCode As Integer, Shift As Integer)
If vbKeyDelete = KeyCode Then CmbVinculo.ListIndex = -1
'Me.Procesar
End Sub

Private Sub CmbVinculo_LostFocus()
Me.Procesar
End Sub

Private Sub CmdClear_Click()
    For Each Col In GrdCon.Columns
        Col.FilterText = ""
    Next Col
    AdoCon.Recordset.Filter = adFilterNone
    GrdCon.Refresh
End Sub

Private Sub CmdConsultar_Click()
Me.Procesar
End Sub

Private Sub CmdFind_Click()
If OptCodigo(1).Value = False Then
    MsgBox "Elija tipo de consulta por Codigo de Trabajador", vbExclamation
    Exit Sub
End If
Unload Frmgrdpla
Load Frmgrdpla
Frmgrdpla.Show vbModal

End Sub

Private Sub CmdRep_Click()
Reporte
End Sub

Private Sub Command_Click()
Dim Sql As String
Dim NroTrans As Integer
NroTrans = 0
On Error GoTo ErrorTrans

Sql = "select placod,tvia,nomvia,nrokmmza,intdptolote,tzona,nomzona,ubigeo,referencia"
Sql = Sql & " ,dbo.fc_nombre_ubigeo_sunat(ubigeo) as nom_ubigeo"
Sql = Sql & "  from planillas  where status<>'*'"
Dim Rq As ADODB.Recordset
If fAbrRst(Rq, Sql) Then
    
    cn.BeginTrans
    NroTrans = 1
    Do While Not Rq.EOF
        Dim sDato As String
        sDato = ""
        
        If Left(Trim(Rq!nrokmmza), 3) = "MZ." Or Left(Trim(Rq!nrokmmza), 3) = "MZ-" Then
            sDato = "nro_manzana1='" & Trim(Mid(Trim(Rq!nrokmmza), 4, 6) & "'")
        ElseIf Left(Trim(Rq!nrokmmza), 2) = "M." Then
            sDato = "nro_manzana1='" & Trim(Mid(Trim(Rq!nrokmmza), 2, 4) & "'")
        ElseIf Left(Trim(Rq!nrokmmza), 2) = "Nº" Then
            sDato = "nro_via1='" & Trim(Mid(Trim(Rq!nrokmmza), 3, 4) & "'")
        ElseIf Left(Trim(Rq!nrokmmza), 4) = "LOTE" Then
            sDato = "nro_lote1='" & Trim(Mid(Trim(Rq!nrokmmza), 5, 4) & "'")
        ElseIf Left(Trim(Rq!nrokmmza), 3) = "KM." Then
            sDato = "nro_kilometro1='" & Trim(Mid(Trim(Rq!nrokmmza), 4, 4) & "'")
        Else
            If Len(Trim(Rq!nrokmmza)) > 4 Then
                sDato = "nom_via= RTRIM(nom_via)+ ' " & Trim(Rq!nrokmmza) & "'"
            Else
                sDato = "nro_via1='" & Trim(Rq!nrokmmza) & "'"
            End If
        End If
        
        Sql = " Update pladerechohab"
        Sql = Sql & " set " & sDato
        Sql = Sql & " where plAcod='" & Rq!PlaCod & "' AND STATUS<>'*'"
        cn.Execute Sql, 64
                
        '/**/
        
        sDato = ""
        
        If Left(Trim(Rq!intdptolote), 2) = "LT" Then
            sDato = "nro_lote1='" & Trim(Mid(Trim(Rq!intdptolote), 3, 4) & "'")
        ElseIf Left(Trim(Rq!intdptolote), 3) = "LT." Or Left(Trim(Rq!intdptolote), 3) = "LT-" Then
            sDato = "nro_lote1='" & Trim(Mid(Trim(Rq!intdptolote), 4, 4) & "'")
        ElseIf Left(Trim(Rq!intdptolote), 4) = "LOTE" Then
            sDato = "nro_lote1='" & Trim(Mid(Trim(Rq!intdptolote), 5, 4) & "'")
                   
        Else
            If Len(Trim(Rq!intdptolote)) > 4 Then
                sDato = "nom_via= RTRIM(nom_via)+ ' " & Trim(Rq!intdptolote) & "'"
            Else
                sDato = "nro_lote1='" & Trim(Rq!intdptolote) & "'"
            End If
        End If
        Sql = " Update pladerechohab"
        Sql = Sql & " set " & sDato
        Sql = Sql & " where plAcod='" & Rq!PlaCod & "' AND STATUS<>'*'"
        cn.Execute Sql, 64
        
        Rq.MoveNext
    Loop
    
    cn.CommitTrans
    MsgBox "ACTUALIZACION FINALIZADA ", vbExclamation, Me.Caption
End If
Rq.Close
Set Rq = Nothing

Exit Sub
ErrorTrans:
If NroTrans = 1 Then
    cn.RollbackTrans
End If
MsgBox Err.Description, vbCritical, Me.Caption
End Sub

Private Sub Command1_Click()
Reportes2
End Sub

Private Sub Form_Activate()
Me.Procesar
End Sub

Private Sub Form_Load()
Me.Top = 0: Me.Left = 0: Me.Width = 11865: Me.Height = 7215
Call rCarCbo(Cmbcia, Carga_Cia, "C", "00")
Call rUbiIndCmbBox(Cmbcia, wcia, "00")
Call fc_Descrip_Maestros2("01071", "", CmbVinculo, True)
GrdCon.CaptionStyle.BackColor = &H80000001
FecIni(0).Value = Date
FecFin(1).Value = Date

End Sub

Public Sub Procesar()
Dim xDni As String
Dim xCodigo As String
xDni = "*": xCodigo = "*"
If Trim(TxtNro.Text) <> "" Then
    If OptDNI(0).Value = True Then xDni = Trim(TxtNro.Text)
    If OptCodigo(1).Value = True Then xCodigo = Trim(TxtNro.Text)
End If
Dim xCodVinculo As String
xCodVinculo = "*"
If Me.CmbVinculo.ListIndex > -1 Then
    xCodVinculo = fc_CodigoComboBox(Me.CmbVinculo, 2)
End If
Dim xSituacion As String
xSituacion = "null"

Dim xFecIni As String
Dim xFecFin As String
If Me.chkFechas.Value = 0 Then
    xFecIni = "null"
    xFecFin = "null"
Else
    If OptActivo(0).Value = True Then xSituacion = "1"
    If OptBaja(1).Value = True Then xSituacion = "0"
    xFecIni = "'" & Format(Me.FecIni(0).Value, "MM/DD/YYYY") & " 12:00:00 AM" & "'"
    xFecFin = "'" & Format(Me.FecFin(1).Value, "MM/DD/YYYY") & " 11:59:59 PM" & "'"
End If


'/***Dim XEstado As Integer
XEstado = 0
If OptEstado(0).Value = True Then XEstado = 1
If OptInactivo(0).Value = True Then XEstado = 2
If optTodos(0).Value = True Then XEstado = 3

Dim Sql As String
Sql = "usp_pla_consultar_derechohabientes '" & wcia & "','" & xCodigo & "','" & xDni & "','" & xCodVinculo & "'," & xSituacion & "," & xFecIni & "," & xFecFin & "," & XEstado & ""

Screen.MousePointer = 11
If fAbrRst(Rq, Sql) Then
    Set AdoCon.Recordset = Rq
Else
    Set AdoCon.Recordset = cn.Execute("select getdate()", 64)
End If
'AdoCon.Refresh
Screen.MousePointer = 0


End Sub

Private Sub GrdCon_DblClick()
Unload aFrmDH
aFrmDH.LblIdInternodh.Caption = Trim(GrdCon.Columns(9))
aFrmDH.TxtCodTrab.Text = AdoCon.Recordset!PlaCod
aFrmDH.LblNomTrab.Caption = Trim(AdoCon.Recordset!ap_pat & "") + " " + Trim(AdoCon.Recordset!ap_mat & "") & " " & Trim(AdoCon.Recordset!nombres & "")
aFrmDH.CargaDerechohabiente
End Sub

Private Sub GrdCon_FilterChange()
'Gets called when an action is performed on the filter bar

On Error GoTo ErrHandler

Set Cols = GrdCon.Columns

Dim c As Integer

c = GrdCon.Col

GrdCon.HoldFields

AdoCon.Recordset.Filter = getFilter()

GrdCon.Col = c

GrdCon.EditActive = True

Exit Sub

 

ErrHandler:

    MsgBox Err.Source & ":" & vbCrLf & Err.Description

    Call CmdClear_Click
    
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

Private Sub OptEstado_Click(Index As Integer)
Me.Procesar
End Sub

Private Sub OptInactivo_Click(Index As Integer)
Me.Procesar
End Sub

Private Sub OptTodos_Click(Index As Integer)
Me.Procesar
End Sub

Private Sub Txtnro_Change()
LblNomTrab.Caption = ""
If OptCodigo(1).Value Then
    If Len(Trim(TxtNro.Text)) >= 5 Then
        TxtNro_LostFocus
        CmdConsultar_Click
    End If
End If

If OptDNI(0).Value Then
    If Len(Trim(TxtNro.Text)) = 8 Then
        TxtNro_LostFocus
        CmdConsultar_Click
    End If
End If

End Sub

Private Sub txtnro_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Sendkeys "{TAB}": Me.Procesar
End Sub

Private Sub TxtNro_LostFocus()
If Trim(TxtNro.Text) = "" Then Exit Sub
Dim Sql As String
Dim xDni As String
Dim xCodigo As String
xDni = "": xCodigo = ""
If Trim(TxtNro.Text) <> "" Then
    If OptDNI(0).Value = True Then xCodigo = " and nro_doc='" & Trim(TxtNro.Text) & "' "
    If OptCodigo(1).Value = True Then xCodigo = " and placod='" & Trim(TxtNro.Text) & "' "
End If
Sql = "select dbo.fc_Razsoc_Trabajador(placod,cia) as nom_trab,fcese,PLACOD "
Sql = Sql & " from planillas where cia='" & wcia & "'" & xCodigo & " AND status<>'*'"
Dim Rq As ADODB.Recordset
Screen.MousePointer = 11
If fAbrRst(Rq, Sql) Then
    If Not IsNull(Rq!fcese) Then
        MsgBox "El Trabajador ya fue Cesado", vbExclamation, Me.Caption
        GoTo Salir:
    End If
    LblNomTrab.Caption = Rq(0)
    LblPlacod.Caption = Trim(Rq!PlaCod & "")
Else
    LblNomTrab.Caption = ""
    MsgBox "No existe codigo,verifique", vbExclamation, Me.Caption
End If
Salir:
Rq.Close
Set Rq = Nothing
Screen.MousePointer = 0

End Sub

Public Sub Reporte()
'On Error GoTo MsgError:

Dim xDni As String
Dim xCodigo As String
xDni = "*": xCodigo = "*"
If Trim(TxtNro.Text) <> "" Then
    If OptDNI(0).Value = True Then xDni = Trim(TxtNro.Text)
    If OptCodigo(1).Value = True Then xCodigo = Trim(TxtNro.Text)
End If
Dim xCodVinculo As String
xCodVinculo = "*"
If Me.CmbVinculo.ListIndex > -1 Then
    xCodVinculo = fc_CodigoComboBox(Me.CmbVinculo, 2)
End If
Dim xSituacion As String
xSituacion = "null"

Dim xFecIni As String
Dim xFecFin As String
If Me.chkFechas.Value = 0 Then
    xFecIni = "null"
    xFecFin = "null"
Else
    If OptActivo(0).Value = True Then xSituacion = "1"
    If OptBaja(1).Value = True Then xSituacion = "0"
    xFecIni = "'" & Format(Me.FecIni(0).Value, "MM/DD/YYYY") & " 12:00:00 AM" & "'"
    xFecFin = "'" & Format(Me.FecFin(1).Value, "MM/DD/YYYY") & " 11:59:59 PM" & "'"
End If

Screen.MousePointer = 11
Dim l_oReporte As ClsExcel
Set l_oReporte = New ClsExcel
If l_oReporte.p_AbrirExcel() = False Then
    MsgBox "No se pudo abrir el Excel", vbCritical, Me.Caption
    Exit Sub
End If
l_oReporte.NombreCia = Cmbcia.Text


XEstado = 0
If OptEstado(0).Value = True Then XEstado = 1
If OptInactivo(0).Value = True Then XEstado = 2
If optTodos(0).Value = True Then XEstado = 3

Call l_oReporte.Reporte_DerechoHabientes(wcia, xDni, xCodigo, xCodVinculo, xSituacion, xFecIni, xFecFin, XEstado)
Call l_oReporte.p_CerrarExcel
Set l_oReporte = Nothing
Termina:
Screen.MousePointer = 0
'Rq.Close
'Set Rq = Nothing

    
Exit Sub


End Sub

Public Sub Reportes2()

Screen.MousePointer = 11
Dim l_oReporte As ClsExcel
Set l_oReporte = New ClsExcel
If l_oReporte.p_AbrirExcel() = False Then
    MsgBox "No se pudo abrir el Excel", vbCritical, Me.Caption
    Exit Sub
End If
l_oReporte.NombreCia = Cmbcia.Text

XEstado = 0
If OptEstado(0).Value = True Then XEstado = 1
If OptInactivo(0).Value = True Then XEstado = 2
If optTodos(0).Value = True Then XEstado = 1

Call l_oReporte.Reporte_Trabajadores(wcia, XEstado)
Call l_oReporte.p_CerrarExcel
Set l_oReporte = Nothing
Termina:
Screen.MousePointer = 0
Exit Sub



MsgError:
    'ErrorLog "", False, "", Err.Number & Space(1) & Err.Description
    Screen.MousePointer = vbDefault
End Sub
