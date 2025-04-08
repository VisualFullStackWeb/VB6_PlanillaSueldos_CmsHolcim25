VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmExamenes 
   Caption         =   "Registrar Examenes"
   ClientHeight    =   8085
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11400
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8085
   ScaleWidth      =   11400
   Begin MSComCtl2.DTPicker dtpFecha 
      Height          =   375
      Left            =   1200
      TabIndex        =   30
      Top             =   1560
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   16646145
      CurrentDate     =   43671
   End
   Begin VB.ListBox LstAreas 
      Height          =   2010
      Left            =   11640
      TabIndex        =   28
      Top             =   600
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.TextBox txtRucProveedor 
      Height          =   315
      Left            =   2160
      MaxLength       =   15
      TabIndex        =   27
      Text            =   "20518132947"
      Top             =   120
      Width           =   3135
   End
   Begin VB.TextBox txtProveedor 
      BackColor       =   &H80000016&
      Height          =   315
      Left            =   1200
      MaxLength       =   150
      TabIndex        =   25
      Text            =   "SERVICIOS MEDICOS EL TREBOL S.A.C."
      Top             =   480
      Width           =   4095
   End
   Begin VB.TextBox txtCodProv 
      Height          =   315
      Left            =   1200
      MaxLength       =   6
      TabIndex        =   24
      Text            =   "PS1418"
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox txtRangoFin 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   2640
      TabIndex        =   22
      Text            =   "AZ100"
      Top             =   2370
      Width           =   855
   End
   Begin VB.TextBox txtTotal 
      BackColor       =   &H80000016&
      Height          =   315
      Left            =   4200
      Locked          =   -1  'True
      TabIndex        =   20
      Text            =   "0"
      Top             =   7560
      Width           =   1215
   End
   Begin VB.TextBox txtIGV 
      BackColor       =   &H80000016&
      Height          =   315
      Left            =   4200
      Locked          =   -1  'True
      TabIndex        =   19
      Text            =   "0"
      Top             =   7200
      Width           =   1215
   End
   Begin VB.TextBox txtMontoTotal 
      BackColor       =   &H80000016&
      Height          =   315
      Left            =   4200
      Locked          =   -1  'True
      TabIndex        =   17
      Text            =   "0"
      Top             =   6840
      Width           =   1215
   End
   Begin VB.TextBox txtNroPersonas 
      Height          =   315
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   16
      Text            =   "0"
      Top             =   7200
      Width           =   1215
   End
   Begin VB.TextBox txtRangoInicio 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   1200
      TabIndex        =   13
      Text            =   "A12"
      Top             =   2370
      Width           =   855
   End
   Begin VB.TextBox txtNroHojas 
      Height          =   315
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   10
      Text            =   "0"
      Top             =   6840
      Width           =   1215
   End
   Begin TrueOleDBGrid70.TDBGrid GrdPlanilla 
      Height          =   7815
      Left            =   5520
      TabIndex        =   8
      Top             =   120
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   13785
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Código"
      Columns(0).DataField=   "placod"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Apellidos y Nombres"
      Columns(1).DataField=   "nombres"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Nro.Doc"
      Columns(2).DataField=   "dni"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "CodArea"
      Columns(3).DataField=   "codarea"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   129
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Area"
      Columns(4).DataField=   "area"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Costo"
      Columns(5).DataField=   "precio"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   6
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   688
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=6"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1773"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1693"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=8196"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=6165"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=6085"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=8196"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(1)._MinWidth=267"
      Splits(0)._ColumnProps(12)=   "Column(2).Width=2117"
      Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=2037"
      Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(16)=   "Column(3).Width=2725"
      Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=2646"
      Splits(0)._ColumnProps(19)=   "Column(3).Visible=0"
      Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(21)=   "Column(4).Width=4419"
      Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=4339"
      Splits(0)._ColumnProps(24)=   "Column(4).Button=1"
      Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(26)=   "Column(5).Width=2117"
      Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=2037"
      Splits(0)._ColumnProps(29)=   "Column(5).Order=6"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   14215660
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=188,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
      _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.locked=-1"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.locked=-1"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=50,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=54,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=51,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=52,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=53,.parent=17"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=58,.parent=13"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=55,.parent=14"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=56,.parent=15"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=57,.parent=17"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=46,.parent=13"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=43,.parent=14"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=44,.parent=15"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=45,.parent=17"
      _StyleDefs(60)  =   "Named:id=33:Normal"
      _StyleDefs(61)  =   ":id=33,.parent=0"
      _StyleDefs(62)  =   "Named:id=34:Heading"
      _StyleDefs(63)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(64)  =   ":id=34,.wraptext=-1"
      _StyleDefs(65)  =   "Named:id=35:Footing"
      _StyleDefs(66)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(67)  =   "Named:id=36:Selected"
      _StyleDefs(68)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(69)  =   "Named:id=37:Caption"
      _StyleDefs(70)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(71)  =   "Named:id=38:HighlightRow"
      _StyleDefs(72)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(73)  =   "Named:id=39:EvenRow"
      _StyleDefs(74)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(75)  =   "Named:id=40:OddRow"
      _StyleDefs(76)  =   ":id=40,.parent=33"
      _StyleDefs(77)  =   "Named:id=41:RecordSelector"
      _StyleDefs(78)  =   ":id=41,.parent=34"
      _StyleDefs(79)  =   "Named:id=42:FilterBar"
      _StyleDefs(80)  =   ":id=42,.parent=33"
   End
   Begin TrueOleDBGrid70.TDBGrid GrdHojas 
      Height          =   3975
      Left            =   120
      TabIndex        =   7
      Top             =   2760
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   7011
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   4
      Columns(0)._MaxComboItems=   5
      Columns(0).DataField=   "sel"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Pestaña"
      Columns(1).DataField=   "hoja"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   2
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   688
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=2"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=529"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=450"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=7938"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=7858"
      Splits(0)._ColumnProps(8)=   "Column(1)._ColStyle=8196"
      Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   14215660
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
      _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.locked=-1"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(44)  =   "Named:id=33:Normal"
      _StyleDefs(45)  =   ":id=33,.parent=0"
      _StyleDefs(46)  =   "Named:id=34:Heading"
      _StyleDefs(47)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(48)  =   ":id=34,.wraptext=-1"
      _StyleDefs(49)  =   "Named:id=35:Footing"
      _StyleDefs(50)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(51)  =   "Named:id=36:Selected"
      _StyleDefs(52)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(53)  =   "Named:id=37:Caption"
      _StyleDefs(54)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(55)  =   "Named:id=38:HighlightRow"
      _StyleDefs(56)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(57)  =   "Named:id=39:EvenRow"
      _StyleDefs(58)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(59)  =   "Named:id=40:OddRow"
      _StyleDefs(60)  =   ":id=40,.parent=33"
      _StyleDefs(61)  =   "Named:id=41:RecordSelector"
      _StyleDefs(62)  =   ":id=41,.parent=34"
      _StyleDefs(63)  =   "Named:id=42:FilterBar"
      _StyleDefs(64)  =   ":id=42,.parent=33"
   End
   Begin VB.CommandButton cmbAbrirArchivo 
      BackColor       =   &H80000013&
      Caption         =   "..."
      Height          =   315
      Left            =   4920
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   6
      Top             =   2040
      Width           =   375
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   7440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtRutaArchivo 
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   2000
      Width           =   3615
   End
   Begin VB.TextBox txtNumDoc 
      Height          =   315
      Left            =   2040
      MaxLength       =   16
      TabIndex        =   4
      Text            =   "0000000000"
      Top             =   1200
      Width           =   3375
   End
   Begin VB.TextBox txtSerie 
      Height          =   315
      Left            =   1200
      MaxLength       =   4
      TabIndex        =   3
      Text            =   "F001"
      Top             =   1200
      Width           =   855
   End
   Begin VB.ComboBox cboTipo 
      Height          =   315
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   840
      Width           =   4215
   End
   Begin VB.Label Label12 
      Caption         =   "Fecha:"
      Height          =   315
      Left            =   240
      TabIndex        =   29
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label11 
      Caption         =   "Proveedor"
      Height          =   255
      Left            =   240
      TabIndex        =   26
      Top             =   165
      Width           =   855
   End
   Begin VB.Label Label10 
      Caption         =   "Fin:"
      Height          =   255
      Left            =   2160
      TabIndex        =   23
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label Label9 
      Caption         =   "Total"
      Height          =   255
      Left            =   2880
      TabIndex        =   21
      Top             =   7560
      Width           =   1095
   End
   Begin VB.Label Label8 
      Caption         =   "IGV"
      Height          =   255
      Left            =   2880
      TabIndex        =   18
      Top             =   7200
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "Nro.Personas:"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   7200
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "Monto Total:"
      Height          =   255
      Left            =   2880
      TabIndex        =   14
      Top             =   6840
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Inicio:"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "Archivo:"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   2040
      Width           =   855
   End
   Begin MSForms.Label Label3 
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   6840
      Width           =   1335
      Caption         =   "Nro de Hojas"
      Size            =   "2355;450"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label2 
      Caption         =   "Nro.Factura"
      Height          =   315
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Tipo"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   855
   End
End
Attribute VB_Name = "FrmExamenes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsdatos As New ADODB.Recordset
Dim rsHojas As New ADODB.Recordset
Dim RUTA As String
Dim editar As Boolean

Dim documento As String

Dim Tecla As Boolean

Private Sub cboTipo_Click()

    documento = fc_CodigoComboBox(cboTipo, 2)
    
    If documento = "01" Then
        'Terapias Fisicas
        txtMontoTotal.BackColor = &HC0FFFF
        txtMontoTotal.Locked = False
        
        txtRangoInicio.Text = "A5"
        txtRangoFin.Text = "F100"

    Else
        'Examenes medicos
        txtMontoTotal.BackColor = &H80000016
        txtMontoTotal.Locked = True
        
        txtRangoInicio.Text = "A12"
        txtRangoFin.Text = "AZ100"
    End If
    
End Sub

Private Sub cmbAbrirArchivo_Click()

    If documento = "00" Then
        MsgBox "Debe seleccionar el tipo de documento", vbExclamation, Me.Caption
        cboTipo.SetFocus
        Exit Sub
    End If
    If Len(Trim(txtSerie.Text)) = 0 Or Len(Trim(txtNumDoc.Text)) = 0 Then
        MsgBox "Ingrese el número de la factura", vbExclamation, Me.Caption
        txtSerie.SetFocus
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    If documento = "01" Then
        If Len(Trim(txtMontoTotal.Text)) = 0 Then
            MsgBox "Ingrese el monto de la factura", vbExclamation, Me.Caption
            txtMontoTotal.SetFocus
            Exit Sub
        End If
        If CCur(Trim(txtMontoTotal.Text)) = 0 Then
            MsgBox "Ingrese el monto de la factura", vbExclamation, Me.Caption
            txtMontoTotal.SetFocus
            Exit Sub
        End If
    End If
    
    CommonDialog1.Filter = "Archivo de Excel (*.xlsx)|*.xlsx"
    CommonDialog1.DefaultExt = "xlsx"
    CommonDialog1.DialogTitle = "Seleccione el archivo"
    CommonDialog1.ShowOpen
    
    txtRutaArchivo.Text = CommonDialog1.FileName
    
    Call ImportarExcel
End Sub

Private Sub Form_Load()
    If Not WindowState = vbNormal Then WindowState = vbNormal
    Me.Top = 0: Me.Left = 0: Me.Width = 15570: Me.Height = 8600
        
    Call CargarListaAreas
    Call CargarTipos
    Call Crear_Rs
    Call Nuevo
    
    dtpFecha.Value = Now
    
End Sub

Sub Crear_Rs()
    If rsdatos.State = 1 Then rsdatos.Close: Set rsdatos = Nothing
    rsdatos.Fields.Append "placod", adVarChar, 8, adFldIsNullable
    rsdatos.Fields.Append "nombres", adVarChar, 250, adFldIsNullable
    rsdatos.Fields.Append "dni", adVarChar, 20, adFldIsNullable
    rsdatos.Fields.Append "codarea", adVarChar, 4, adFldIsNullable
    rsdatos.Fields.Append "area", adVarChar, 150, adFldIsNullable
    rsdatos.Fields.Append "precio", adCurrency, 20, adFldIsNullable
    rsdatos.Open
    Set GrdPlanilla.DataSource = rsdatos
    
    GrdPlanilla.Columns(3).NumberFormat = "####.0000"
        
    If rsHojas.State = 1 Then rsHojas.Close: Set rsHojas = Nothing
    rsHojas.Fields.Append "sel", adBoolean, , adFldIsNullable
    rsHojas.Fields.Append "hoja", adVarChar, 250, adFldIsNullable
    rsHojas.Open
           
    Set GrdHojas.DataSource = rsHojas
    
End Sub

Sub CargarListaAreas()
On Error GoTo Err
    Dim rcs As New ADODB.Recordset
    
    Sql = "select * from pla_areas where status<>'*'"
    Dim c As Integer
    If (fAbrRst(rcs, Sql)) Then
        rcs.MoveFirst
        Do While Not rcs.EOF
            LstAreas.AddItem rcs!DPTO & Space(100) & rcs!cod_area
            LstAreas.ItemData(LstAreas.NewIndex) = Format(rcs!cod_area, "000")
            c = c + 1
            If c > 200 Then Exit Sub
            
            rcs.MoveNext
        Loop
    End If
    
    If rcs.State = adStateOpen Then
        rcs.Close
    End If
    Exit Sub
Err:
    MsgBox Err.Number & "-" & Err.Description, vbCritical, Me.Caption
End Sub

Sub CargarTipos()

    cboTipo.AddItem "-- Seleccione --"
    cboTipo.ItemData(cboTipo.NewIndex) = "00"
    
    cboTipo.AddItem "Terapia Fisica"
    cboTipo.ItemData(cboTipo.NewIndex) = "01"
    
    cboTipo.AddItem "Exámen Médico"
    cboTipo.ItemData(cboTipo.NewIndex) = "02"
    
End Sub

Sub ImportarExcel()

    On Error GoTo Err

    RUTA = Trim(txtRutaArchivo.Text)

    If Len(RUTA) = 0 Then Exit Sub

    LimpiarRsT rsHojas, GrdHojas

    Dim xlApp2 As Excel.Application
    Dim xlApp1 As Excel.Application
    Dim xLibro As Excel.Workbook
    
    Set xlApp1 = GetObject("", "Excel.Application")
    If xlApp1 Is Nothing Then
        'Si excel no esta corriendo, creamos una nueva instancia.
        Set xlApp1 = CreateObject("Excel.Application")
    End If
    
    'Variable de tipo Aplicación de Excel
    
    Set xlApp2 = xlApp1.Application
    
    'Una variable de tipo Libro de Excel
    
    Dim Col As Integer, Fila As Integer
    
    Set xLibro = xlApp2.Workbooks.Open(RUTA)
    
    Dim Inicio As Integer, fin As Integer
    Inicio = 1
    fin = xLibro.Worksheets.count
    
    For Inicio = 1 To fin
        rsHojas.AddNew
        rsHojas!SEL = False
        rsHojas!hoja = xLibro.Worksheets(Inicio).Name
        rsHojas.Update
    Next

    rsHojas.Filter = adFilterNone
    rsHojas.MoveFirst
    
    GrdHojas.Refresh
    
    txtNroHojas.Text = rsHojas.RecordCount
    
    xlApp2.Visible = False
    
    xLibro.Close
    xlApp2.Quit
    
    If Not xlApp2 Is Nothing Then Set xlApp2 = Nothing
    If Not xlApp1 Is Nothing Then Set xlApp1 = Nothing
    If Not xLibro Is Nothing Then Set xlBook = Nothing
    
Exit Sub

Err:

    MsgBox Err.Number & "-" & Err.Description, vbCritical, Me.Caption

End Sub

Sub LimpiarRsT(ByRef pRs As ADODB.Recordset, ByRef pDgrd As Object)
    pDgrd.Refresh
    If pRs.State = 1 Then
        If pRs.RecordCount > 0 Then
            pRs.MoveFirst
            Do While Not pRs.EOF
                pRs.Delete
                If Not pRs.EOF Then pRs.MoveNext
            Loop
        End If
    End If
    pDgrd.Refresh
End Sub

Private Sub GrdHojas_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    If ColIndex = 0 Then
        
    End If
End Sub

Private Sub GrdHojas_AfterColUpdate(ByVal ColIndex As Integer)
    
    On Error GoTo Err
    
    Dim hoja As String, rango As String
    Dim Monto As Currency, nroTotal As Integer, punit As Currency
    
    Dim pDni As String, pNombres As String, pPrecio As Currency
    
    ' Monto = CCur(txtMontoTotal.Text)
    Monto = 0
    nroTotal = 0
    punit = 0
    
    Dim conexion As ADODB.Connection, rs As ADODB.Recordset
  
    Set conexion = New ADODB.Connection
    If rsHojas.RecordCount = 0 Then Exit Sub
    
    rsHojas.MoveFirst
'    documento = fc_CodigoComboBox(cboTipo, 2)
    
'    conexion.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
'                  "Data Source=" & RUTA & _
'                  ";Extended Properties=""Excel 8.0;HDR=Yes;"""
    If UCase(Right(RUTA, 5)) = ".XLSX" Then
        'XLSX
        conexion.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                    "Data Source=" & RUTA & ";" & _
                    "Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1;"""
    Else
        'XLS
        conexion.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                    "Data Source=" & RUTA & ";" & _
                    "Extended Properties=""Excel 8.0;HDR=YES;IMEX=1;"""
    End If
    
    If ColIndex = 0 Then
        
        LimpiarRsT rsdatos, GrdPlanilla
        
        If rsHojas.RecordCount > 0 Then
            Do While Not rsHojas.EOF
                If rsHojas!SEL Then
                
                    hoja = rsHojas!hoja
                    rango = Trim(txtRangoInicio.Text) & ":" & Trim(txtRangoFin.Text)
                    Set rs = New ADODB.Recordset
            
                    With rs
                        .CursorLocation = adUseClient
                        .CursorType = adOpenStatic
                        .LockType = adLockOptimistic
                    End With
                    
                    rs.Open "SELECT * FROM [" & hoja & "$" & rango & "]", conexion, , , adCmdText
                    
                    If rs.RecordCount > 0 Then
                    
                        Do While Not rs.EOF
                       
                            pDni = "": pNombres = "": pPrecio = 0#
                            
                            If documento = "01" Then
                                pPrecio = Format((Monto / rs.RecordCount), "###,###,##0.00")
                                pNombres = Trim(rs!Trabajador & "")
                            End If
                            If documento = "02" Then
                                If Len(Trim(rs!costo & "")) = 0 Then
                                    pPrecio = 0
                                Else
                                    pPrecio = Format(CCur(rs!costo), "###,###,##0.00")
                                End If
                                pNombres = Trim(rs(1) & "")
                            End If
                            
                            pDni = Format(Trim(rs(2) & ""), "00000000")
                            
                            If ValidarDatosEmpleado(pDni, pNombres, pPrecio) Then
                                nroTotal = nroTotal + 1
                                Monto = Monto + pPrecio
                            End If
                            
                            rs.MoveNext
                        Loop
                    End If
                                                     
                    If rs.State = adStateOpen Then
                        rs.Close
                    End If
                                                     
                    Set rs = Nothing
                End If
                rsHojas.MoveNext
            Loop
        End If
    End If
    

    If documento = "01" And nroTotal > 0 Then
        Monto = CCur(txtMontoTotal.Text)
        punit = Monto / nroTotal
        punit = Format(punit, "###,###,##0.00")

        Monto = 0

        rsdatos.MoveFirst
        Do While Not rsdatos.EOF
            rsdatos!precio = punit
            rsdatos.Update

            Monto = Monto + punit

            rsdatos.MoveNext
        Loop
    End If

    
    Call CalcularTotales(Monto, nroTotal)
    
    If conexion.State = 1 Then
        conexion.Close
    End If
         
    Set conexion = Nothing
         
Exit Sub
       
Err:
    MsgBox Err.Number & "-" & Err.Description, vbCritical, Me.Caption
End Sub

Private Function ValidarDatosEmpleado(ByVal pNroDoc As String, ByVal pNombres As String, ByVal pPrecio As Currency) As Boolean
On Error GoTo Err
    Dim rsEmpleado As New ADODB.Recordset
    
    Sql = "exec rrhh_validar_detalle_documento "
    Sql = Sql & "@codcia='" & wcia & "',"
    Sql = Sql & "@dni='" & pNroDoc & "'"
    
    If Len(pNroDoc) = 0 And Len(pNombres) = 0 Then
        'no es una fila valida , puede ser un subtotal o algo
        ValidarDatosEmpleado = False
    Else
        If (fAbrRst(rsEmpleado, Sql)) Then
            rsEmpleado.MoveFirst
            
            rsdatos.AddNew
            rsdatos!PlaCod = Trim(rsEmpleado!PlaCod & "")
            rsdatos!nombres = Trim(rsEmpleado!nombres & "")
            rsdatos!DNI = Trim(rsEmpleado!DNI & "")
            rsdatos!codarea = Trim(rsEmpleado!cod_area & "")
            rsdatos!Area = Trim(rsEmpleado!Area & "")
            rsdatos!precio = Format(pPrecio, "###,###,##0.00")
            rsdatos.Update
        Else
            rsdatos.AddNew
            rsdatos!PlaCod = ""
            rsdatos!nombres = pNombres
            rsdatos!DNI = pNroDoc
            rsdatos!codarea = ""
            rsdatos!Area = ""
            rsdatos!precio = Format(pPrecio, "###,###,##0.00")
            rsdatos.Update
        End If
        
        ValidarDatosEmpleado = True
    End If
        
    If rsEmpleado.State = adStateOpen Then
        rsEmpleado.Close
    End If
    Exit Function
Err:
    MsgBox Err.Number & "-" & Err.Description, vbCritical, Me.Caption
End Function


Private Sub CalcularTotales(pMonto As Currency, pNroPersonas As Integer)
    txtMontoTotal.Text = Format(pMonto, "###,###,##0.00")
    txtIGV.Text = Format(pMonto * 0.18, "###,###,##0.00")
    txtTotal.Text = Format(pMonto + (pMonto * 0.18), "###,###,##0.00")
    txtNroPersonas.Text = pNroPersonas
End Sub

Public Sub CargarDatosFactura()

    On Error GoTo Err

    Dim numdoc As String, codprov As String, xdoc As String
    Dim Monto As Currency, contador As Integer
    
    Dim rs As Recordset
    numdoc = Trim(txtSerie.Text) & Trim(txtNumDoc.Text)
    codprov = Trim(txtCodProv.Text)
    
    If Len(numdoc) = 0 Or Len(codprov) = 0 Then Exit Sub
    
    contador = 0
    Monto = 0#
    xdoc = "00"
    
    Set rs = New ADODB.Recordset
    
    LimpiarRsT rsdatos, GrdPlanilla
    
    Sql = "exec rrhh_listar_detalle_documentos "
    Sql = Sql & "@codcia='" & wcia & "',"
    Sql = Sql & "@codprov='" & codprov & "',"
    Sql = Sql & "@numdoc='" & numdoc & "'"
        
    If fAbrRst(rs, Sql) Then
    
        dtpFecha.Value = CDate(rs!fecha)
    
        Do While Not rs.EOF
            rsdatos.AddNew
            rsdatos!PlaCod = rs!PlaCod
            rsdatos!nombres = rs!nombres
            rsdatos!DNI = rs!DNI
            rsdatos!codarea = rs!cod_area
            rsdatos!Area = rs!Area
            rsdatos!precio = rs!precio
            rsdatos.Update
            
            xdoc = rs!documento
            contador = contador + 1
            Monto = Monto + CCur(rs!precio)
            
            rs.MoveNext
        Loop
        
        Call rUbiIndCmbBox(cboTipo, xdoc, "00")
    End If
    
    If contador = 0 Then
        editar = False
    Else
        editar = True
    End If
        
    If rs.State = adStateOpen Then
        rs.Close
    End If
    
    Set rs = Nothing
        
    Call CalcularTotales(Monto, contador)
    
Exit Sub
       
Err:
    MsgBox Err.Number & "-" & Err.Description, vbCritical, Me.Caption
End Sub

Private Sub GrdPlanilla_ButtonClick(ByVal ColIndex As Integer)
On Error GoTo Err
    Dim indexRow As Integer, xtop As Integer, xleft As Integer
    
    indexRow = GrdPlanilla.Row
        
    If ColIndex = 4 Then ' area
        xtop = GrdPlanilla.Top + GrdPlanilla.RowTop(indexRow) + GrdPlanilla.RowHeight
        xleft = GrdPlanilla.Left + GrdPlanilla.Columns(4).Left
        
        If indexRow < 8 Then
            LstAreas.Top = xtop
        Else
            LstAreas.Top = GrdPlanilla.Top + GrdPlanilla.RowTop(indexRow) - LstAreas.Height
        End If
        
        LstAreas.Left = xleft
        LstAreas.Visible = True
        LstAreas.SetFocus
        LstAreas.ZOrder 0
    End If
    
    Exit Sub
Err:
    MsgBox Err.Number & "-" & Err.Description, vbCritical, Me.Caption
End Sub

Private Sub GrdPlanilla_Scroll(Cancel As Integer)
    LstAreas.Visible = False
End Sub

Private Sub LstAreas_Click()
    If LstAreas.ListIndex > -1 And Tecla Then
        LstAreas.Visible = False
    End If
    GrdPlanilla.Columns(4).Value = LstAreas.Text
    GrdPlanilla.Columns(3).Value = fc_CodigoComboBox(LstAreas, 4)
    GrdPlanilla.Update
End Sub

Private Sub LstAreas_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 40 Or KeyCode = 38 Then Tecla = False: Exit Sub
    If KeyCode = 27 Then LstAreas.Visible = False: GrdPlanilla.Col = 1: GrdPlanilla.SetFocus
    If KeyCode = 13 Then Tecla = True: LstAreas_Click
End Sub

Private Sub LstAreas_LostFocus()
    LstAreas.Visible = False
End Sub

Private Sub LstAreas_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Tecla = True
End Sub

Private Sub txtCodProv_Keypress(KeyAscii As Integer)
    If KeyAscii <> 13 Then Return
    Call CargarDatosProveedor(Trim(txtCodProv.Text))
End Sub

Private Sub txtMontoTotal_GotFocus()
    Call ResaltarTexto(txtMontoTotal)
End Sub

Private Sub txtMontoTotal_LostFocus()
'    Dim Monto As Currency, contador As Integer
'    Monto = 0
'    contador = CInt(txtNroPersonas.Text)
'    If IsNumeric(Trim(txtMontoTotal.Text)) Then
'        Monto = CCur(Trim(txtMontoTotal.Text))
'    End If
'    Call CalcularTotales(Monto, contador)
    Call txtMontoTotal_Keypress(13)
End Sub

Private Sub txtMontoTotal_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If documento = "01" Then
            Dim Monto As Currency
            Dim nroTotal As Integer
            
            Monto = CCur(txtMontoTotal.Text)
            nroTotal = rsdatos.RecordCount
            
            If nroTotal > 0 Then
                punit = Monto / nroTotal
                punit = Format(punit, "###,###,##0.00")
        
                Monto = 0
        
                rsdatos.MoveFirst
                Do While Not rsdatos.EOF
                    rsdatos!precio = punit
                    rsdatos.Update
        
                    Monto = Monto + punit
        
                    rsdatos.MoveNext
                Loop
                
                Call CalcularTotales(Monto, nroTotal)
            End If
        
        End If
    End If
End Sub

Private Sub txtRucProveedor_Keypress(KeyAscii As Integer)
    If KeyAscii <> 13 Then Return
    Call CargarDatosProveedor(Trim(txtRucProveedor.Text))
End Sub

Private Sub CargarDatosProveedor(Codigo As String)
On Error GoTo Err:
    Dim rs As Recordset
    Set rs = New ADODB.Recordset
    
    txtCodProv.Text = ""
    txtRucProveedor.Text = ""
    txtProveedor.Text = ""
    
    If Len(Codigo) = 6 Then
        Sql = "SP_PROVEEDORES_con_codigo '" & wcia & "','" & Codigo & "'"
    ElseIf Len(Codigo) = 13 Then
        Sql = "SP_PROVEEDORES_con_RUC '" & wcia & "','" & Codigo & "'"
    Else
        Exit Sub
    End If
    
    If fAbrRst(rs, Sql) Then
        txtCodProv.Text = rs!cod_prov & ""
        txtRucProveedor.Text = rs!RUC
        txtProveedor.Text = Trim(rs!razsoc & "" & Trim(rs!ap_pat) & Space(1) & Trim(rs!ap_mat) & Space(1) & Trim(rs!pri_nom))
                
        Call CargarDatosFactura
        
    End If
    
    If rs.State = adStateOpen Then
        rs.Close
    End If
    
    Set rs = Nothing
    
Exit Sub
Err:
    MsgBox Err.Number & "-" & Err.Description, vbCritical, Me.Caption
End Sub

Private Sub txtNumDoc_GotFocus()
    Call ResaltarTexto(txtNumDoc)
End Sub

Private Sub txtNumDoc_KeyPress(KeyAscii As Integer)
    txtNumDoc.Text = txtNumDoc.Text + fc_ValNumeros(KeyAscii)
    If KeyAscii = 13 And txtNumDoc.Text <> "" Then
        txtNumDoc = Format(txtNumDoc, "0000000000000000")
        Call CargarDatosFactura
    End If
End Sub

Private Sub txtNumDoc_LostFocus()
    Call txtNumDoc_KeyPress(13)
End Sub

Private Sub txtSerie_LostFocus()
    Call txtSerie_KeyPress(13)
End Sub

Private Sub txtSerie_Change()
    If Len(txtSerie) = 4 Then txtSerie = Format(txtSerie, "0000"): If txtNumDoc.Visible Then txtNumDoc.SetFocus
    If Len(txtSerie) = 0 Then txtNumDoc.Text = ""
End Sub

Private Sub txtSerie_GotFocus()
    Call ResaltarTexto(txtSerie)
End Sub

Private Sub txtSerie_KeyPress(KeyAscii As Integer)
'    txtSerie.Text = txtSerie.Text + fc_ValNumeros(KeyAscii)
    If KeyAscii = 13 And txtSerie.Text <> "" Then txtSerie = Format(txtSerie, "0000"): txtNumDoc.SetFocus
End Sub

Public Sub Eliminar()
On Error GoTo Error

    Screen.MousePointer = vbHourglass
    Dim codprov As String, numdoc As String
    Dim iniciada As Boolean
    Dim DNI As String, precio As Currency
    
    codprov = Trim(txtCodProv.Text)
    numdoc = Trim(txtSerie.Text) & Trim(txtNumDoc.Text)
    documento = fc_CodigoComboBox(cboTipo, 2)
    
    If Len(codprov) = 0 Then
        MsgBox "Ingrese el código del Proveedor", vbExclamation, Me.Caption
        txtCodProv.SetFocus
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    If Len(Trim(txtSerie.Text)) = 0 Or Len(Trim(txtNumDoc.Text)) = 0 Then
        MsgBox "Ingrese el número de la factura", vbExclamation, Me.Caption
        txtSerie.SetFocus
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    If documento = "00" Then
        MsgBox "Seleccione el tipo de Exámen", vbExclamation, Me.Caption
        cboTipo.SetFocus
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    If MsgBox("Desea eliminar el registro de la factura" & numdoc & "?", vbQuestion + vbYesNo + vbDefaultButton1, "Eliminar Factura") = vbNo Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
        
    Sql = "BEGIN TRANSACTION"
    cn.Execute Sql, 64
    
    iniciada = True
    
    Sql = "update rrhh_detalle_documentos set status='*',"
    Sql = Sql & "user_modif='" & wuser & "',"
    Sql = Sql & "fec_modi=GETDATE() "
    Sql = Sql & "where tipo_doc='" & documento & "' "
    Sql = Sql & "and cod_prov='" & codprov & "' "
    Sql = Sql & "and num_doc='" & numdoc & "' "
    Sql = Sql & "and status<>'*' "
    cn.Execute Sql, 64
        
    Sql = "COMMIT TRANSACTION"
    cn.Execute Sql, 64
    

    MsgBox "Se ha eliminado la factura: " & numdoc, vbExclamation, Me.Caption
    
    Call Nuevo
        
    Screen.MousePointer = vbDefault
    
Exit Sub
Error:
    If iniciada Then
        Sql = "ROLLBACK TRANSACTION"
        cn.Execute Sql, 64
    End If

    MsgBox Err.Number & "-" & Err.Description, vbCritical, Me.Caption
    
End Sub

Public Sub Grabar()

On Error GoTo Error
    Screen.MousePointer = vbHourglass
    
    Dim codprov As String, numdoc As String, TipoDoc As String
    Dim iniciada As Boolean
    Dim PlaCod As String, DNI As String, nombres As String, codarea As String, precio As Currency
    Dim fecha As Date
    
    codprov = Trim(txtCodProv.Text)
    numdoc = Trim(txtSerie.Text) & Trim(txtNumDoc.Text)
    documento = fc_CodigoComboBox(cboTipo, 2)
    TipoDoc = "01" 'FACTURAS
    fecha = dtpFecha.Value
    
    If Len(codprov) = 0 Then
        MsgBox "Ingrese el código del Proveedor", vbExclamation, Me.Caption
        txtCodProv.SetFocus
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    If Len(Trim(txtSerie.Text)) = 0 Or Len(Trim(txtNumDoc.Text)) = 0 Then
        MsgBox "Ingrese el número de la factura", vbExclamation, Me.Caption
        txtSerie.SetFocus
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    If documento = "00" Then
        MsgBox "Seleccione el tipo de Exámen", vbExclamation, Me.Caption
        cboTipo.SetFocus
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    If rsdatos.RecordCount = 0 Then
        MsgBox "Ingrese empleados para generar el detalle del documento", vbExclamation, Me.Caption
        txtRutaArchivo.SetFocus
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    rsdatos.MoveFirst
    Do While Not rsdatos.EOF
'
        If IsNumeric(Trim(rsdatos!precio & "")) Then
            precio = CCur(Trim(rsdatos!precio & ""))
        Else
            precio = 0
        End If
'
        DNI = Trim(rsdatos!DNI & "")
        If Len(DNI) = 0 Then
            MsgBox "El Nro. de DNI " & DNI & " ingresado no es válido", vbExclamation, Me.Caption
            GrdPlanilla.SetFocus
            Exit Sub
        End If
'
        codarea = Trim(rsdatos!codarea & "")
        If Len(codarea) = 0 Then
            MsgBox "Debe registrar el area  para el DNI:" & DNI & "", vbExclamation, Me.Caption
            GrdPlanilla.SetFocus
            Exit Sub
        End If
'
        If precio = 0 Then
            MsgBox "El costo del Nro. de DNI " & DNI & " no es válido", vbExclamation, Me.Caption
            GrdPlanilla.SetFocus
            Exit Sub
        End If
'
        rsdatos.MoveNext
    Loop
    
    If editar Then
        If MsgBox("Desea actualizar el detalle de la factura " & numdoc & "?", vbQuestion + vbYesNo + vbDefaultButton1, "Actualizar Datos") = vbNo Then
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    Else
        If MsgBox("Desea registrar el detalle de la factura " & numdoc & "?", vbQuestion + vbYesNo + vbDefaultButton1, "Registrar Datos") = vbNo Then
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    End If
    
    Sql = "BEGIN TRANSACTION"
    cn.Execute Sql, 64
    
    iniciada = True
    
    Sql = "update rrhh_detalle_documentos set status='*',"
    Sql = Sql & "user_modif='" & wuser & "',"
    Sql = Sql & "fec_modi=GETDATE() "
    Sql = Sql & "where tipo_doc='" & documento & "' "
    Sql = Sql & "and cod_prov='" & codprov & "' "
    Sql = Sql & "and num_doc='" & numdoc & "' "
    Sql = Sql & "and status<>'*' "
    cn.Execute Sql, 64
        
    rsdatos.MoveFirst
        
    Do While Not rsdatos.EOF
    
        codarea = Trim(rsdatos!codarea & "")
        DNI = Trim(rsdatos!DNI & "")
        nombres = Trim(rsdatos!nombres & "")
        precio = CCur(rsdatos!precio)
        PlaCod = Trim(rsdatos!PlaCod & "")
        
        If Not RegistrarDetalle(codprov, TipoDoc, numdoc, fecha, _
            PlaCod, DNI, nombres, documento, _
            codarea, precio) Then
            GoTo Error
        End If
    
        rsdatos.MoveNext
    Loop
    
    Sql = "COMMIT TRANSACTION"
    cn.Execute Sql, 64
    
    If editar Then
        MsgBox "Se han actualizado los detalles de la factura: " & numdoc, vbExclamation, Me.Caption
    Else
        MsgBox "Se han registrado los detalles de la factura: " & numdoc, vbExclamation, Me.Caption
    End If
    
    Call CargarDatosFactura
    editar = True
        
    Screen.MousePointer = vbDefault
        
Exit Sub
Error:
    If iniciada Then
        Sql = "ROLLBACK TRANSACTION"
        cn.Execute Sql, 64
    End If
    Screen.MousePointer = vbDefault
    MsgBox Err.Number & "-" & Err.Description, vbCritical, Me.Caption
End Sub

Private Function RegistrarDetalle(ByVal pCodprov As String, _
    ByVal pTipoDoc As String, ByVal pNumDoc As String, ByVal pFecha As Date, _
    ByVal pPlacod As String, ByVal pDni As String, ByVal pNombres As String, _
    ByVal pDocumento As String, ByVal pCodArea As String, ByVal pPrecio As Currency) As Boolean
    
    On Error GoTo Error

    Sql = "exec rrhh_registrar_detalle_documento "
    Sql = Sql & "@codcia='" & wcia & "',"
    Sql = Sql & "@codprov='" & pCodprov & "',"
    Sql = Sql & "@tipodoc='" & pTipoDoc & "',"
    Sql = Sql & "@numdoc='" & pNumDoc & "',"
    Sql = Sql & "@fecha='" & Format(pFecha, "YYYYMMDD") & "',"
    Sql = Sql & "@placod='" & pPlacod & "',"
    Sql = Sql & "@dni='" & pDni & "',"
    Sql = Sql & "@nombres='" & pNombres & "',"
    Sql = Sql & "@documento='" & pDocumento & "',"
    Sql = Sql & "@codarea='" & pCodArea & "',"
    Sql = Sql & "@precio='" & pPrecio & "',"
    Sql = Sql & "@user='" & wuser & "'"
    Debug.Print Sql
    cn.Execute Sql, 64

    RegistrarDetalle = True
Exit Function
Error:
    MsgBox Err.Number & "-" & Err.Description, vbCritical, Me.Caption

    RegistrarDetalle = False
End Function

Public Sub Nuevo()

    editar = False
    
'    txtCodProv.Text = ""
'    txtRucProveedor.Text = ""
'    txtProveedor.Text = ""
    
    txtSerie.Text = "F001"
    txtNumDoc.Text = ""
    txtRutaArchivo.Text = ""
    
    Call rUbiIndCmbBox(cboTipo, "00", "00")
    
    txtRangoInicio.Text = ""
    txtRangoFin.Text = ""
    
    LimpiarRsT rsHojas, GrdHojas
    LimpiarRsT rsdatos, GrdPlanilla
    
    dtpFecha.Value = Now
    
    Call CalcularTotales(0, 0)
             
End Sub
