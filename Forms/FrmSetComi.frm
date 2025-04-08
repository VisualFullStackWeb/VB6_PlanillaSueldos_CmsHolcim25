VERSION 5.00
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmSetComi 
   Caption         =   "Tabla de parametros - Comision de venta"
   ClientHeight    =   9075
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15210
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9075
   ScaleWidth      =   15210
   Begin VB.CommandButton CmdClear 
      Caption         =   "L"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   360
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   2055
      Left            =   7800
      MousePointer    =   2  'Cross
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "FrmSetComi.frx":0000
      Top             =   4680
      Width           =   7095
   End
   Begin TrueOleDBGrid70.TDBGrid Grd 
      Height          =   8655
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   15266
      _LayoutType     =   4
      _RowHeight      =   17
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Escala"
      Columns(0).DataField=   "escala"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "PorcTonObjetivo"
      Columns(1).DataField=   "PorcTonObjetivo"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "PorcSolesObjetivo"
      Columns(2).DataField=   "PorcSolesObjetivo"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "factor_comision_previa"
      Columns(3).DataField=   "factor_comision_previa"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   4
      Splits(0)._UserFlags=   0
      Splits(0).MarqueeStyle=   3
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).FetchRowStyle=   -1  'True
      Splits(0).DividerColor=   13160660
      Splits(0).FilterBar=   -1  'True
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=4"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=3069"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2990"
      Splits(0)._ColumnProps(14)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(16)=   "Column(3).Width=2725"
      Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=2646"
      Splits(0)._ColumnProps(19)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
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
      Caption         =   "Tabla de parametros - Comision de venta"
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
      _StyleDefs(14)  =   ":id=2,.bold=-1,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(15)  =   ":id=2,.fontname=Arial"
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
      _StyleDefs(34)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8,.fgcolor=&H80000012&"
      _StyleDefs(35)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(36)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(37)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(38)  =   "Splits(0).FilterBarStyle:id=24,.parent=12,.bgcolor=&H80FFFF&"
      _StyleDefs(39)  =   ":id=24,.fgcolor=&H80000001&"
      _StyleDefs(40)  =   "Splits(0).Columns(0).Style:id=32,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
      _StyleDefs(52)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
      _StyleDefs(53)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
      _StyleDefs(54)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
      _StyleDefs(55)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
      _StyleDefs(56)  =   "Named:id=33:Normal"
      _StyleDefs(57)  =   ":id=33,.parent=0"
      _StyleDefs(58)  =   "Named:id=34:Heading"
      _StyleDefs(59)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(60)  =   ":id=34,.wraptext=-1"
      _StyleDefs(61)  =   "Named:id=35:Footing"
      _StyleDefs(62)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(63)  =   "Named:id=36:Selected"
      _StyleDefs(64)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(65)  =   "Named:id=37:Caption"
      _StyleDefs(66)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(67)  =   "Named:id=38:HighlightRow"
      _StyleDefs(68)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(69)  =   "Named:id=39:EvenRow"
      _StyleDefs(70)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(71)  =   "Named:id=40:OddRow"
      _StyleDefs(72)  =   ":id=40,.parent=33"
      _StyleDefs(73)  =   "Named:id=41:RecordSelector"
      _StyleDefs(74)  =   ":id=41,.parent=34"
      _StyleDefs(75)  =   "Named:id=42:FilterBar"
      _StyleDefs(76)  =   ":id=42,.parent=33"
   End
   Begin TrueOleDBGrid70.TDBGrid GrdCar 
      Height          =   4335
      Left            =   7800
      TabIndex        =   1
      Top             =   120
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   7646
      _LayoutType     =   4
      _RowHeight      =   17
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Cod. Trabajador"
      Columns(0).DataField=   "cod_trab"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Colaborador"
      Columns(1).DataField=   "Colaborador"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Meta Soles"
      Columns(2).DataField=   "pp_meta"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   3
      Splits(0)._UserFlags=   0
      Splits(0).MarqueeStyle=   3
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).FetchRowStyle=   -1  'True
      Splits(0).DividerColor=   13160660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=3"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2090"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2011"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=6826"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=6747"
      Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=2990"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2910"
      Splits(0)._ColumnProps(14)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
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
      Caption         =   "Metas PP 2023"
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
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37,.bgcolor=&HFF&,.fgcolor=&H80000014&"
      _StyleDefs(10)  =   ":id=4,.bold=-1,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(11)  =   ":id=4,.fontname=MS Sans Serif"
      _StyleDefs(12)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bgcolor=&HEFEFEF&,.appearance=0"
      _StyleDefs(13)  =   ":id=2,.bold=-1,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(14)  =   ":id=2,.fontname=Arial"
      _StyleDefs(15)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(16)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(17)  =   ":id=3,.fontname=MS Sans Serif"
      _StyleDefs(18)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(19)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(20)  =   "EditorStyle:id=7,.parent=1,.bgcolor=&HFFECD9&"
      _StyleDefs(21)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(22)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(23)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40,.bgcolor=&HFFEBD7&"
      _StyleDefs(24)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41,.bgcolor=&HEFEFEF&"
      _StyleDefs(25)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(26)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(27)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(28)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(29)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(30)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(31)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(32)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(33)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(34)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(35)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(36)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(37)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=32,.parent=13"
      _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
      _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
      _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
      _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
      _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
      _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
      _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
      _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
      _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
      _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
      _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
      _StyleDefs(50)  =   "Named:id=33:Normal"
      _StyleDefs(51)  =   ":id=33,.parent=0"
      _StyleDefs(52)  =   "Named:id=34:Heading"
      _StyleDefs(53)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(54)  =   ":id=34,.wraptext=-1"
      _StyleDefs(55)  =   "Named:id=35:Footing"
      _StyleDefs(56)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(57)  =   "Named:id=36:Selected"
      _StyleDefs(58)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(59)  =   "Named:id=37:Caption"
      _StyleDefs(60)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(61)  =   "Named:id=38:HighlightRow"
      _StyleDefs(62)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(63)  =   "Named:id=39:EvenRow"
      _StyleDefs(64)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(65)  =   "Named:id=40:OddRow"
      _StyleDefs(66)  =   ":id=40,.parent=33"
      _StyleDefs(67)  =   "Named:id=41:RecordSelector"
      _StyleDefs(68)  =   ":id=41,.parent=34"
      _StyleDefs(69)  =   "Named:id=42:FilterBar"
      _StyleDefs(70)  =   ":id=42,.parent=33"
   End
End
Attribute VB_Name = "FrmSetComi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim Col As TrueOleDBGrid70.Column
Dim Cols As TrueOleDBGrid70.Columns

Dim vPeriodo As String
Dim Rs As New ADODB.Recordset
Dim RsCar As New ADODB.Recordset


Private Sub CmdClear_Click()
LimpiarFiltros
End Sub

Private Sub Form_Load()
Me.Width = 15450: Me.Height = 9660
Crea_Rs
Exp_Estructura01
Exp_Estructura02

End Sub

Public Sub Crea_Rs()
    If Rs.State = 1 Then Rs.Close
    Rs.Fields.Append "escala", adChar, 12, adFldIsNullable
    Rs.Fields.Append "PorcTonObjetivo", adChar, 12, adFldIsNullable
    Rs.Fields.Append "PorcSolesObjetivo", adChar, 12, adFldIsNullable
    Rs.Fields.Append "factor_comision_previa", adChar, 12, adFldIsNullable
    Rs.Open
    
    If RsCar.State = 1 Then RsCar.Close
    RsCar.Fields.Append "cod_trab", adChar, 8, adFldIsNullable
    RsCar.Fields.Append "Colaborador", adVarChar, 120, adFldIsNullable
    RsCar.Fields.Append "pp_meta", adChar, 20, adFldIsNullable
  
    RsCar.Open
    
End Sub

Public Sub Exp_Estructura01()
'//*** ESTRUCTURA1: "Datos de Establecimientos Propios"
'On Error GoTo MsgErr:
Dim Reg As String
Dim I As Integer
I = 0
'xRutaFile = App.Path & "\reports\" & xRucCia & ".esp"

Dim Rq As ADODB.Recordset
Sql = "usp_pla_listar_parametros_comision_ventas '" & wcia & "'," & vPeriodo & ""
'Open xRutaFile For Output As #1
Screen.MousePointer = 11
If fAbrRst(Rq, Sql) Then
    'Barra.Max = Rq.RecordCount
    'Barra.Min = 1 - 1
    Do While Not Rq.EOF
    
        Rs.AddNew
        Rs!escala = Rq!escala
        Rs!PorcTonObjetivo = Rq!PorcTonObjetivo
        Rs!PorcSolesObjetivo = Rq!PorcSolesObjetivo
        Rs!factor_comision_previa = Rq!factor_comision_previa
       
        Rq.MoveNext
    Loop
    
    Set Grd.DataSource = Rs
    Grd.Refresh
    
    'Add_Mensaje LstError, "Estructura " & rs!NRO & ": ( " & CStr(I) & " ) registros exportados correctamente.", True, vbBlue
Else
    'Add_Mensaje LstError, "Estructura " & rs!NRO & ": no existen datos para exportar.", True, vbBlack
End If
'Barra.Value = 0

Screen.MousePointer = 0
'Close #1
Exit Sub
MsgErr:
'Close #1
MsgBox Err.Number & " - " & Err.Description, vbCritical, Me.Caption
Salir:
On Error Resume Next:
'Close #1
'If Dir(xRutaFile) <> "" Then DeleteAFile xRutaFile
'Add_Mensaje LstError, "Se canceló exportación " & Rs!ESTRUCTURA, True, vbRed
Screen.MousePointer = 0
End Sub


'Public Property Get PeriodoAno() As Variant
'PeriodoAno = vNewValue
'End Property

Public Property Let PeriodoAno(ByVal vNewValue As Variant)
vPeriodo = vNewValue
End Property

Private Sub Form_Unload(Cancel As Integer)
Rs.Close
Set Rs = Nothing
End Sub

Public Sub Exp_Estructura02()
'//*** ESTRUCTURA1: "Datos de Establecimientos Propios"
'On Error GoTo MsgErr:
Dim Reg As String
Dim I As Integer
I = 0
'xRutaFile = App.Path & "\reports\" & xRucCia & ".esp"

Dim Rq As ADODB.Recordset
Sql = "usp_pla_listar_parametros_comision_ventas_carteras '" & wcia & "'," & vPeriodo & ""
'Open xRutaFile For Output As #1
Screen.MousePointer = 11
If fAbrRst(Rq, Sql) Then
    'Barra.Max = Rq.RecordCount
    'Barra.Min = 1 - 1
    Do While Not Rq.EOF
    
        RsCar.AddNew
        RsCar!cod_trab = Trim(Rq!cod_vendedor & "")
        RsCar!Colaborador = Trim(Rq!nomtrabajador & "")
        RsCar!pp_meta = Format(Rq!Imp_Soles_Meta_Precio_Promedio_Por_Cartera, "###,###,##0.00")

        
       
        Rq.MoveNext
    Loop
    
    Set GrdCar.DataSource = RsCar
    GrdCar.Refresh
    
    'Add_Mensaje LstError, "Estructura " & rs!NRO & ": ( " & CStr(I) & " ) registros exportados correctamente.", True, vbBlue
Else
    'Add_Mensaje LstError, "Estructura " & rs!NRO & ": no existen datos para exportar.", True, vbBlack
End If
'Barra.Value = 0

Screen.MousePointer = 0
'Close #1
Exit Sub
MsgErr:
'Close #1
MsgBox Err.Number & " - " & Err.Description, vbCritical, Me.Caption
Salir:
On Error Resume Next:
'Close #1
'If Dir(xRutaFile) <> "" Then DeleteAFile xRutaFile
'Add_Mensaje LstError, "Se canceló exportación " & Rs!ESTRUCTURA, True, vbRed
Screen.MousePointer = 0
End Sub

Private Sub Grd_FilterChange()
On Error GoTo ErrHandler

Set Cols = Grd.Columns

Dim c As Integer

c = Grd.Col

Grd.HoldFields

Rs.Filter = getFilter()

Grd.Col = c

Grd.EditActive = True

Exit Sub

 

ErrHandler:

    MsgBox Err.Source & ":" & vbCrLf & Err.Description
    Me.LimpiarFiltros
End Sub


Public Sub LimpiarFiltros()
    For Each Col In Me.Grd.Columns
        Col.FilterText = ""
    Next Col
    Rs.Filter = adFilterNone
    Me.Grd.Refresh
    'If Not RsBi.EOF Then RsBi.MoveFirst
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

'Private Sub Grd_FilterChange()
'On Error GoTo ErrHandler
'
'Set Cols = Grd.Columns
'
'Dim c As Integer
'
'c = Grd.Col
'
'Grd.HoldFields
'
'Rs.Filter = getFilter()
'
'Grd.Col = c
'
'Grd.EditActive = True
'
'Exit Sub
'
'
'
'ErrHandler:
'
'    MsgBox Err.Source & ":" & vbCrLf & Err.Description
'    Me.LimpiarFiltros
'
'End Sub
