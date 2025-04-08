VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0D623638-DBA2-11D1-B5DF-0060976089D0}#7.0#0"; "tdbg7.ocx"
Begin VB.Form frmboletamasiva 
   BackColor       =   &H00EBFEFC&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Generacion de Boletas Masivas"
   ClientHeight    =   6075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   11895
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox pctsemanas 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   2385
      ScaleHeight     =   480
      ScaleWidth      =   5925
      TabIndex        =   7
      Top             =   540
      Visible         =   0   'False
      Width           =   5955
      Begin VB.TextBox Txtsemana 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   990
         TabIndex        =   8
         Top             =   105
         Width           =   375
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   375
         Left            =   1350
         TabIndex        =   9
         Top             =   90
         Width           =   240
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.DTPicker Cmbal 
         Height          =   255
         Left            =   4290
         TabIndex        =   10
         Top             =   150
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   450
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   70778881
         CurrentDate     =   37265
      End
      Begin MSComCtl2.DTPicker Cmbdel 
         Height          =   255
         Left            =   2505
         TabIndex        =   11
         Top             =   150
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   450
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   70778881
         CurrentDate     =   37267
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Al"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00EBFEFC&
         Height          =   210
         Left            =   4005
         TabIndex        =   14
         Top             =   150
         Width           =   180
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Del"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00EBFEFC&
         Height          =   210
         Left            =   2100
         TabIndex        =   13
         Top             =   150
         Width           =   285
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Semana"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00EBFEFC&
         Height          =   210
         Left            =   180
         TabIndex        =   12
         Top             =   150
         Width           =   720
      End
   End
   Begin TrueDBGrid70.TDBGrid tdbg_boletas 
      Height          =   4875
      Left            =   90
      TabIndex        =   0
      Top             =   1125
      Width           =   11715
      _ExtentX        =   20664
      _ExtentY        =   8599
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   68
      Columns(0)._MaxComboItems=   5
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Personal"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Aportaciones"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Descuentos"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   4
      Splits(0)._UserFlags=   0
      Splits(0).SizeMode=   1
      Splits(0).Size  =   3495.118
      Splits(0).Size.vt=   4
      Splits(0).RecordSelectorWidth=   503
      Splits(0).Caption=   "Datos Personales"
      Splits(0).DividerColor=   13160660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=4"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=423"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=344"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=8414"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=8334"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=516"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=2725"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2646"
      Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=516"
      Splits(0)._ColumnProps(15)=   "Column(2).Visible=0"
      Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(17)=   "Column(3).Width=2725"
      Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=2646"
      Splits(0)._ColumnProps(20)=   "Column(3)._ColStyle=516"
      Splits(0)._ColumnProps(21)=   "Column(3).Order=4"
      Splits(1)._UserFlags=   0
      Splits(1).SizeMode=   1
      Splits(1).Size  =   4995.213
      Splits(1).Size.vt=   4
      Splits(1).RecordSelectors=   0   'False
      Splits(1).RecordSelectorWidth=   503
      Splits(1).Caption=   "Variables de Calculo"
      Splits(1).DividerColor=   13160660
      Splits(1).SpringMode=   0   'False
      Splits(1)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(1)._ColumnProps(0)=   "Columns.Count=4"
      Splits(1)._ColumnProps(1)=   "Column(0).Width=529"
      Splits(1)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(1)._ColumnProps(3)=   "Column(0)._WidthInPix=450"
      Splits(1)._ColumnProps(4)=   "Column(0)._ColStyle=516"
      Splits(1)._ColumnProps(5)=   "Column(0).Visible=0"
      Splits(1)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(1)._ColumnProps(7)=   "Column(1).Width=8414"
      Splits(1)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(1)._ColumnProps(9)=   "Column(1)._WidthInPix=8334"
      Splits(1)._ColumnProps(10)=   "Column(1)._ColStyle=516"
      Splits(1)._ColumnProps(11)=   "Column(1).Visible=0"
      Splits(1)._ColumnProps(12)=   "Column(1).Order=2"
      Splits(1)._ColumnProps(13)=   "Column(2).Width=2725"
      Splits(1)._ColumnProps(14)=   "Column(2).DividerColor=0"
      Splits(1)._ColumnProps(15)=   "Column(2)._WidthInPix=2646"
      Splits(1)._ColumnProps(16)=   "Column(2)._ColStyle=516"
      Splits(1)._ColumnProps(17)=   "Column(2).Order=3"
      Splits(1)._ColumnProps(18)=   "Column(3).Width=2725"
      Splits(1)._ColumnProps(19)=   "Column(3).DividerColor=0"
      Splits(1)._ColumnProps(20)=   "Column(3)._WidthInPix=2646"
      Splits(1)._ColumnProps(21)=   "Column(3)._ColStyle=516"
      Splits(1)._ColumnProps(22)=   "Column(3).Visible=0"
      Splits(1)._ColumnProps(23)=   "Column(3).Order=4"
      Splits(2)._UserFlags=   0
      Splits(2).RecordSelectors=   0   'False
      Splits(2).RecordSelectorWidth=   503
      Splits(2).Caption=   "Descuentos "
      Splits(2).DividerColor=   13160660
      Splits(2).SpringMode=   0   'False
      Splits(2)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(2)._ColumnProps(0)=   "Columns.Count=4"
      Splits(2)._ColumnProps(1)=   "Column(0).Width=529"
      Splits(2)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(2)._ColumnProps(3)=   "Column(0)._WidthInPix=450"
      Splits(2)._ColumnProps(4)=   "Column(0)._ColStyle=516"
      Splits(2)._ColumnProps(5)=   "Column(0).Visible=0"
      Splits(2)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(2)._ColumnProps(7)=   "Column(1).Width=8414"
      Splits(2)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(2)._ColumnProps(9)=   "Column(1)._WidthInPix=8334"
      Splits(2)._ColumnProps(10)=   "Column(1)._ColStyle=516"
      Splits(2)._ColumnProps(11)=   "Column(1).Visible=0"
      Splits(2)._ColumnProps(12)=   "Column(1).Order=2"
      Splits(2)._ColumnProps(13)=   "Column(2).Width=2725"
      Splits(2)._ColumnProps(14)=   "Column(2).DividerColor=0"
      Splits(2)._ColumnProps(15)=   "Column(2)._WidthInPix=2646"
      Splits(2)._ColumnProps(16)=   "Column(2)._ColStyle=516"
      Splits(2)._ColumnProps(17)=   "Column(2).Visible=0"
      Splits(2)._ColumnProps(18)=   "Column(2).Order=3"
      Splits(2)._ColumnProps(19)=   "Column(3).Width=2725"
      Splits(2)._ColumnProps(20)=   "Column(3).DividerColor=0"
      Splits(2)._ColumnProps(21)=   "Column(3)._WidthInPix=2646"
      Splits(2)._ColumnProps(22)=   "Column(3)._ColStyle=516"
      Splits(2)._ColumnProps(23)=   "Column(3).Order=4"
      Splits.Count    =   3
      PrintInfos(0)._StateFlags=   0
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowDelete     =   -1  'True
      AllowAddNew     =   -1  'True
      Appearance      =   0
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      Caption         =   "Planilla de Personal"
      TabAcrossSplits =   -1  'True
      TabAction       =   2
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   13160660
      RowDividerColor =   13160660
      RowSubDividerColor=   13160660
      DirectionAfterEnter=   1
      MaxRows         =   250000
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33"
      _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
      _StyleDefs(9)   =   "FooterStyle:id=3,.parent=1,.namedParent=35"
      _StyleDefs(10)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(11)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(12)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(13)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(14)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(15)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(16)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(17)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(18)  =   "Splits(0).Style:id=47,.parent=1"
      _StyleDefs(19)  =   "Splits(0).CaptionStyle:id=56,.parent=4"
      _StyleDefs(20)  =   "Splits(0).HeadingStyle:id=48,.parent=2"
      _StyleDefs(21)  =   "Splits(0).FooterStyle:id=49,.parent=3"
      _StyleDefs(22)  =   "Splits(0).InactiveStyle:id=50,.parent=5"
      _StyleDefs(23)  =   "Splits(0).SelectedStyle:id=52,.parent=6"
      _StyleDefs(24)  =   "Splits(0).EditorStyle:id=51,.parent=7"
      _StyleDefs(25)  =   "Splits(0).HighlightRowStyle:id=53,.parent=8"
      _StyleDefs(26)  =   "Splits(0).EvenRowStyle:id=54,.parent=9"
      _StyleDefs(27)  =   "Splits(0).OddRowStyle:id=55,.parent=10"
      _StyleDefs(28)  =   "Splits(0).RecordSelectorStyle:id=57,.parent=11"
      _StyleDefs(29)  =   "Splits(0).FilterBarStyle:id=58,.parent=12"
      _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=62,.parent=47"
      _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=59,.parent=48"
      _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=60,.parent=49"
      _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=61,.parent=51"
      _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=66,.parent=47"
      _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=63,.parent=48"
      _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=64,.parent=49"
      _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=65,.parent=51"
      _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=70,.parent=47"
      _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=67,.parent=48"
      _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=68,.parent=49"
      _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=69,.parent=51"
      _StyleDefs(42)  =   "Splits(0).Columns(3).Style:id=74,.parent=47"
      _StyleDefs(43)  =   "Splits(0).Columns(3).HeadingStyle:id=71,.parent=48"
      _StyleDefs(44)  =   "Splits(0).Columns(3).FooterStyle:id=72,.parent=49"
      _StyleDefs(45)  =   "Splits(0).Columns(3).EditorStyle:id=73,.parent=51"
      _StyleDefs(46)  =   "Splits(1).Style:id=13,.parent=1"
      _StyleDefs(47)  =   "Splits(1).CaptionStyle:id=22,.parent=4"
      _StyleDefs(48)  =   "Splits(1).HeadingStyle:id=14,.parent=2"
      _StyleDefs(49)  =   "Splits(1).FooterStyle:id=15,.parent=3"
      _StyleDefs(50)  =   "Splits(1).InactiveStyle:id=16,.parent=5"
      _StyleDefs(51)  =   "Splits(1).SelectedStyle:id=18,.parent=6"
      _StyleDefs(52)  =   "Splits(1).EditorStyle:id=17,.parent=7"
      _StyleDefs(53)  =   "Splits(1).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(54)  =   "Splits(1).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(55)  =   "Splits(1).OddRowStyle:id=21,.parent=10"
      _StyleDefs(56)  =   "Splits(1).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(57)  =   "Splits(1).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(58)  =   "Splits(1).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(59)  =   "Splits(1).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(60)  =   "Splits(1).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(61)  =   "Splits(1).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(62)  =   "Splits(1).Columns(1).Style:id=32,.parent=13"
      _StyleDefs(63)  =   "Splits(1).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(64)  =   "Splits(1).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(65)  =   "Splits(1).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(66)  =   "Splits(1).Columns(2).Style:id=46,.parent=13"
      _StyleDefs(67)  =   "Splits(1).Columns(2).HeadingStyle:id=43,.parent=14"
      _StyleDefs(68)  =   "Splits(1).Columns(2).FooterStyle:id=44,.parent=15"
      _StyleDefs(69)  =   "Splits(1).Columns(2).EditorStyle:id=45,.parent=17"
      _StyleDefs(70)  =   "Splits(1).Columns(3).Style:id=78,.parent=13"
      _StyleDefs(71)  =   "Splits(1).Columns(3).HeadingStyle:id=75,.parent=14"
      _StyleDefs(72)  =   "Splits(1).Columns(3).FooterStyle:id=76,.parent=15"
      _StyleDefs(73)  =   "Splits(1).Columns(3).EditorStyle:id=77,.parent=17"
      _StyleDefs(74)  =   "Splits(2).Style:id=79,.parent=1"
      _StyleDefs(75)  =   "Splits(2).CaptionStyle:id=88,.parent=4"
      _StyleDefs(76)  =   "Splits(2).HeadingStyle:id=80,.parent=2"
      _StyleDefs(77)  =   "Splits(2).FooterStyle:id=81,.parent=3"
      _StyleDefs(78)  =   "Splits(2).InactiveStyle:id=82,.parent=5"
      _StyleDefs(79)  =   "Splits(2).SelectedStyle:id=84,.parent=6"
      _StyleDefs(80)  =   "Splits(2).EditorStyle:id=83,.parent=7"
      _StyleDefs(81)  =   "Splits(2).HighlightRowStyle:id=85,.parent=8"
      _StyleDefs(82)  =   "Splits(2).EvenRowStyle:id=86,.parent=9"
      _StyleDefs(83)  =   "Splits(2).OddRowStyle:id=87,.parent=10"
      _StyleDefs(84)  =   "Splits(2).RecordSelectorStyle:id=89,.parent=11"
      _StyleDefs(85)  =   "Splits(2).FilterBarStyle:id=90,.parent=12"
      _StyleDefs(86)  =   "Splits(2).Columns(0).Style:id=94,.parent=79"
      _StyleDefs(87)  =   "Splits(2).Columns(0).HeadingStyle:id=91,.parent=80"
      _StyleDefs(88)  =   "Splits(2).Columns(0).FooterStyle:id=92,.parent=81"
      _StyleDefs(89)  =   "Splits(2).Columns(0).EditorStyle:id=93,.parent=83"
      _StyleDefs(90)  =   "Splits(2).Columns(1).Style:id=98,.parent=79"
      _StyleDefs(91)  =   "Splits(2).Columns(1).HeadingStyle:id=95,.parent=80"
      _StyleDefs(92)  =   "Splits(2).Columns(1).FooterStyle:id=96,.parent=81"
      _StyleDefs(93)  =   "Splits(2).Columns(1).EditorStyle:id=97,.parent=83"
      _StyleDefs(94)  =   "Splits(2).Columns(2).Style:id=102,.parent=79"
      _StyleDefs(95)  =   "Splits(2).Columns(2).HeadingStyle:id=99,.parent=80"
      _StyleDefs(96)  =   "Splits(2).Columns(2).FooterStyle:id=100,.parent=81"
      _StyleDefs(97)  =   "Splits(2).Columns(2).EditorStyle:id=101,.parent=83"
      _StyleDefs(98)  =   "Splits(2).Columns(3).Style:id=106,.parent=79"
      _StyleDefs(99)  =   "Splits(2).Columns(3).HeadingStyle:id=103,.parent=80"
      _StyleDefs(100) =   "Splits(2).Columns(3).FooterStyle:id=104,.parent=81"
      _StyleDefs(101) =   "Splits(2).Columns(3).EditorStyle:id=105,.parent=83"
      _StyleDefs(102) =   "Named:id=33:Normal"
      _StyleDefs(103) =   ":id=33,.parent=0"
      _StyleDefs(104) =   "Named:id=34:Heading"
      _StyleDefs(105) =   ":id=34,.parent=33,.alignment=2,.valignment=2,.bgcolor=&HF1FEFD&"
      _StyleDefs(106) =   ":id=34,.fgcolor=&H80000012&,.wraptext=-1,.bold=-1,.fontsize=900,.italic=0"
      _StyleDefs(107) =   ":id=34,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(108) =   ":id=34,.fontname=Tahoma"
      _StyleDefs(109) =   "Named:id=35:Footing"
      _StyleDefs(110) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(111) =   "Named:id=36:Selected"
      _StyleDefs(112) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(113) =   "Named:id=37:Caption"
      _StyleDefs(114) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(115) =   "Named:id=38:HighlightRow"
      _StyleDefs(116) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(117) =   "Named:id=39:EvenRow"
      _StyleDefs(118) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(119) =   "Named:id=40:OddRow"
      _StyleDefs(120) =   ":id=40,.parent=33"
      _StyleDefs(121) =   "Named:id=41:RecordSelector"
      _StyleDefs(122) =   ":id=41,.parent=34"
      _StyleDefs(123) =   "Named:id=42:FilterBar"
      _StyleDefs(124) =   ":id=42,.parent=33"
   End
   Begin VB.ComboBox Cmbtipo 
      Height          =   315
      Left            =   1620
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   135
      Width           =   2295
   End
   Begin VB.ComboBox Cmbtipotrabajador 
      Height          =   315
      Left            =   6300
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   135
      Width           =   2055
   End
   Begin MSComCtl2.DTPicker Cmbfecha 
      Height          =   255
      Left            =   780
      TabIndex        =   2
      Top             =   705
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   450
      _Version        =   393216
      Format          =   70778881
      CurrentDate     =   37265
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo de Boleta"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00EBFEFC&
      Height          =   210
      Left            =   180
      TabIndex        =   6
      Top             =   225
      Width           =   1320
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00EBFEFC&
      Height          =   210
      Left            =   225
      TabIndex        =   5
      Top             =   705
      Width           =   510
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo Trabajador"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00EBFEFC&
      Height          =   210
      Left            =   4740
      TabIndex        =   4
      Top             =   225
      Width           =   1440
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      Height          =   6000
      Left            =   45
      Top             =   45
      Width           =   11805
   End
End
Attribute VB_Name = "frmboletamasiva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

