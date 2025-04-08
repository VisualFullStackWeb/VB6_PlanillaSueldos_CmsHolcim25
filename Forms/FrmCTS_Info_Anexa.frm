VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmCTS_Info_Anexa 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "» Parametos Certificado de CTS «"
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8130
   Icon            =   "FrmCTS_Info_Anexa.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   8130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Chk_01 
      Caption         =   "Reimpresión"
      ForeColor       =   &H00000000&
      Height          =   200
      Left            =   6465
      TabIndex        =   18
      Top             =   150
      Width           =   1320
   End
   Begin VB.Frame frm 
      Height          =   1020
      Index           =   1
      Left            =   6465
      TabIndex        =   15
      Top             =   360
      Width           =   1485
      Begin MSForms.CommandButton btn_Salir 
         Height          =   375
         Left            =   105
         TabIndex        =   17
         Top             =   555
         Width           =   1245
         Caption         =   "         Salir"
         PicturePosition =   327683
         Size            =   "2196;661"
         Picture         =   "FrmCTS_Info_Anexa.frx":030A
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton btn_Aceptar 
         Height          =   375
         Left            =   105
         TabIndex        =   16
         Top             =   195
         Width           =   1245
         Caption         =   "     Aceptar"
         PicturePosition =   327683
         Size            =   "2196;661"
         Picture         =   "FrmCTS_Info_Anexa.frx":08A4
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
   End
   Begin VB.CheckBox Chk 
      BackColor       =   &H80000010&
      Caption         =   "SeleccionarTodos"
      ForeColor       =   &H00FFFFFF&
      Height          =   200
      Left            =   225
      TabIndex        =   14
      Top             =   2295
      Width           =   1635
   End
   Begin VB.Frame frm 
      Caption         =   "Parametros del Depósito"
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
      Height          =   780
      Index           =   0
      Left            =   200
      TabIndex        =   6
      Top             =   1380
      Width           =   5145
      Begin VB.TextBox txt_T_Cambio 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   3735
         TabIndex        =   7
         Top             =   315
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker dtp_Fecha 
         Height          =   315
         Left            =   840
         TabIndex        =   8
         Top             =   315
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   62914561
         CurrentDate     =   40731
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Cambio"
         Height          =   195
         Index           =   4
         Left            =   2430
         TabIndex        =   10
         Top             =   315
         Width           =   1110
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha"
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   9
         Top             =   315
         Width           =   450
      End
   End
   Begin VB.ComboBox Cbo_Tipo_Moneda 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1620
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   930
      Width           =   3735
   End
   Begin VB.ComboBox Cbo_Cta_Cte 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1620
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   525
      Width           =   3735
   End
   Begin VB.ComboBox Cbo_Banco 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1620
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
   Begin TrueOleDBGrid70.TDBGrid grid 
      Height          =   4545
      Left            =   180
      TabIndex        =   11
      Top             =   2265
      Width           =   7755
      _ExtentX        =   13679
      _ExtentY        =   8017
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   68
      Columns(0)._MaxComboItems=   5
      Columns(0).DataField=   "ESTADO"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Código"
      Columns(1).DataField=   "PLACOD"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Personal"
      Columns(2).DataField=   "NOMBRE"
      Columns(2).NumberFormat=   "Standard"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Fec .Depósito"
      Columns(3).DataField=   "FECHA_DDEPOSITO"
      Columns(3).NumberFormat=   "Short Date"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "T. Cambio"
      Columns(4).DataField=   "TIPO_DCAMBIO"
      Columns(4).NumberFormat=   "General Number"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   5
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   13160660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=5"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=529"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=450"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=131585"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=1244"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=1164"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=8705"
      Splits(0)._ColumnProps(10)=   "Column(1).Visible=0"
      Splits(0)._ColumnProps(11)=   "Column(1).AllowFocus=0"
      Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(13)=   "Column(2).Width=7938"
      Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=7858"
      Splits(0)._ColumnProps(16)=   "Column(2)._ColStyle=8704"
      Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(18)=   "Column(3).Width=2302"
      Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=2223"
      Splits(0)._ColumnProps(21)=   "Column(3)._ColStyle=8705"
      Splits(0)._ColumnProps(22)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(23)=   "Column(4).Width=1773"
      Splits(0)._ColumnProps(24)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(25)=   "Column(4)._WidthInPix=1693"
      Splits(0)._ColumnProps(26)=   "Column(4)._ColStyle=8706"
      Splits(0)._ColumnProps(27)=   "Column(4).Order=5"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      Appearance      =   2
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      Caption         =   "Seleccione Personal a Procesar"
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
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37,.bgcolor=&H80000010&"
      _StyleDefs(10)  =   ":id=4,.fgcolor=&H80000014&,.appearance=0,.bold=-1,.fontsize=825,.italic=0"
      _StyleDefs(11)  =   ":id=4,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(12)  =   ":id=4,.fontname=Arial"
      _StyleDefs(13)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bgcolor=&H8000000F&,.bold=-1"
      _StyleDefs(14)  =   ":id=2,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(15)  =   ":id=2,.fontname=Arial"
      _StyleDefs(16)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bgcolor=&H8000000F&,.bold=0"
      _StyleDefs(17)  =   ":id=3,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(18)  =   ":id=3,.fontname=MS Sans Serif"
      _StyleDefs(19)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(20)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(21)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(22)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(23)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(24)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(25)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(26)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(27)  =   "Splits(0).Style:id=13,.parent=1,.bold=0,.fontsize=825,.italic=0,.underline=0"
      _StyleDefs(28)  =   ":id=13,.strikethrough=0,.charset=0"
      _StyleDefs(29)  =   ":id=13,.fontname=MS Sans Serif"
      _StyleDefs(30)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(31)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(32)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(33)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(34)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(35)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(36)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(37)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(38)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(39)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(40)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(41)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.alignment=2"
      _StyleDefs(42)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14,.bold=-1,.fontsize=825"
      _StyleDefs(43)  =   ":id=25,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(44)  =   ":id=25,.fontname=Arial"
      _StyleDefs(45)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15,.alignment=2,.appearance=0"
      _StyleDefs(46)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17,.alignment=2"
      _StyleDefs(47)  =   "Splits(0).Columns(1).Style:id=46,.parent=13,.alignment=2,.locked=-1"
      _StyleDefs(48)  =   "Splits(0).Columns(1).HeadingStyle:id=43,.parent=14"
      _StyleDefs(49)  =   "Splits(0).Columns(1).FooterStyle:id=44,.parent=15"
      _StyleDefs(50)  =   "Splits(0).Columns(1).EditorStyle:id=45,.parent=17"
      _StyleDefs(51)  =   "Splits(0).Columns(2).Style:id=62,.parent=13,.alignment=0,.locked=-1,.bold=0"
      _StyleDefs(52)  =   ":id=62,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(53)  =   ":id=62,.fontname=MS Sans Serif"
      _StyleDefs(54)  =   "Splits(0).Columns(2).HeadingStyle:id=59,.parent=14,.fgcolor=&H80000012&,.bold=-1"
      _StyleDefs(55)  =   ":id=59,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(56)  =   ":id=59,.fontname=Arial"
      _StyleDefs(57)  =   "Splits(0).Columns(2).FooterStyle:id=60,.parent=15"
      _StyleDefs(58)  =   "Splits(0).Columns(2).EditorStyle:id=61,.parent=17"
      _StyleDefs(59)  =   "Splits(0).Columns(3).Style:id=32,.parent=13,.alignment=2,.locked=-1"
      _StyleDefs(60)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14,.bold=-1,.fontsize=825"
      _StyleDefs(61)  =   ":id=29,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(62)  =   ":id=29,.fontname=Arial"
      _StyleDefs(63)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
      _StyleDefs(64)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
      _StyleDefs(65)  =   "Splits(0).Columns(4).Style:id=50,.parent=13,.alignment=1,.locked=-1"
      _StyleDefs(66)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14"
      _StyleDefs(67)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
      _StyleDefs(68)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
      _StyleDefs(69)  =   "Named:id=33:Normal"
      _StyleDefs(70)  =   ":id=33,.parent=0"
      _StyleDefs(71)  =   "Named:id=34:Heading"
      _StyleDefs(72)  =   ":id=34,.parent=33,.alignment=2,.valignment=2,.bgcolor=&H8000000F&"
      _StyleDefs(73)  =   ":id=34,.fgcolor=&H80000012&,.wraptext=-1,.bold=-1,.fontsize=900,.italic=0"
      _StyleDefs(74)  =   ":id=34,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(75)  =   ":id=34,.fontname=Arial"
      _StyleDefs(76)  =   "Named:id=35:Footing"
      _StyleDefs(77)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(78)  =   "Named:id=36:Selected"
      _StyleDefs(79)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(80)  =   "Named:id=37:Caption"
      _StyleDefs(81)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(82)  =   "Named:id=38:HighlightRow"
      _StyleDefs(83)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(84)  =   "Named:id=39:EvenRow"
      _StyleDefs(85)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(86)  =   "Named:id=40:OddRow"
      _StyleDefs(87)  =   ":id=40,.parent=33"
      _StyleDefs(88)  =   "Named:id=41:RecordSelector"
      _StyleDefs(89)  =   ":id=41,.parent=34"
      _StyleDefs(90)  =   "Named:id=42:FilterBar"
      _StyleDefs(91)  =   ":id=42,.parent=33"
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Height          =   195
      Index           =   6
      Left            =   7575
      TabIndex        =   13
      Top             =   1965
      Width           =   345
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Registros encontrados :"
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
      Height          =   210
      Index           =   5
      Left            =   5550
      TabIndex        =   12
      Top             =   1965
      Width           =   1995
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo de Moneda"
      Height          =   195
      Index           =   2
      Left            =   200
      TabIndex        =   5
      Top             =   930
      Width           =   1170
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo de Cta. Cte."
      Height          =   195
      Index           =   1
      Left            =   200
      TabIndex        =   4
      Top             =   525
      Width           =   1200
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Entidad Bancaria"
      Height          =   195
      Index           =   0
      Left            =   200
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "FrmCTS_Info_Anexa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Id_Bco      As String
Dim Id_CtaCte   As String
Dim T_Moneda    As String
Public iYear    As Integer
Public iMonth   As Integer
Public bBoolean As Boolean
Dim rsTemporal  As ADODB.Recordset

Private Sub Trae_Ent_Bancaria()
    Cadena = "SP_TRAE_ENT_BANCARIA"
    Call rCarCbo(Cbo_Banco, Cadena, "XX", "00")
    Cbo_Banco.ListIndex = 0
End Sub

Private Sub Trae_Tipo_Moneda()
    Cadena = "SP_TIPO_MONEDA"
    Call rCarCbo(Cbo_Tipo_Moneda, Cadena, "XX", "00")
    Cbo_Tipo_Moneda.ListIndex = 0
End Sub

Private Sub Trae_Tipo_CtaCte()
    Cadena = "SP_TIPO_CTA_CTE"
    Call rCarCbo(Cbo_Cta_Cte, Cadena, "XX", "00")
    Cbo_Cta_Cte.ListIndex = 0
End Sub

Private Sub Trae_Informacion()
    'Exit Sub
    Call mRecordSet
    Dim rs As ADODB.Recordset
    Cadena = "SP_TRAE_INFO_CTS " & _
            "'" & wcia & "', " & _
            "" & iYear & ", " & _
            "" & iMonth & ", " & _
            "'" & Id_Bco & "', " & _
            "'" & T_Moneda & "', " & _
            "'" & Id_CtaCte & "'"
    Set rs = OpenRecordset(Cadena, cn)
    If Not rs.EOF Then
        Do While Not rs.EOF
            DoEvents
            With rsTemporal
                .AddNew
                !estado = IIf(Val(rs!estado) = 0, False, True)
                !PlaCod = rs!PlaCod
                !nombre = rs!nombre
                !Fecha_dDeposito = rs!Fecha_dDeposito
                !TIPO_DCAMBIO = rs!TIPO_DCAMBIO
            End With
            rs.MoveNext
        Loop
    End If
    lbl(6).Caption = rs.RecordCount
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
        Set rs = Nothing
    End If
    grid.Refresh
End Sub

Private Sub mRecordSet()
    Set rsTemporal = New ADODB.Recordset
    With rsTemporal
        If .State = adStateOpen Then .Close
        With .Fields
            .Append "ESTADO", adBoolean, 50, adFldIsNullable
            .Append "PLACOD", adVarChar, 10, adFldIsNullable
            .Append "NOMBRE", adVarChar, 250, adFldIsNullable
            .Append "FECHA_DDEPOSITO", adDate, , adFldIsNullable
            .Append "TIPO_DCAMBIO", adDouble, , adFldIsNullable
        End With
        .Open
    End With
    Set grid.DataSource = rsTemporal
End Sub

Public Sub Procesar()
If MsgBox("Desea Continuar?", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
    Dim Cad_Aux         As String
    Dim InTrans         As Boolean
    Dim bSolo_Impresion As Boolean
    Dim mFechaAux       As Date
On Error GoTo MyErr
    bSolo_Impresion = Chk_01.Value
    If bSolo_Impresion = False Then
        mFechaAux = DateSerial(iYear, iMonth + 1, 0)
        If (Not IsDate(dtp_Fecha.Value)) Or dtp_Fecha.Value < mFechaAux Then
            MsgBox "FECHA : Valor incorrecto, Verifique.", vbExclamation + vbOKOnly + Me.Caption
            Exit Sub
        End If
        If (Not IsNumeric(txt_T_Cambio.Text)) Or Val(txt_T_Cambio.Text) = 0 Or Val(txt_T_Cambio.Text) < 0 Then
            MsgBox "TIPO DE CAMBIO : Valor incorrecto, Verifique.", vbExclamation + vbOKOnly + Me.Caption
            txt_T_Cambio.SetFocus
            Exit Sub
        End If
    
        cn.CommandTimeout = 0
        cn.BeginTrans: InTrans = True
    End If
    With rsTemporal
        If .RecordCount > 0 Then
            Cad_Aux = "("
            .MoveFirst
            Do While Not .EOF
                Debug.Print .AbsolutePosition & Space(1) & !PlaCod & Space(1) & !estado
                If !estado Then
                    mBoolean = True
                    Cad_Aux = Cad_Aux & "'" & Trim(!PlaCod) & "'"
'                    If bSolo_Impresion = False Then
'                        Cadena = "SP_SALVA_DATOS_CTS " & _
'                                "'" & wcia & "', " & _
'                                "" & iYear & ", " & _
'                                "" & iMonth & ", " & _
'                                "'" & Trim(!PlaCod) & "', " & _
'                                "" & CDbl(txt_T_Cambio.Text) & ", " & _
'                                "'" & mFormato_Fecha(dtp_Fecha.Value) & "'"
'                        cn.Execute (Cadena)
'                    End If
                End If
                rsTemporal.MoveNext
                If mBoolean Then Cad_Aux = Cad_Aux & ",": mBoolean = False
            Loop
            Cad_Aux = Mid(Cad_Aux, 1, Len(Trim(Cad_Aux)) - 1)
            Cad_Aux = Cad_Aux & ")"
            Debug.Print Cad_Aux
        End If
        If Len(Cad_Aux) < 7 Then MsgBox "Debe de Seleccionar al menos un trabajador, Verifique.", vbExclamation + vbOKOnly, Me.Caption: GoTo MyErr
'        If bSolo_Impresion = False Then
'            Cadena = "SP_SALVAR_CERTIFICADO_CTS '" & wcia & "', " & iYear & ", " & iMonth & ", " & CDbl(txt_T_Cambio.Text) & ", '" & mFormato_Fecha(dtp_Fecha.Value) & "', '" & wuser & "'"
'            cn.Execute (Cadena)
'        End If
    End With
    If bSolo_Impresion = False Then cn.CommitTrans: InTrans = False
    bBoolean = True
        With FrmCts
            .Id_Trab = Empty
            .Id_Trab = Cad_Aux
            .T_dCambio = CDbl(Val(txt_T_Cambio.Text))
            .Fecha_dDeposito = dtp_Fecha.Value
            .Interes_Moratorio = 0 'PARA CUANDO SEA NECESARIO AGREGARLO AL CALCULO
        End With
MyErr:
    If bSolo_Impresion = False Then
        If InTrans Then cn.RollbackTrans
        
        If ERR.Number <> 0 Then
            MsgBox ERR.Description, vbCritical, Me.Caption
            ERR.Clear
        End If
    End If
    Unload Me
End If
End Sub

Private Sub Inicia_Form()
    Call Trae_Ent_Bancaria
    Call Trae_Tipo_Moneda
    Call Trae_Tipo_CtaCte
End Sub

Private Sub btn_aceptar_Click()
    Call Procesar
End Sub

Private Sub btn_salir_Click()
    Unload Me
End Sub

Private Sub Cbo_Banco_Click()
    Id_Bco = Empty
    Id_Bco = Trim(fc_CodigoComboBox(Cbo_Banco, 2))
    If Id_Bco = "99" Then Id_Bco = Empty
    Call Trae_Informacion
End Sub

Private Sub Cbo_Cta_Cte_Click()
    Id_CtaCte = Empty
    Id_CtaCte = Trim(fc_CodigoComboBox(Cbo_Cta_Cte, 2))
    If Id_CtaCte = "99" Then Id_CtaCte = Empty
    Call Trae_Informacion
End Sub

Private Sub Cbo_Tipo_Moneda_Click()
    T_Moneda = Empty
    T_Moneda = Trim(fc_CodigoComboBox(Cbo_Tipo_Moneda, 2))
    If T_Moneda = "99" Then T_Moneda = Empty Else T_Moneda = Mid(Cbo_Tipo_Moneda.Text, 2, 3)
    Call Trae_Informacion
End Sub

Private Sub Chk_01_Click()
    Dim mChecked As Boolean
    mChecked = Chk_01.Value
    dtp_Fecha.Value = Date
    txt_T_Cambio.Text = Empty
    dtp_Fecha.Enabled = Not mChecked
    txt_T_Cambio.Enabled = Not mChecked
    frm(0).Enabled = Not mChecked
End Sub

Private Sub Chk_Click()
    Dim mChecked As Boolean
    mChecked = Chk.Value
        If rsTemporal.RecordCount > 0 Then
            rsTemporal.MoveFirst
            Do While Not rsTemporal.EOF
                rsTemporal!estado = mChecked
                rsTemporal.MoveNext
            Loop
        End If
    grid.Refresh
End Sub

Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 0
    Me.Height = 7350
    Me.Width = 8220
    dtp_Fecha.Value = Date
    bBoolean = False
    Call mRecordSet
    Call Inicia_Form
End Sub
