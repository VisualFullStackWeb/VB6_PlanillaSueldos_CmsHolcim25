VERSION 5.00
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmCategoriaVendedor 
   Caption         =   "Categoria de Ventas"
   ClientHeight    =   3750
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   23565
   LinkTopic       =   "Categoria de Ventas"
   MDIChild        =   -1  'True
   ScaleHeight     =   3750
   ScaleWidth      =   23565
   Begin VB.ComboBox CmbCategoria 
      Height          =   315
      ItemData        =   "FrmCategoriaVendedor.frx":0000
      Left            =   1080
      List            =   "FrmCategoriaVendedor.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   5535
   End
   Begin TrueOleDBGrid70.TDBGrid GrdCategoriaVendedor 
      Height          =   2895
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   23535
      _ExtentX        =   41513
      _ExtentY        =   5106
      _LayoutType     =   0
      _RowHeight      =   17
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   2
      Splits(0)._UserFlags=   0
      Splits(0).MarqueeStyle=   3
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).FetchRowStyle=   -1  'True
      Splits(0).DividerColor=   13160660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=2"
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
      _StyleDefs(14)  =   ":id=2,.bold=-1,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(15)  =   ":id=2,.fontname=Arial"
      _StyleDefs(16)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(17)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(18)  =   ":id=3,.fontname=MS Sans Serif"
      _StyleDefs(19)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(20)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36,.fgcolor=&H80000007&"
      _StyleDefs(21)  =   "EditorStyle:id=7,.parent=1,.bgcolor=&HFFECD9&,.fgcolor=&H0&"
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Categoria"
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
      TabIndex        =   2
      Top             =   165
      Width           =   825
   End
End
Attribute VB_Name = "FrmCategoriaVendedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsCategoria As New ADODB.Recordset
Dim VCategoria As String
Private Sub Crea_Rs()
    If rsCategoria.State = 1 Then rsCategoria.Close
        rsCategoria.Fields.Append "F80", adDouble, 6, adFldIsNullable
        rsCategoria.Fields.Append "F85", adDouble, 6, adFldIsNullable
        rsCategoria.Fields.Append "F90", adDouble, 6, adFldIsNullable
        rsCategoria.Fields.Append "F95", adDouble, 6, adFldIsNullable
        rsCategoria.Fields.Append "F100", adDouble, 6, adFldIsNullable
        rsCategoria.Fields.Append "F110", adDouble, 6, adFldIsNullable
        rsCategoria.Fields.Append "F120", adDouble, 6, adFldIsNullable
        rsCategoria.Fields.Append "F130", adDouble, 6, adFldIsNullable
        rsCategoria.Fields.Append "F140", adDouble, 6, adFldIsNullable
        rsCategoria.Fields.Append "F150", adDouble, 6, adFldIsNullable
        rsCategoria.Fields.Append "F160", adDouble, 6, adFldIsNullable
        rsCategoria.Fields.Append "F170", adDouble, 6, adFldIsNullable
        rsCategoria.Fields.Append "F180", adDouble, 6, adFldIsNullable
        rsCategoria.Fields.Append "F190", adDouble, 4, adFldIsNullable
        rsCategoria.Fields.Append "F200", adDouble, 4, adFldIsNullable
        
        rsCategoria.Open
        Set GrdCategoriaVendedor.DataSource = rsCategoria
End Sub
Public Function fAbrRst(ByRef pRs As ADODB.Recordset, ByVal pSql As String) As Boolean
     cn.CommandTimeout = 0
     Set pRs = cn.Execute(pSql)
     fAbrRst = Not pRs.EOF
End Function
Private Sub Cargar_Categorias()
    Call Crea_Rs
    VCategoria = fc_CodigoComboBox(CmbCategoria, 2)
    Sql$ = "usp_Listar_Vta_CategoriaVendedor '" & wcia & "'," & "'" & VCategoria & "'"
    If fAbrRst(rs, Sql$) Then
        rs.MoveFirst
        Do While Not rs.EOF
            rsCategoria.AddNew
            rsCategoria!F80 = rs!F80
            rsCategoria!F85 = rs!F85
            rsCategoria!F90 = rs!F90
            rsCategoria!F95 = rs!F95
            rsCategoria!F100 = rs!F100
            rsCategoria!F110 = rs!F110
            rsCategoria!F120 = rs!F120
            rsCategoria!F130 = rs!F130
            rsCategoria!F140 = rs!F140
            rsCategoria!F150 = rs!F150
            rsCategoria!F160 = rs!F160
            rsCategoria!F170 = rs!F170
            rsCategoria!F180 = rs!F180
            rsCategoria!F190 = rs!F190
            rsCategoria!F200 = rs!F200
 
            rsCategoria.Update
            rs.MoveNext
        Loop
        rsCategoria.MoveFirst
    End If
    'Set GrdCategoriaVendedor.DataSource = rsCategoria
    rs.Close
    
    Set rs = Nothing
End Sub

Private Sub Cargar_CboCategoria()
   Call fc_Descrip_Maestros2("01185", "", CmbCategoria, False)
End Sub

Private Sub CmbCategoria_Click()
If CmbCategoria.Text <> "" Then
   Set GrdCategoriaVendedor.DataSource = Nothing
   Call Cargar_Categorias
End If
End Sub

Private Sub Form_Load()
Me.Top = 0: Me.Left = 0
Me.Width = 24825: Me.Height = 4260
Crea_Rs
Call Cargar_CboCategoria
End Sub


Public Sub Graba_Categoria()
Dim Sql As String
Dim NroTrans As Integer
On Error GoTo Salir
NroTrans = 0
If cn.State = True Then cn.Close
If CmbCategoria.ListIndex < 0 Then MsgBox "Seleccione La Categoria", vbInformation: Exit Sub
If rsCategoria.RecordCount <= 0 Then MsgBox "Ingrese al menos un Registro a la Categoria, vbInformation: Exit Sub"
cn.BeginTrans
NroTrans = 0

VCategoria = fc_CodigoComboBox(CmbCategoria, 2)

'Sql = "usp_pla_PlaBcoCta 0,'" & wcia & "','" & VBcoPago & "','','','','','" & wuser & "',3"
'cn.Execute Sql, 64

rsCategoria.MoveFirst
Do While Not rsCategoria.EOF
    If IsNull(rsCategoria!F80) = True Or IsNumeric(rsCategoria!F80) = False Then
      NroTrans = 1
      GoTo Salir
    End If
    If IsNull(rsCategoria!F85) = True Or IsNumeric(rsCategoria!F85) = False Then
      NroTrans = 2
      GoTo Salir
    End If
    If IsNull(rsCategoria!F90) = True Or IsNumeric(rsCategoria!F90) = False Then
      NroTrans = 3
      GoTo Salir
    End If
    If IsNull(rsCategoria!F95) = True Or IsNumeric(rsCategoria!F95) = False Then
      NroTrans = 4
      GoTo Salir
    End If
    If IsNull(rsCategoria!F100) = True Or IsNumeric(rsCategoria!F100) = False Then
      NroTrans = 5
      GoTo Salir
    End If
    If IsNull(rsCategoria!F110) = True Or IsNumeric(rsCategoria!F110) = False Then
      NroTrans = 6
      GoTo Salir
    End If
    If IsNull(rsCategoria!F120) = True Or IsNumeric(rsCategoria!F120) = False Then
      NroTrans = 7
      GoTo Salir
    End If
    If IsNull(rsCategoria!F130) = True Or IsNumeric(rsCategoria!F130) = False Then
      NroTrans = 8
      GoTo Salir
    End If
    If IsNull(rsCategoria!F140) = True Or IsNumeric(rsCategoria!F140) = False Then
      NroTrans = 9
      GoTo Salir
    End If
    If IsNull(rsCategoria!F150) = True Or IsNumeric(rsCategoria!F150) = False Then
      NroTrans = 10
      GoTo Salir
    End If
    If IsNull(rsCategoria!F160) = True Or IsNumeric(rsCategoria!F160) = False Then
      NroTrans = 11
      GoTo Salir
    End If
    If IsNull(rsCategoria!F170) = True Or IsNumeric(rsCategoria!F170) = False Then
      NroTrans = 12
      GoTo Salir
    End If
    If IsNull(rsCategoria!F180) = True Or IsNumeric(rsCategoria!F180) = False Then
      NroTrans = 13
      GoTo Salir
    End If
    If IsNull(rsCategoria!F190) = True Or IsNumeric(rsCategoria!F190) = False Then
      NroTrans = 14
      GoTo Salir
    End If
    If IsNull(rsCategoria!F200) = True Or IsNumeric(rsCategoria!F200) = False Then
      NroTrans = 15
      GoTo Salir
    End If
    
    Sql = "usp_Inserta_Vta_CategoriaVendedor '" & wcia & "','" & VCategoria & "'," & Trim(rsCategoria!F80) & "," & Trim(rsCategoria!F85) & "," & Trim(rsCategoria!F90) & "," & Trim(rsCategoria!F95) & "," & Trim(rsCategoria!F100) & "," & Trim(rsCategoria!F110) & "," & Trim(rsCategoria!F120) & "," & Trim(rsCategoria!F130) & "," & Trim(rsCategoria!F140) & "," & Trim(rsCategoria!F150) & "," & Trim(rsCategoria!F160) & "," & Trim(rsCategoria!F170) & "," & Trim(rsCategoria!F180) & "," & Trim(rsCategoria!F190) & "," & Trim(rsCategoria!F200)
    cn.Execute Sql, 64
    rsCategoria.MoveNext
Loop

cn.CommitTrans

MsgBox "Grabación Satisfactoria", vbInformation



Call Cargar_Categorias
Exit Sub

Salir:

If NroTrans > 0 Then
    cn.RollbackTrans
End If
Select Case NroTrans
Case 1
    MsgBox "Ingrese Factor para el 80 %", vbCritical, Me.Caption
Case 2
    MsgBox "Ingrese Factor para el 85 %", vbCritical, Me.Caption
Case 3
    MsgBox "Ingrese Factor para el 90 %", vbCritical, Me.Caption
Case 4
    MsgBox "Ingrese Factor para el 95 %", vbCritical, Me.Caption
Case 5
    MsgBox "Ingrese Factor para el 100 %", vbCritical, Me.Caption
Case 6
    MsgBox "Ingrese Factor para el 110 %", vbCritical, Me.Caption
Case 7
    MsgBox "Ingrese Factor para el 120 %", vbCritical, Me.Caption
Case 8
    MsgBox "Ingrese Factor para el 130 %", vbCritical, Me.Caption
Case 9
    MsgBox "Ingrese Factor para el 140 %", vbCritical, Me.Caption
Case 10
    MsgBox "Ingrese Factor para el 150 %", vbCritical, Me.Caption
Case 11
    MsgBox "Ingrese Factor para el 160 %", vbCritical, Me.Caption
Case 12
    MsgBox "Ingrese Factor para el 170 %", vbCritical, Me.Caption
Case 13
    MsgBox "Ingrese Factor para el 180 %", vbCritical, Me.Caption
Case 14
    MsgBox "Ingrese Factor para el 190 %", vbCritical, Me.Caption
Case 15
    MsgBox "Ingrese Factor para el 2000 %", vbCritical, Me.Caption
End Select

End Sub


