VERSION 5.00
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form Frmseteoimpresion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seteo de Impresion de Boletas"
   ClientHeight    =   7035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6750
   Icon            =   "Frmseteoimpresion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   6750
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkTodos 
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   5040
      TabIndex        =   8
      Top             =   1560
      Width           =   255
   End
   Begin VB.ComboBox Cmbtipo 
      Height          =   315
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   960
      Width           =   5415
   End
   Begin VB.ComboBox Cmbconcepto 
      Height          =   315
      ItemData        =   "Frmseteoimpresion.frx":030A
      Left            =   1200
      List            =   "Frmseteoimpresion.frx":0317
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   600
      Width           =   5415
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6615
      Begin VB.ComboBox Cmbcia 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   80
         Width           =   5295
      End
      Begin VB.Label Label2 
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
         ForeColor       =   &H8000000E&
         Height          =   210
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   825
      End
   End
   Begin TrueOleDBGrid70.TDBGrid GrdSeteo 
      Height          =   5535
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   9763
      _LayoutType     =   4
      _RowHeight      =   17
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Descripción"
      Columns(0).DataField=   "concepto"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   4
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Sel"
      Columns(1).DataField=   "chk"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Codigo"
      Columns(2).DataField=   "codigo"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Orden"
      Columns(3).DataField=   "orden"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   4
      Splits(0)._UserFlags=   0
      Splits(0).MarqueeStyle=   3
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).FetchRowStyle=   -1  'True
      Splits(0).DividerColor=   13160660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=4"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=7938"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=7858"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=873"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=794"
      Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(10)=   "Column(1)._ColStyle=1"
      Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(12)=   "Column(2).Width=2117"
      Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=2037"
      Splits(0)._ColumnProps(15)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(16)=   "Column(2)._ColStyle=8193"
      Splits(0)._ColumnProps(17)=   "Column(2).Visible=0"
      Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(19)=   "Column(3).Width=1402"
      Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=1323"
      Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=0"
      Splits(0)._ColumnProps(24)=   "Column(3).WrapText=1"
      Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
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
      HeadLines       =   2
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      MultiSelect     =   0
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
      _StyleDefs(39)  =   "Splits(0).Columns(0).Style:id=54,.parent=13"
      _StyleDefs(40)  =   "Splits(0).Columns(0).HeadingStyle:id=51,.parent=14"
      _StyleDefs(41)  =   "Splits(0).Columns(0).FooterStyle:id=52,.parent=15"
      _StyleDefs(42)  =   "Splits(0).Columns(0).EditorStyle:id=53,.parent=17"
      _StyleDefs(43)  =   "Splits(0).Columns(1).Style:id=46,.parent=13,.alignment=2"
      _StyleDefs(44)  =   "Splits(0).Columns(1).HeadingStyle:id=43,.parent=14"
      _StyleDefs(45)  =   "Splits(0).Columns(1).FooterStyle:id=44,.parent=15,.alignment=3"
      _StyleDefs(46)  =   "Splits(0).Columns(1).EditorStyle:id=45,.parent=17"
      _StyleDefs(47)  =   "Splits(0).Columns(2).Style:id=50,.parent=13,.alignment=2,.locked=-1"
      _StyleDefs(48)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=14"
      _StyleDefs(49)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=15"
      _StyleDefs(50)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=17"
      _StyleDefs(51)  =   "Splits(0).Columns(3).Style:id=28,.parent=13,.alignment=0,.wraptext=-1,.locked=0"
      _StyleDefs(52)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=14"
      _StyleDefs(53)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=15"
      _StyleDefs(54)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=17"
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
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "T. Trabajador"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Concepto"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   690
   End
End
Attribute VB_Name = "Frmseteoimpresion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsconcepto As New Recordset
Dim VConcepto As String
Dim VTipo As String
Dim VTipotrab As String
Dim RsSeteo As New ADODB.Recordset

Private Sub chkTodos_Click()
If RsSeteo.RecordCount > 0 Then
    RsSeteo.MoveFirst
    Do While Not RsSeteo.EOF
        RsSeteo!Chk = CBool(chkTodos.Value)
        RsSeteo.Update
        RsSeteo.MoveNext
    Loop
End If
End Sub

Private Sub Cmbcia_Click()
wcia = Trim(Right("00" & CmbCia.ItemData(CmbCia.ListIndex), 2))
Call fc_Descrip_Maestros2("01055", "", Cmbtipo)
End Sub

Private Sub Crea_Rs()

If RsSeteo.State = 1 Then RsSeteo.Close
    RsSeteo.Fields.Append "concepto", adVarChar, 500, adFldIsNullable
    RsSeteo.Fields.Append "chk", adBoolean, , adFldIsNullable
    RsSeteo.Fields.Append "codigo", adChar, 2, adFldIsNullable
    RsSeteo.Fields.Append "orden", adInteger, , adFldIsNullable
    
    RsSeteo.Open
    Set GrdSeteo.DataSource = RsSeteo
End Sub

Private Sub Cmbconcepto_Click()
If CmbConcepto.ListIndex = 0 Then
   VConcepto = "02"
   VTipo = "I"
Else
   VConcepto = "03"
   If CmbConcepto.ListIndex = 1 Then VTipo = "D" Else VTipo = "A"
End If
Procesa
End Sub

Private Sub CmbTipo_Click()
VTipotrab = fc_CodigoComboBox(Cmbtipo, 2)
Procesa
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
Me.Width = 6795
Me.Height = 7515
Crea_Rs
Call rCarCbo(CmbCia, Carga_Cia, "C", "00")
Call rUbiIndCmbBox(CmbCia, wcia, "00")
End Sub
Private Sub Procesa()

Call Crea_Rs
If CmbConcepto.ListIndex < 0 Then Exit Sub
If Cmbtipo.ListIndex < 0 Then Exit Sub

Sql$ = "SELECT pa.codinterno,pa.descripcion,CASE WHEN NOT pp.tipo IS NULL THEN 'I' ELSE '' END AS tipo,PP.ORDEN FROM" & _
    " placonstante pa LEFT OUTER JOIN plaseteoprint pp ON (pp.codigo=pa.codinterno and pp.cia='" & wcia & "' and pp.tipo='" & VTipo & "' " & _
    " and pp.tipo_trab='" & VTipotrab & "' and pp.status<>'*' ) WHERE pa.tipomovimiento='" & VConcepto & "' and pa.status<>'*'  AND pa.cia='" & wcia & "' " & _
    " ORDER BY pa.codinterno "

Set rs = cn.Execute(Sql)

If Not rs.EOF Then

    Do While Not rs.EOF
        RsSeteo.AddNew
        RsSeteo!concepto = rs!Descripcion
        RsSeteo!Chk = IIf(rs!tipo = "I", True, False)
        RsSeteo!Codigo = rs!codinterno
        If IsNull(rs!Orden) = True Then
            RsSeteo!Orden = 0
        Else
            RsSeteo!Orden = Val(rs!Orden)
        End If
    
        RsSeteo.Update
        rs.MoveNext
    Loop
        rs.Close
End If

Set rs = Nothing

End Sub

Public Sub Grabar_Seteo_Print()
Dim mTipo As String
Dim NroTrans As Integer
On Error GoTo ErrorTrans
NroTrans = 0
If CmbConcepto.ListIndex < 0 Then MsgBox "Debe Seleccionar Concepto", vbInformation, "Seteo de Impresion": CmbConcepto.SetFocus: Exit Sub
If Cmbtipo.ListIndex < 0 Then MsgBox "Debe Seleccionar Tipo de Trabajador", vbInformation, "Seteo de Impresion": Cmbtipo.SetFocus: Exit Sub
Mgrab = MsgBox("Seguro de Grabar Seteo", vbYesNo + vbQuestion, TitMsg)
If Mgrab <> 6 Then Exit Sub

Screen.MousePointer = vbArrowHourglass

'If rsconcepto.RecordCount > 0 Then rsconcepto.MoveFirst

Dim iFila As Long

cn.BeginTrans
NroTrans = 1

Sql$ = "update plaseteoprint set status='*' where cia='" & wcia & "' and tipo='" & VTipo & "' and status<>'*' and tipo_trab='" & VTipotrab & "'"
cn.Execute Sql$
   
Me.GrdSeteo.Update
GrdSeteo.BatchUpdates = True

If RsSeteo.EOF = False Then RsSeteo.Update
If RsSeteo.RecordCount > 0 Then RsSeteo.MoveFirst
Do While Not RsSeteo.EOF
   If RsSeteo!Chk = True Then
      Sql$ = "INSERT INTO plaseteoprint values('" & wcia & "','" & VTipo & "','" & RsSeteo!Codigo & "','','" & wuser & "'," & FechaSys & ",'" & VTipotrab & "','" & RsSeteo!Orden & "')"
      cn.Execute Sql$
   End If

    RsSeteo.MoveNext
Loop


cn.CommitTrans
Screen.MousePointer = vbDefault
MsgBox "Se guardarón los datos correctamente", vbInformation, Me.Caption
Call Procesa
Exit Sub
ErrorTrans:
If NroTrans = 1 Then
    cn.RollbackTrans
End If
MsgBox ERR.Description, vbCritical, Me.Caption
Screen.MousePointer = vbDefault



End Sub


