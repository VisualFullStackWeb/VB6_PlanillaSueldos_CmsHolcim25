VERSION 5.00
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmPlaBcoCta 
   Caption         =   "Banco y Cuenta de Pago de Haberes"
   ClientHeight    =   3750
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6705
   Icon            =   "FrmPlaBcoCta.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3750
   ScaleWidth      =   6705
   Begin VB.ComboBox CmbBcoPago 
      Height          =   315
      ItemData        =   "FrmPlaBcoCta.frx":030A
      Left            =   1080
      List            =   "FrmPlaBcoCta.frx":030C
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   240
      Width           =   5535
   End
   Begin TrueOleDBGrid70.TDBGrid GrdCuentaBco 
      Height          =   2895
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   5106
      _LayoutType     =   4
      _RowHeight      =   17
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   1
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Moneda"
      Columns(0).DataField=   "moneda"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Sucursal"
      Columns(1).DataField=   "Sucursal"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Nro. Cuenta"
      Columns(2).DataField=   "cuentabco"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "IdBcoCta"
      Columns(3).DataField=   "IdBcoCta"
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
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2117"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2037"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=1"
      Splits(0)._ColumnProps(6)=   "Column(0).Button=1"
      Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(8)=   "Column(1).Width=1773"
      Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=1693"
      Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(13)=   "Column(2).Width=6006"
      Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=5927"
      Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=0"
      Splits(0)._ColumnProps(18)=   "Column(2).WrapText=1"
      Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(20)=   "Column(3).Width=2725"
      Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=2646"
      Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=8196"
      Splits(0)._ColumnProps(25)=   "Column(3).Visible=0"
      Splits(0)._ColumnProps(26)=   "Column(3).Order=4"
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
      _StyleDefs(39)  =   "Splits(0).Columns(0).Style:id=50,.parent=13,.alignment=2,.locked=0"
      _StyleDefs(40)  =   "Splits(0).Columns(0).HeadingStyle:id=47,.parent=14"
      _StyleDefs(41)  =   "Splits(0).Columns(0).FooterStyle:id=48,.parent=15"
      _StyleDefs(42)  =   "Splits(0).Columns(0).EditorStyle:id=49,.parent=17"
      _StyleDefs(43)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
      _StyleDefs(44)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(45)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(46)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(47)  =   "Splits(0).Columns(2).Style:id=28,.parent=13,.alignment=0,.wraptext=-1,.locked=0"
      _StyleDefs(48)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
      _StyleDefs(49)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
      _StyleDefs(50)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
      _StyleDefs(51)  =   "Splits(0).Columns(3).Style:id=46,.parent=13,.locked=-1"
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
   Begin VB.Label Label1 
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
      Left            =   360
      TabIndex        =   1
      Top             =   285
      Width           =   555
   End
End
Attribute VB_Name = "FrmPlaBcoCta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsBcoCta As New ADODB.Recordset
Dim VBcoPago As String

Private Sub Crea_Rs()

If rsBcoCta.State = 1 Then rsBcoCta.Close
    rsBcoCta.Fields.Append "moneda", adChar, 3, adFldIsNullable
    rsBcoCta.Fields.Append "sucursal", adVarChar, 5, adFldIsNullable
    rsBcoCta.Fields.Append "cuentabco", adVarChar, 50, adFldIsNullable
    rsBcoCta.Fields.Append "IdBcoCta", adInteger, , adFldIsNullable
    rsBcoCta.Open
    Set GrdCuentaBco.DataSource = rsBcoCta
End Sub


Private Sub CmbBcoPago_Click()
Call Cargar_Cuenta_Banco
End Sub

Private Sub Form_Load()
Me.Top = 0: Me.Left = 0
Me.Width = 6825: Me.Height = 4260
Call fc_Descrip_Maestros2("01007", "", CmbBcoPago, False)
CmbBcoPago.ListIndex = 0
Call Cargar_Cuenta_Banco
Sql$ = "Select flag1 From maestros_2 where ciamaestro='" & wcia & "006' and status<>'*'"
If fAbrRst(rs, Sql$) Then
    rs.MoveFirst
    Do While Not rs.EOF
        Dim VALOR As New TrueOleDBGrid70.ValueItem
        VALOR.Value = Mid(Trim(rs!flag1), 2, 3)
        VALOR.DisplayValue = Mid(Trim(rs!flag1), 2, 3)
        GrdCuentaBco.Columns(0).ValueItems.Add VALOR
        rs.MoveNext
    Loop
End If
rs.Close
Set rs = Nothing

End Sub

Private Sub Cargar_Cuenta_Banco()
VBcoPago = fc_CodigoComboBox(CmbBcoPago, 2)
Call Crea_Rs
Sql$ = "usp_pla_PlaBcoCta 0,'" & wcia & "','" & VBcoPago & "','','','','','" & wuser & "',4"
If fAbrRst(rs, Sql$) Then
    rs.MoveFirst
    Do While Not rs.EOF
        rsBcoCta.AddNew
        rsBcoCta!moneda = rs!moneda
        rsBcoCta!Sucursal = rs!Sucursal
        rsBcoCta!cuentabco = rs!cuentabco
        rsBcoCta!idbcocta = rs!idbcocta
        rsBcoCta.Update
        rs.MoveNext
    Loop
    rsBcoCta.MoveFirst
End If
rs.Close
Set rs = Nothing

End Sub

Public Sub Graba_BcoCta()
Dim Sql As String
Dim NroTrans As Integer
On Error GoTo Salir
NroTrans = 0

If CmbBcoPago.ListIndex < 0 Then MsgBox "Seleccione el Banco", vbInformation: Exit Sub
If rsBcoCta.RecordCount <= 0 Then MsgBox "Ingrese al menos una cuenta bancaria", vbInformation: Exit Sub
cn.BeginTrans
NroTrans = 1

VBcoPago = fc_CodigoComboBox(CmbBcoPago, 2)

Sql = "usp_pla_PlaBcoCta 0,'" & wcia & "','" & VBcoPago & "','','','','','" & wuser & "',3"
cn.Execute Sql, 64

rsBcoCta.MoveFirst
Do While Not rsBcoCta.EOF
    If IsNull(rsBcoCta!moneda) = True Then
        NroTrans = 2
        GoTo Salir
    ElseIf Trim(rsBcoCta!moneda) = "" Then
        NroTrans = 2
        GoTo Salir
    End If
    If IsNull(rsBcoCta!cuentabco) = True Then
        NroTrans = 3
        GoTo Salir
    ElseIf Len(Trim(rsBcoCta!cuentabco)) < 10 Then
        NroTrans = 3
        GoTo Salir
    End If
    If IsNull(rsBcoCta!idbcocta) = True Then
        Sql = "usp_pla_PlaBcoCta 0,'" & wcia & "','" & VBcoPago & "','" & Trim(rsBcoCta!moneda) & "','" & Trim(rsBcoCta!Sucursal) & "','" & Trim(rsBcoCta!cuentabco) & "','','" & wuser & "',1"
    ElseIf rsBcoCta!idbcocta <= 0 Then
        Sql = "usp_pla_PlaBcoCta 0,'" & wcia & "','" & VBcoPago & "','" & Trim(rsBcoCta!moneda) & "','" & Trim(rsBcoCta!Sucursal) & "','" & Trim(rsBcoCta!cuentabco) & "','','" & wuser & "',1"
    Else
        Sql = "usp_pla_PlaBcoCta " & rsBcoCta!idbcocta & ",'" & wcia & "','" & VBcoPago & "','" & Trim(rsBcoCta!moneda) & "','" & Trim(rsBcoCta!Sucursal) & "','" & Trim(rsBcoCta!cuentabco) & "','','" & wuser & "',2"
    End If
    cn.Execute Sql, 64
    rsBcoCta.MoveNext
Loop

cn.CommitTrans

MsgBox "Grabación Satisfactoria", vbInformation

Call Cargar_Cuenta_Banco

Exit Sub
Salir:

If NroTrans > 0 Then
    cn.RollbackTrans
End If
If NroTrans = 2 Then
    MsgBox "Ingrese la Moneda", vbCritical, Me.Caption
ElseIf NroTrans = 3 Then
    MsgBox "Ingrese Correctamente Cuenta del Banco", vbCritical, Me.Caption
Else
    MsgBox ERR.Description, vbCritical, Me.Caption
End If

End Sub

