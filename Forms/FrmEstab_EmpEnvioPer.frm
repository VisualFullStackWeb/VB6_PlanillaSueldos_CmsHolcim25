VERSION 5.00
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmEstab_EmpEnvioPer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Establecimientos"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5745
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   5745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdCancelar 
      BackColor       =   &H00FFFFFF&
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3600
      Width           =   1695
   End
   Begin VB.CommandButton CmdAceptar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Frame Frame2 
      Caption         =   "Declarante"
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
      Height          =   735
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   5535
      Begin VB.Label LblDecRazsoc 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   1
         Left            =   1560
         TabIndex        =   15
         Top             =   240
         Width           =   3855
      End
      Begin VB.Label LblDecRuc 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Centro de Riesgo"
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
      Height          =   2895
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Top             =   2760
      Width           =   3375
      Begin VB.OptionButton OptNo 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Caption         =   "No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   5
         Top             =   360
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton OptSi 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Caption         =   "Si"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   615
      End
      Begin TrueOleDBGrid70.TDBGrid DgrdTasa 
         Height          =   1935
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   3413
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Tasas (%)"
         Columns(0).DataField=   "porc"
         Columns(0).NumberFormat=   "##0.00"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   1
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   13160660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=1"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2355"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2275"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
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
         DefColWidth     =   0
         Enabled         =   0   'False
         HeadLines       =   2
         FootLines       =   1
         Caption         =   "Tasas  SCTR-ESSALUD"
         MultipleLines   =   0
         CellTipsWidth   =   0
         DeadAreaBackColor=   13160660
         RowDividerColor =   13160660
         RowSubDividerColor=   13160660
         DirectionAfterEnter=   2
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
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HE0E0E0&,.bold=0,.fontsize=825"
         _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
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
         _StyleDefs(43)  =   "Named:id=33:Normal"
         _StyleDefs(44)  =   ":id=33,.parent=0"
         _StyleDefs(45)  =   "Named:id=34:Heading"
         _StyleDefs(46)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(47)  =   ":id=34,.wraptext=-1"
         _StyleDefs(48)  =   "Named:id=35:Footing"
         _StyleDefs(49)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(50)  =   "Named:id=36:Selected"
         _StyleDefs(51)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(52)  =   "Named:id=37:Caption"
         _StyleDefs(53)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(54)  =   "Named:id=38:HighlightRow"
         _StyleDefs(55)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(56)  =   "Named:id=39:EvenRow"
         _StyleDefs(57)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(58)  =   "Named:id=40:OddRow"
         _StyleDefs(59)  =   ":id=40,.parent=33"
         _StyleDefs(60)  =   "Named:id=41:RecordSelector"
         _StyleDefs(61)  =   ":id=41,.parent=34"
         _StyleDefs(62)  =   "Named:id=42:FilterBar"
         _StyleDefs(63)  =   ":id=42,.parent=33"
      End
   End
   Begin VB.Frame Frame1 
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
      Height          =   1695
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   5535
      Begin VB.ComboBox CmbTipEst 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1200
         Width           =   5295
      End
      Begin VB.TextBox TxtDenominacion 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1200
         TabIndex        =   1
         Top             =   600
         Width           =   4215
      End
      Begin VB.TextBox TxtCodigo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   120
         MaxLength       =   4
         TabIndex        =   0
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo"
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
         Index           =   2
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Denominación"
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
         Index           =   1
         Left            =   1320
         TabIndex        =   11
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Código"
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
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   975
      End
   End
End
Attribute VB_Name = "FrmEstab_EmpEnvioPer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public RsTasa As New ADODB.Recordset
Dim vAccion As Integer

Private Sub CmbTipEst_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{tab}"
End Sub

Private Sub CmdAceptar_Click(Index As Integer)
If Trim(TxtCodigo.Text) = "" Then
    MsgBox "Ingrese Código del establecimiento", vbExclamation, Me.Caption
    TxtCodigo.SetFocus
    Exit Sub
ElseIf Len(Trim(TxtCodigo.Text)) > 4 Then
    MsgBox "Ingrese Código máx 4 caracteres.", vbExclamation, Me.Caption
    TxtCodigo.SetFocus
    Exit Sub
ElseIf Trim(TxtDenominacion.Text) = "" Then
    MsgBox "Ingrese Nombre de establecimiento", vbExclamation, Me.Caption
    TxtDenominacion.SetFocus
    Exit Sub
ElseIf Len(Trim(TxtDenominacion.Text)) > 40 Then
    MsgBox "Ingrese Denomincación máx 40 caracteres.", vbExclamation, Me.Caption
    TxtDenominacion.SetFocus
    Exit Sub
ElseIf Me.CmbTipEst.ListIndex = -1 Then
    MsgBox " Elija tipo de establecimiento", vbExclamation, Me.Caption
    CmbTipEst.SetFocus
    Exit Sub
ElseIf OptSi(0).Value = False And OptNo(1).Value = False Then
    MsgBox " Elija tipo de establecimiento", vbExclamation, Me.Caption
    CmbTipEst.SetFocus
    Exit Sub
ElseIf OptSi(0).Value = True And RsTasa.RecordCount = 0 Then
    MsgBox "Si el establecimiento es Centro de Riesgo debe ingresar por lo menos un (%)Tasa", vbExclamation, Me.Caption
    Me.DgrdTasa.SetFocus
    Exit Sub
End If

With RsTasa
    If .RecordCount > 0 Then
        .MoveFirst
        Do While Not .EOF
                If Not IsNumeric(!porc) Then
                    MsgBox "Ingrese porcentaje correctamente", vbExclamation, Me.Caption
                    Me.DgrdTasa.Col = 0
                    Me.DgrdTasa.SetFocus
                    Exit Sub
                ElseIf CCur(!porc) < 0 Or CCur(!porc) > 100 Then
                    MsgBox "Ingrese porcentaje correctamente (0-100%)", vbExclamation, Me.Caption
                    Me.DgrdTasa.Col = 0
                    Me.DgrdTasa.SetFocus
                    Exit Sub
                End If
            .MoveNext
        Loop
    End If
End With

Select Case vAccion
    Case 1 'nuevo
        If FrmEmpEnvioPer.RsEstablecimientos.RecordCount > 0 Then
            FrmEmpEnvioPer.RsEstablecimientos.MoveFirst
            FrmEmpEnvioPer.RsEstablecimientos.FIND "CODEST='" & Trim(TxtCodigo.Text) & "'"
            If Not FrmEmpEnvioPer.RsEstablecimientos.EOF Then
                MsgBox "El código de establecimiento ya fue ingresado", vbExclamation, Me.Caption
                TxtCodigo.SetFocus
                Exit Sub
            End If
        End If

        FrmEmpEnvioPer.RsEstablecimientos.AddNew
End Select
FrmEmpEnvioPer.RsEstablecimientos!codest = Trim(Me.TxtCodigo.Text)
FrmEmpEnvioPer.RsEstablecimientos!nomest = Trim(Me.TxtDenominacion.Text)
FrmEmpEnvioPer.RsEstablecimientos!tipest = fc_CodigoComboBox(CmbTipEst, 2)
FrmEmpEnvioPer.RsEstablecimientos!nomtipest = CmbTipEst.Text
FrmEmpEnvioPer.RsEstablecimientos!centro_riesgo = IIf(OptSi(0).Value = True, True, False)

With RsTasa
    
        If FrmEmpEnvioPer.RsEstTasa.RecordCount > 0 Then FrmEmpEnvioPer.RsEstTasa.MoveFirst
        Do While Not FrmEmpEnvioPer.RsEstTasa.EOF
            If Trim(FrmEmpEnvioPer.RsEstTasa!codest) = Trim(Me.TxtCodigo.Text) Then
                FrmEmpEnvioPer.RsEstTasa.Delete
            End If
           FrmEmpEnvioPer.RsEstTasa.MoveNext
        Loop
    If .RecordCount > 0 Then
        .MoveFirst
        Do While Not .EOF
                FrmEmpEnvioPer.RsEstTasa.AddNew
                FrmEmpEnvioPer.RsEstTasa!codest = Trim(Me.TxtCodigo.Text)
                FrmEmpEnvioPer.RsEstTasa!porc = Format(!porc, "##0.00")
            .MoveNext
        Loop
    End If
End With
Select Case vAccion
Case 1 'NUEVO
    Limpiar
    Me.TxtCodigo.Text = genera_correlativo
    TxtCodigo.Enabled = True
    Me.TxtDenominacion.SetFocus
Case 2, 3
    Unload Me
End Select
End Sub

Private Sub CmdCancelar_Click(Index As Integer)
Unload Me
End Sub

Private Sub DgrdTasa_OnAddNew()
'Me.RsTasa.AddNew
End Sub

Private Sub Form_Load()
Crea_Rs
Call fc_Descrip_Maestros2(wcia & "138", "", CmbTipEst, True)
Select Case vAccion
Case 1 'NUEVO
    Limpiar
    'TxtCodigo.BackColor = vbWhite
    'TxtCodigo.Enabled = True
    TxtCodigo.Text = Me.genera_correlativo
    
    TxtCodigo.Enabled = False
    TxtCodigo.BackColor = &HE0E0E0
    
Case 2 'MODIFICA
    TxtCodigo.Enabled = False
    TxtCodigo.BackColor = &HE0E0E0

End Select
End Sub

Private Sub Crea_Rs()
    'Tasa
    If RsTasa.State = 1 Then RsTasa.Close
    RsTasa.Fields.Append "codest", adChar, 2, adFldIsNullable
    RsTasa.Fields.Append "porc", adChar, 6, adFldIsNullable
    RsTasa.Open
    Set DgrdTasa.DataSource = RsTasa
    
End Sub

Private Sub Grd_OnAddNew()
RsTasa.AddNew
End Sub

Private Sub Option1_Click(Index As Integer)


End Sub

Private Sub OptNo_Click(Index As Integer)
DgrdTasa.Enabled = False
DgrdTasa.BackColor = &HC0C0C0
If RsTasa.RecordCount > 0 Then RsTasa.MoveFirst
Do While Not RsTasa.EOF
   RsTasa.Delete
   RsTasa.MoveNext
Loop

End Sub

Private Sub OptNo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{tab}"
End Sub

Private Sub OptSi_Click(Index As Integer)
DgrdTasa.Enabled = True
DgrdTasa.BackColor = vbWhite
If RsTasa.RecordCount = 0 Then RsTasa.AddNew
DgrdTasa.Col = 0
If DgrdTasa.Visible Then DgrdTasa.SetFocus
End Sub

Private Sub OptSi_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{tab}"
End Sub

Private Sub TxtCodigo_GotFocus()
ResaltarTexto TxtCodigo
End Sub

Private Sub TxtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{tab}"
End Sub

Private Sub Txtcodigo_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TxtDenominacion_GotFocus()
ResaltarTexto TxtDenominacion
End Sub

Private Sub TxtDenominacion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{tab}"
End Sub

Public Property Get MantAccion() As Variant
    MantAccion = vAccion
End Property

Public Property Let MantAccion(ByVal vNewValue As Variant)
    vAccion = vNewValue
End Property

Private Sub TxtDenominacion_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Public Sub Limpiar()
    Me.TxtCodigo.Text = ""
    Me.TxtDenominacion.Text = ""
    Me.OptSi(0).Value = False
    Me.OptNo(1).Value = True
    Me.CmbTipEst.ListIndex = -1
    If RsTasa.RecordCount > 0 Then RsTasa.MoveFirst
    Do While Not RsTasa.EOF
       RsTasa.Delete
       RsTasa.MoveNext
    Loop
End Sub

Public Function genera_correlativo() As String
If FrmEmpEnvioPer.RsEstablecimientos.RecordCount > 0 Then
    FrmEmpEnvioPer.RsEstablecimientos.Sort = "CODEST"
    FrmEmpEnvioPer.RsEstablecimientos.MoveLast
    genera_correlativo = "9" + Format(Val(Right(FrmEmpEnvioPer.RsEstablecimientos!codest, 3)) + 1, "000")
    
Else
    genera_correlativo = "9001"
End If
End Function
