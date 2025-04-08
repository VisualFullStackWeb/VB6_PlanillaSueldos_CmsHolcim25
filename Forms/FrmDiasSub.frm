VERSION 5.00
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmDiasSub 
   Caption         =   "Dias Subsidiados del Trabajador"
   ClientHeight    =   7470
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11115
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7470
   ScaleWidth      =   11115
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   120
      TabIndex        =   10
      Top             =   0
      Width           =   10935
      Begin VB.ComboBox CmbCia 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   240
         Width           =   7095
      End
      Begin VB.Label Label1 
         Caption         =   "Compañia"
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
         Index           =   3
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Datos Subsidio"
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
      Height          =   1575
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   10935
      Begin VB.CommandButton CmdFind 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2760
         TabIndex        =   2
         Top             =   720
         Width           =   400
      End
      Begin VB.TextBox TxtCodTrab 
         Height          =   285
         Left            =   1680
         TabIndex        =   1
         Top             =   720
         Width           =   975
      End
      Begin MSMask.MaskEdBox TxtPeriodo 
         Height          =   315
         Left            =   1680
         TabIndex        =   0
         Top             =   360
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   7
         Mask            =   "##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Caption         =   "Trabajador"
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
         TabIndex        =   9
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label LblNomTrab 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1680
         TabIndex        =   8
         Top             =   1080
         Width           =   9135
      End
      Begin VB.Label Label1 
         Caption         =   "Periodo"
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
         TabIndex        =   7
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Cod. Trabajador"
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
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4815
      Left            =   120
      TabIndex        =   4
      Top             =   2520
      Width           =   10935
      Begin TrueOleDBGrid70.TDBGrid Grd 
         Height          =   4335
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   7646
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Fecha Inicio Subsidio"
         Columns(0).DataField=   "fecini"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Fecha Final Subsidio"
         Columns(1).DataField=   "fecfin"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Nro CITT"
         Columns(2).DataField=   "nrocitt"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   49
         Columns(3)._MaxComboItems=   15
         Columns(3).Caption=   "Tipo de Suspension de la Relación Laboral"
         Columns(3).DataField=   "cod_suspension"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   4
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   13160660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=4"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2355"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2275"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=2302"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2223"
         Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(9)=   "Column(2).Width=3598"
         Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=3519"
         Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(13)=   "Column(3).Width=10821"
         Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=10742"
         Splits(0)._ColumnProps(16)=   "Column(3).Button=1"
         Splits(0)._ColumnProps(17)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(18)=   "Column(3).AutoDropDown=1"
         Splits(0)._ColumnProps(19)=   "Column(3).AutoCompletion=1"
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
         HeadLines       =   2
         FootLines       =   1
         Caption         =   "Dias Subsidiados"
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
         _StyleDefs(43)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(44)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(45)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(46)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(47)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
         _StyleDefs(48)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(49)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(50)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(51)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
         _StyleDefs(52)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
         _StyleDefs(53)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
         _StyleDefs(54)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
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
   End
End
Attribute VB_Name = "FrmDiasSub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SwNuevo As Boolean
Dim Sql As String
Dim Rs As New ADODB.Recordset

Private Sub CmdFind_Click()
Unload Frmgrdpla
Load Frmgrdpla
Frmgrdpla.Show vbModal
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
Me.Width = 11235
Me.Height = 7875
Call rCarCbo(CmbCia, Carga_Cia, "C", "00")
Call rUbiIndCmbBox(CmbCia, wcia, "00")
Crea_Rs
TxtPeriodo.Text = Format(Month(Date), "00") & "/" & Year(Date)
Sql = "select cod_maestro2,descrip from maestros_2 where right(ciamaestro,3)='151' and status<>'*' and rtrim(isnull(codsunat,'')) IN ('21','22') order by DESCRIP"
TrueDbgrid_CargarCombo Grd, 3, Sql, -1

End Sub

Public Sub Crea_Rs()
   'canteras
    If Rs.State = 1 Then Rs.Close
    Rs.Fields.Append "fecini", adChar, 10, adFldIsNullable
    Rs.Fields.Append "fecfin", adChar, 10, adFldIsNullable
    Rs.Fields.Append "nrocitt", adVarChar, 16, adFldIsNullable
    Rs.Fields.Append "cod_suspension", adChar, 2, adFldIsNullable
    Rs.Open
    Set Grd.DataSource = Rs

End Sub

Private Sub Grd_OnAddNew()
'Rs.AddNew
End Sub




Public Sub TrueDbgrid_CargarCombo(ByRef Tdbgrd As TrueOleDBGrid70.TDBGrid, ByVal Col As Integer, ByVal Strsql As String, ByVal default As Integer)
    Dim vItem As New TrueOleDBGrid70.ValueItem
    Dim vItems As TrueOleDBGrid70.ValueItems
    Dim AdoRs As New ADODB.Recordset
    
    Set vItems = Tdbgrd.Columns(Col).ValueItems
    AdoRs.CursorLocation = adUseClient
    AdoRs.Open Strsql, cn, adOpenForwardOnly, adLockReadOnly
    Set AdoRs.ActiveConnection = Nothing
    
    If Not AdoRs.EOF Then
        Do While Not AdoRs.EOF
           vItem.Value = Trim(AdoRs(0))
           vItem.DisplayValue = Trim(AdoRs(1))
           vItems.Add vItem
           AdoRs.MoveNext
        Loop
        AdoRs.Close
    End If
    Set AdoRs = Nothing
    If default <> -1 Then vItems.DefaultItem = default
End Sub

Private Sub TxtCodTrab_Change()
If Len(TxtCodTrab.Text) > 1 Then
    LblNomTrab.Caption = ""
    LimpiarRs Rs, Grd
End If
End Sub

Private Sub TxtCodTrab_GotFocus()
ResaltarTexto TxtCodTrab
End Sub

Private Sub TxtCodTrab_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{tab}"
End Sub

Private Sub TxtCodTrab_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TxtCodTrab_LostFocus()
If Trim(TxtCodTrab.Text) <> "" Then
    Dim Rq As ADODB.Recordset
    'OBTENER NOMBRE DE EMPLEADO
    Sql$ = Funciones.nombre()
    Sql$ = Sql$ & "codauxinterno,a.status,a.tipotrabajador,a.fingreso," & _
     "a.fcese,a.codafp,a.numafp,a.area,a.placod," & _
     "a.codauxinterno,b.descrip,a.tipotasaextra," & _
     "a.cargo,a.altitud,a.vacacion,a.area,a.fnacimiento," & _
     "a.fec_jubila,a.sindicato,a.ESSALUDVIDA,a.quinta " & _
     "from planillas a,maestros_2 b where a.status<>'*' "
     Sql$ = Sql$ & " AND right(b.ciamaestro,3)='055' "
     Sql$ = Sql$ & " and a.tipotrabajador=b.cod_maestro2 " _
     & "and cia='" & wcia & "' AND placod='" & Trim(TxtCodTrab.Text) & "' "
     Sql$ = Sql$ & " order by nombre"
     Screen.MousePointer = 11
     LimpiarRs Rs, Grd
     If fAbrRst(Rq, Sql$) Then
        If Not IsNull(Rq!fcese) Then
            MsgBox "El trabajador ya fue cesado", vbExclamation, Me.Caption
            LblNomTrab.Caption = ""
            TxtCodTrab.SetFocus
            GoTo Termina:
        End If
        TxtCodTrab.Text = UCase(TxtCodTrab.Text)
        LblNomTrab.Caption = Trim(Rq!nombre & "")
        Carga_Susidio
        If Rs.RecordCount = 0 Then Rs.AddNew
     Else
        MsgBox "No existe codigo de Trabajador"
        LblNomTrab.Caption = ""
        TxtCodTrab.SetFocus
        GoTo Termina:
     End If
     Grd.Col = 0
     Grd.SetFocus
Termina:
     Rq.Close
     Set Rq = Nothing
     Screen.MousePointer = 0
End If
End Sub

Private Sub TxtPeriodo_GotFocus()
ResaltarTexto TxtPeriodo
End Sub

Private Sub TxtPeriodo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{tab}"
End Sub


Public Sub Grabar()

'On Error GoTo ErrMsg:
Dim xLimiteIni As String
Dim xLimiteFin As String
xLimiteIni = "01/" & Left(TxtPeriodo.Text, 2) & "/" & Val(Right(TxtPeriodo.Text, 4))
xLimiteFin = fMaxDay(Val(Left(TxtPeriodo.Text, 2)), Val(Right(TxtPeriodo.Text, 4))) & "/" & Left(TxtPeriodo.Text, 2) & "/" & Val(Right(TxtPeriodo.Text, 4))

If Not IsNumeric(Right(TxtPeriodo.Text, 4)) Then
    MsgBox "El año del periodo no es correcto", vbExclamation, Me.Caption
    TxtPeriodo.SetFocus
    Exit Sub
ElseIf Not IsNumeric(Left(TxtPeriodo.Text, 2)) Then
    MsgBox "El mes del periodo no es correcto", vbExclamation, Me.Caption
    TxtPeriodo.SetFocus
    Exit Sub
ElseIf Trim(TxtCodTrab.Text) = "" Or Trim(Me.LblNomTrab.Caption) = "" Then
    MsgBox "Ingrese código de trabajador correctamente", vbExclamation, Me.Caption
    TxtCodTrab.SetFocus
    Exit Sub
ElseIf Rs.RecordCount = 0 Then
    MsgBox "Ingrese detalle de subsidios del trabajador", vbExclamation, Me.Caption
    Grd.SetFocus
    Exit Sub
End If
With Rs
    .MoveFirst
    Do While Not .EOF
        If Not IsDate(!FecIni) Then
            MsgBox "Ingrese fecha de inicio", vbExclamation, Me.Caption
            Grd.Col = 0
            Grd.SetFocus
            Exit Sub
'        ElseIf CDate(!FecIni) > Date Then
'            MsgBox "La fecha de inicio no puede ser mayor al día actual", vbExclamation, Me.Caption
'            Grd.Col = 0
'            Grd.SetFocus
'            Exit Sub
        ElseIf Not IsDate(!FecFin) Then
            MsgBox "Ingrese fecha final ", vbExclamation, Me.Caption
            Grd.Col = 1
            Grd.SetFocus
            Exit Sub
'        ElseIf CDate(!FecFin) > Date Then
'            MsgBox "La fecha de final no puede ser mayor al día actual", vbExclamation, Me.Caption
'            Grd.Col = 1
'            Grd.SetFocus
'            Exit Sub
'        ElseIf CDate(!FecIni) < CDate(xLimiteIni) Then
'            MsgBox "La fecha de inicio está fuera del rango del periodo", vbExclamation, Me.Caption
'            Grd.Col = 0
'            Grd.SetFocus
'            Exit Sub
'        ElseIf CDate(!FecFin) > CDate(xLimiteFin) Then
'            MsgBox "La fecha de final está fuera del rango del periodo", vbExclamation, Me.Caption
'            Grd.Col = 1
'            Grd.SetFocus
'            Exit Sub
        ElseIf CDate(!FecIni) > CDate(!FecFin) Then
            MsgBox "La fecha de Inicial no puede ser mayor a la fecha final", vbExclamation, Me.Caption
            Grd.Col = 0
            Grd.SetFocus
            Exit Sub
        ElseIf Trim(!nrocitt) = "" Then
            MsgBox "Ingrese Número de CITT asignado al trabajador correctamente", vbExclamation, Me.Caption
            Grd.Col = 2
            Grd.SetFocus
            Exit Sub
        ElseIf Trim(!cod_suspension & "") = "" Then
            MsgBox "Elija motivo de subsido del trabajador correctamente", vbExclamation, Me.Caption
            Grd.Col = 3
            Grd.SetFocus
            Exit Sub
        End If
        .MoveNext
    Loop
End With


If MsgBox("Desea Grabar los Datos ? ", vbQuestion + vbYesNo + vbDefaultButton1, Me.Caption) = vbNo Then Exit Sub
Screen.MousePointer = 11
Sql = "BEGIN TRANSACTION"
cn.Execute Sql, 64

Sql = "update dias_subsidio_trabajador set status='*',user_modi='" & Trim(wuser) & "',fec_modi=getdate() where cod_cia='" & wcia & "' and año=" & Right(TxtPeriodo.Text, 4) & " and mes=" & Left(TxtPeriodo.Text, 2) & " and cod_per='" & Trim(TxtCodTrab.Text) & "' AND status<>'*'"
cn.Execute Sql, 64

With Rs
    .MoveFirst
    Do While Not .EOF
            Sql = "insert into dias_subsidio_trabajador (cod_cia,año,mes,cod_per,fecini,fecfin,nrocitt,cod_tip_subsidio,status,user_crea,fec_crea,user_modi,fec_modi) "
            Sql = Sql & " values('" & wcia & "'," & Right(TxtPeriodo.Text, 4) & "," & Left(TxtPeriodo.Text, 2) & ",'" & Trim(TxtCodTrab.Text) & "','"
            Sql = Sql & Format(!FecIni, "mm/dd/yyyy") & "','" & Format(!FecFin, "mm/dd/yyyy") & "','" & Trim(!nrocitt) & "','" & Trim(!cod_suspension) & "','','" & Trim(wuser) & "',getdate(),null,null)"
            cn.Execute Sql, 64
        .MoveNext
    Loop
End With

Sql$ = "COMMIT TRANSACTION"
cn.Execute Sql, 64
Screen.MousePointer = 0
Nuevo

Exit Sub
ErrMsg:
    Screen.MousePointer = 11
    'ErrorLog "Se canceló la Grabación", True, Sql, Err & Space(1) & Err.Description
End Sub


Public Sub LimpiarRs(ByRef pRs As ADODB.Recordset, ByRef pDgrd As TrueOleDBGrid70.TDBGrid)
If pRs.State = 1 Then
    If pRs.RecordCount > 0 Then
        pRs.MoveFirst
        Do While Not pRs.EOF
            pRs.Delete
            pRs.MoveNext
        Loop
    End If
End If
pDgrd.Refresh
End Sub

Public Sub Nuevo()
TxtCodTrab.Text = ""
Me.LblNomTrab.Caption = ""
LimpiarRs Rs, Grd
TxtPeriodo.SetFocus
End Sub

Public Sub Elimimar()
'On Error GoTo ErrMsg:

If MsgBox("Desea Eliminar los Datos ? ", vbCritical + vbYesNo + vbDefaultButton1, Me.Caption) = vbNo Then Exit Sub
Screen.MousePointer = 11
Sql = "BEGIN TRANSACTION"
cn.Execute Sql, 64

Sql = "update dias_subsidio_trabajador set status='*',user_modi='" & Trim(wuser) & "',fec_modi=getdate() where cod_cia='" & wcia & "' and año=" & Right(TxtPeriodo.Text, 4) & " and mes=" & Left(TxtPeriodo.Text, 2) & " and cod_per='" & Trim(TxtCodTrab.Text) & "' and status<>'*'"
cn.Execute Sql, 64


Sql$ = "COMMIT TRANSACTION"
cn.Execute Sql, 64
Screen.MousePointer = 0
Nuevo

Exit Sub
ErrMsg:
    Screen.MousePointer = 11
    'ErrorLog "Se canceló la Grabación", True, Sql, Err & Space(1) & Err.Description
End Sub

Public Sub Carga_Susidio()
Dim Rq As ADODB.Recordset
Sql = "select * from dias_subsidio_trabajador where cod_cia='" & wcia & "' and año=" & Right(TxtPeriodo.Text, 4) & " and mes=" & Left(TxtPeriodo.Text, 2) & " and cod_per='" & Trim(TxtCodTrab.Text) & "' AND status<>'*'"
Screen.MousePointer = 11
Me.LimpiarRs Rs, Grd
If fAbrRst(Rq, Sql) Then
    Do While Not Rq.EOF
            Rs.AddNew
            Rs!FecIni = Format(Rq!FecIni, "dd/mm/yyyy")
            Rs!FecFin = Format(Rq!FecFin, "dd/mm/yyyy")
            Rs!nrocitt = Trim(Rq!nrocitt)
            Rs!cod_suspension = Rq!cod_tip_subsidio
        Rq.MoveNext
    Loop
    
End If
Screen.MousePointer = 0
Rq.Close
Set Rq = Nothing
End Sub
