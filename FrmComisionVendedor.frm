VERSION 5.00
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmComisionVendedor 
   Caption         =   "Asignar Vendedor Para Comision"
   ClientHeight    =   7470
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11115
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7470
   ScaleWidth      =   11115
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
      Height          =   1935
      Left            =   0
      TabIndex        =   5
      Top             =   840
      Width           =   10935
      Begin VB.TextBox TxtObjetivo 
         Height          =   285
         Left            =   1680
         TabIndex        =   14
         Top             =   1440
         Width           =   1935
      End
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
         TabIndex        =   8
         Top             =   240
         Width           =   400
      End
      Begin VB.TextBox TxtCodTrab 
         Height          =   285
         Left            =   1680
         TabIndex        =   7
         Top             =   225
         Width           =   975
      End
      Begin VB.ComboBox CmbCategoria 
         Height          =   315
         ItemData        =   "FrmComisionVendedor.frx":0000
         Left            =   1680
         List            =   "FrmComisionVendedor.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1020
         Width           =   2175
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Objetivo"
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
         Index           =   4
         Left            =   120
         TabIndex        =   13
         Top             =   1440
         Width           =   720
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
         TabIndex        =   12
         Top             =   660
         Width           =   1215
      End
      Begin VB.Label LblNomTrab 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1680
         TabIndex        =   11
         Top             =   600
         Width           =   9135
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
         TabIndex        =   10
         Top             =   240
         Width           =   1455
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
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   1080
         Width           =   825
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4335
      Left            =   0
      TabIndex        =   3
      Top             =   2880
      Width           =   10935
      Begin TrueOleDBGrid70.TDBGrid Grd 
         Height          =   3975
         Left            =   0
         TabIndex        =   4
         Top             =   240
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   7011
         _LayoutType     =   0
         _RowHeight      =   -2147483647
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
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   13160660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=2"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
         Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowUpdate     =   0   'False
         DefColWidth     =   0
         HeadLines       =   2
         FootLines       =   1
         Caption         =   "Vendeores Asignados"
         AllowArrows     =   0   'False
         MultipleLines   =   0
         CellTipsWidth   =   0
         MultiSelect     =   2
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
   End
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10935
      Begin VB.ComboBox CmbCia 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   1
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
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
   End
End
Attribute VB_Name = "FrmComisionVendedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SwNuevo As Boolean
Dim Sql As String
Dim rs As New ADODB.Recordset

Private Sub CmdFind_Click()
    Unload Frmgrdpla
    Load Frmgrdpla
    Frmgrdpla.Show vbModal
End Sub
Private Sub Cargar_CboCategoria()
   Call fc_Descrip_Maestros2("01185", "", CmbCategoria, False)
End Sub
Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 0
    Me.Width = 11235
    Me.Height = 7875
    Call Cargar_CboCategoria
    Call rCarCbo(CmbCia, Carga_Cia, "C", "00")
    Call rUbiIndCmbBox(CmbCia, wcia, "00")
    Call Carga_Susidio
   
End Sub
Public Sub Crea_Rs()
    If rs.State = 1 Then rs.Close
    rs.Fields.Append "Codigo", adVarChar, 20, adFldIsNullable
    rs.Fields.Append "Trabajador", adVarChar, 1000, adFldIsNullable
    rs.Fields.Append "Categoria", adVarChar, 10, adFldIsNullable
    rs.Fields.Append "Objetivo", adDecimal, 40, adFldIsNullable
    rs.Open
    Set Grd.DataSource = rs
End Sub
Private Sub Grd_DblClick()
    If Grd.ApproxCount > 0 Then
        Dim VCategoria As String
        VCategoria = Trim(Grd.Columns(2).Value)
        TxtCodTrab.Text = Trim(Grd.Columns(0).Value)
        LblNomTrab.Caption = Trim(Grd.Columns(1).Value)
        'CmbCategoria = VCategoria
        TxtObjetivo.Text = Trim(Grd.Columns(3).Value & "")
    End If
End Sub

Private Sub TxtCodTrab_Change()
    If Len(TxtCodTrab.Text) > 1 Then
        LblNomTrab.Caption = ""
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

Private Sub Txtcodtrab_LostFocus()
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
         LimpiarRs rs, Grd
         If fAbrRst(Rq, Sql$) Then
            If Not IsNull(Rq!fcese) Then
                MsgBox "El trabajador ya fue cesado", vbExclamation, Me.Caption
                LblNomTrab.Caption = ""
                TxtCodTrab.SetFocus
                GoTo Termina:
            End If
            TxtCodTrab.Text = UCase(TxtCodTrab.Text)
            LblNomTrab.Caption = Trim(Rq!nombre & "")
            If rs.RecordCount = 0 Then rs.AddNew
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


Public Sub LimpiarRs(ByRef pRs As ADODB.Recordset, ByRef pDgrd As TrueOleDBGrid70.Tdbgrid)
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
TxtObjetivo.Text = 0#
Me.LblNomTrab.Caption = ""
'LimpiarRs rs, Grd
End Sub

Public Sub Elimimar()
   
    If MsgBox("Desea Eliminar los Datos ? ", vbCritical + vbYesNo + vbDefaultButton1, Me.Caption) = vbNo Then Exit Sub
    Screen.MousePointer = 11
    Sql = "BEGIN TRANSACTION"
    cn.Execute Sql, 64
    
    Sql = "update Vta_ComisionVendedor set status='*' where cia='" & wcia & "'  and codigo='" & Trim(TxtCodTrab.Text) & "' and status<>'*'"
    cn.Execute Sql, 64
        
    Sql$ = "COMMIT TRANSACTION"
    cn.Execute Sql, 64
    Screen.MousePointer = 0
    Nuevo
    
    Exit Sub
ErrMsg:
        Screen.MousePointer = 11
End Sub
Public Sub Carga_Susidio()
Dim Rq As ADODB.Recordset
Sql = "select * from vta_ComisionVendedor where cia='" & wcia & "'  and status<>'*'"
Call Crea_Rs
'and codigo='" & Trim(TxtCodTrab.Text) & "' AND status<>'*'"
Screen.MousePointer = 11
Me.LimpiarRs rs, Grd
If fAbrRst(Rq, Sql) Then
    Do While Not Rq.EOF
            rs.AddNew
            rs!Codigo = UCase(Trim(Rq!Codigo))
            rs!Trabajador = UCase(Trim(Rq!Trabajador))
            rs!Categoria = Trim(Rq!Categoria)
            rs!Objetivo = Rq!Objetivo
        Rq.MoveNext
    Loop
End If
Set Grd.DataSource = rs
Screen.MousePointer = 0
Rq.Close
Set Rq = Nothing
End Sub
Public Sub Grabar()
    Dim xLimiteIni As String
    Dim xLimiteFin As String
    Dim VCategoria As String
    VCategoria = fc_CodigoComboBox(CmbCategoria, 2)
    If Not IsNumeric(TxtObjetivo.Text) Then
        MsgBox "El Importe del Obejtivo no es el Correcto", vbExclamation, Me.Caption
        TxtObjetivo.SetFocus
        Exit Sub
    ElseIf Trim(CmbCategoria.Text) = "" Then
        MsgBox "Asigne la Categoria Correcta al Vendedor", vbExclamation, Me.Caption
        CmbCategoria.SetFocus
        Exit Sub
    ElseIf Trim(TxtCodTrab.Text) = "" Or Trim(Me.LblNomTrab.Caption) = "" Then
        MsgBox "Ingrese código de trabajador correctamente", vbExclamation, Me.Caption
        TxtCodTrab.SetFocus
        Exit Sub
    End If
  '  With rs
  '      .MoveFirst
  '      Do While Not .EOF
            'If Not IsNumeric(!Objetivo) Then
            If Not IsNumeric(TxtObjetivo.Text) Then
                MsgBox "Importe Objetivo no es Correcto", vbExclamation, Me.Caption
                Grd.Col = 1
                Grd.SetFocus
                Exit Sub
            'ElseIf Trim(!Categoria) = "" Then
            ElseIf Trim(CmbCategoria.Text) = "" Then
                MsgBox "Ingrese Categoria Correctamente", vbExclamation, Me.Caption
                Grd.Col = 2
                Grd.SetFocus
                Exit Sub
            End If
  '          .MoveNext
  '      Loop
  '  End With
    
    If MsgBox("Desea Grabar los Datos ? ", vbQuestion + vbYesNo + vbDefaultButton1, Me.Caption) = vbNo Then Exit Sub
    Screen.MousePointer = 11
    Sql = "BEGIN TRANSACTION"
    cn.Execute Sql, 64
    
    Sql = "update Vta_ComisionVendedor set status='*' where cia='" & wcia & "' and codigo='" & Trim(TxtCodTrab.Text) & "' AND status<>'*'"
    cn.Execute Sql, 64
    
   ' With rs
   '     .MoveFirst
   '     Do While Not .EOF
                Sql = "insert into Vta_ComisionVendedor (Cia,codigo,Trabajador,Categoria,Objetivo,Status) "
                Sql = Sql & " values('" & wcia & "','" & Trim(TxtCodTrab.Text) & "','"
                Sql = Sql & LblNomTrab.Caption & "','" & VCategoria & "'," & TxtObjetivo.Text & ",'')"
                cn.Execute Sql, 64
   '         .MoveNext
   '     Loop
   ' End With
    
    Sql$ = "COMMIT TRANSACTION"
    cn.Execute Sql, 64
    Screen.MousePointer = 0
    
    Call Carga_Susidio

    Nuevo
    Exit Sub
ErrMsg:
        Screen.MousePointer = 11
End Sub

