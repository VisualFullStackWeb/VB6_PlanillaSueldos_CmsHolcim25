VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frmpromedio 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "» Calculo de Promedio (platasaanexo) «"
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6450
   Icon            =   "FrmPromedio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   6450
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox Cmbcargo 
      Height          =   315
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   1200
      Visible         =   0   'False
      Width           =   5070
   End
   Begin VB.Frame Frame3 
      Height          =   4935
      Left            =   4215
      TabIndex        =   9
      Top             =   1560
      Width           =   2055
      Begin VB.TextBox txtfactor 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   540
         TabIndex        =   19
         Top             =   2205
         Width           =   495
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   285
         Left            =   990
         TabIndex        =   14
         Top             =   525
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.TextBox Txtperiodo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   510
         TabIndex        =   11
         Top             =   525
         Width           =   495
      End
      Begin MSComCtl2.UpDown UpDownCALC 
         Height          =   285
         Left            =   1020
         TabIndex        =   18
         Top             =   2205
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Factor Meses"
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
         Left            =   540
         TabIndex        =   20
         Top             =   1905
         Width           =   1155
      End
      Begin MSForms.OptionButton OpcMes 
         Height          =   375
         Left            =   480
         TabIndex        =   13
         Top             =   1290
         Width           =   1095
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   5
         Size            =   "1931;661"
         Value           =   "1"
         Caption         =   "Meses"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.OptionButton OpcSemana 
         Height          =   255
         Left            =   480
         TabIndex        =   12
         Top             =   1305
         Visible         =   0   'False
         Width           =   1335
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   5
         Size            =   "2355;450"
         Value           =   "0"
         Caption         =   "Semanas"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   525
         TabIndex        =   10
         Top             =   240
         Width           =   660
      End
   End
   Begin VB.Frame FrameRemunera 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Enabled         =   0   'False
      Height          =   4935
      Left            =   105
      TabIndex        =   7
      Top             =   1560
      Width           =   4095
      Begin VB.ListBox LstRemunera 
         Appearance      =   0  'Flat
         Height          =   1590
         Left            =   435
         TabIndex        =   15
         Top             =   600
         Visible         =   0   'False
         Width           =   3375
      End
      Begin MSDataGridLib.DataGrid DgrdRemunera 
         Height          =   4815
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   8493
         _Version        =   393216
         AllowUpdate     =   -1  'True
         Appearance      =   0
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         AllowAddNew     =   -1  'True
         AllowDelete     =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "descripcion"
            Caption         =   "Remuneraciones"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "codigo"
            Caption         =   "codigo"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               Button          =   -1  'True
               ColumnWidth     =   3390.236
            EndProperty
            BeginProperty Column01 
               Object.Visible         =   0   'False
               ColumnWidth     =   1170.142
            EndProperty
         EndProperty
      End
   End
   Begin VB.ComboBox CmbRemunera 
      Height          =   315
      Left            =   4200
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   720
      Width           =   2070
   End
   Begin VB.ComboBox CmbTipo 
      Height          =   315
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   720
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6495
      Begin VB.ComboBox Cmbcia 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   120
         Width           =   5190
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
         ForeColor       =   &H80000014&
         Height          =   210
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   825
      End
   End
   Begin VB.Label Lblcargo 
      AutoSize        =   -1  'True
      Caption         =   "Cargo"
      Height          =   195
      Left            =   600
      TabIndex        =   16
      Top             =   1200
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Concepto"
      Height          =   195
      Left            =   3360
      TabIndex        =   5
      Top             =   720
      Width           =   690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "T. Trabajador"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   960
   End
End
Attribute VB_Name = "Frmpromedio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim VTipo As String
Dim VCargo As String
Dim VRemunera As String
Dim rsremunera As New Recordset

Private Sub CmbCargo_Click()
VCargo = fc_CodigoComboBox(CmbCargo, 2)
Procesa
End Sub

Private Sub Cmbcia_Click()
wcia = Trim(Right("00" & Cmbcia.ItemData(Cmbcia.ListIndex), 2))
Call fc_Descrip_Maestros2("01055", "", Cmbtipo)
Sql$ = "Select rtrim(codinterno) as codinterno,rtrim(descripcion) as descripcion from placonstante where cia='" & wcia & "' and tipomovimiento='02' and status<>'*' order by codinterno"
Cmbremunera.Clear
LstRemunera.Clear
If (fAbrRst(rs, Sql$)) Then
   If (Not rs.EOF) Then
      Do Until rs.EOF
         If rs(0) = "16" Then
            Cmbremunera.AddItem rs(1)
            Cmbremunera.ItemData(Cmbremunera.NewIndex) = rs(0)
         End If
         LstRemunera.AddItem rs(1) & Space(60) & rs(0)
         rs.MoveNext
       Loop
    End If
    If rs.State = 1 Then rs.Close
End If
Cmbremunera.ListIndex = 0
Cmbremunera.Enabled = False
Procesa
End Sub
Private Sub Cmbremunera_Click()
VRemunera = fc_CodigoComboBox(Cmbremunera, 2)
If Cmbtipo.ListIndex > -1 And Cmbremunera.ListIndex > -1 Then
   FrameRemunera.Enabled = True
   DGrdRemunera.Col = 0
Else
   FrameRemunera.Enabled = False
End If
Procesa
End Sub
Private Sub CmbTipo_Click()
Dim wciamae As String
VTipo = fc_CodigoComboBox(Cmbtipo, 2)
If Cmbtipo.ListIndex > -1 And Cmbremunera.ListIndex > -1 Then
   FrameRemunera.Enabled = True
   DGrdRemunera.Col = 0
Else
   FrameRemunera.Enabled = False
End If
If VTipo = "05" Then
   VCargo = ""
   Lblcargo.Visible = True
   CmbCargo.Visible = True
   Sql$ = "select cod_maestro3,descrip from maestros_3 where ciamaestro='" & wcia & "055" & "' and status!='*'"
   Call rCarCbo(CmbCargo, Sql$, "C", "00")
Else
   VCargo = ""
   Lblcargo.Visible = False
   CmbCargo.Visible = False
End If

Procesa
End Sub

Private Sub DgrdRemunera_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
If DGrdRemunera.Col = 0 Then
        KeyAscii = 0
        Cancel = True
        DgrdRemunera_ButtonClick (ColIndex)
End If
End Sub

Private Sub DgrdRemunera_ButtonClick(ByVal ColIndex As Integer)
'Dim Y As Integer ', xtop As Integer, xleft As Integer
'If DgrdRemunera.Row < 0 Then rsremunera.AddNew
'Y = DgrdRemunera.Row
'xtop = DgrdRemunera.Top + DgrdRemunera.RowTop(Y) + DgrdRemunera.RowHeight
'Select Case ColIndex
'Case 0:
'       xleft = DgrdRemunera.Left + DgrdRemunera.Columns(1).Left + 120
'       With LstRemunera
'       If Y < 8 Then
'         .Top = xtop
'       Else
'         .Top = DgrdRemunera.Top + DgrdRemunera.RowTop(Y) - .Height
'       End If
'        .Left = xleft
'        .Visible = True
'        .SetFocus
'        .ZOrder 0
'       End With
'End Select
    If DGrdRemunera.Row < 0 Then rsremunera.AddNew
    If ColIndex = 0 Then
       LstRemunera.Top = DGrdRemunera.Top + DGrdRemunera.RowTop(DGrdRemunera.Row) + DGrdRemunera.RowHeight
       LstRemunera.Left = DGrdRemunera.Left + DGrdRemunera.Columns(ColIndex).Left
       LstRemunera.Width = DGrdRemunera.Columns(ColIndex).Width
       LstRemunera.Height = 1440
       LstRemunera.Visible = True
       LstRemunera.SetFocus
       LstRemunera.ZOrder 0
    End If

End Sub

Private Sub DgrdRemunera_OnAddNew()
rsremunera.AddNew
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
Me.Height = 7065
Me.Width = 6570
Crea_Rs
Call rCarCbo(Cmbcia, Carga_Cia, "C", "00")
Call rUbiIndCmbBox(Cmbcia, wcia, "00")
End Sub
Private Sub LstRemunera_Click()
Dim m, P As Integer
If LstRemunera.ListIndex > -1 Then
   m = Len(LstRemunera.Text) - 2
   DGrdRemunera.Columns(0) = Trim(Left(LstRemunera.Text, m))
   DGrdRemunera.Columns(1) = Format(Right(LstRemunera.Text, 2), "00")
   DGrdRemunera.SetFocus
   LstRemunera.Visible = False
End If
End Sub

Private Sub LstRemunera_LostFocus()
LstRemunera.Visible = False
End Sub

Private Sub Txtperiodo_KeyPress(KeyAscii As Integer)
TxtPeriodo.Text = TxtPeriodo.Text + fc_ValNumeros(KeyAscii)
End Sub

Private Sub UpDown1_DownClick()
If TxtPeriodo.Text = "" Then TxtPeriodo.Text = "0"
If TxtPeriodo.Text > 0 Then TxtPeriodo = TxtPeriodo - 1
End Sub

Private Sub UpDown1_UpClick()
If TxtPeriodo.Text = "" Then TxtPeriodo.Text = "0"
TxtPeriodo = TxtPeriodo + 1
End Sub
Private Sub Crea_Rs()
    If rsremunera.State = 1 Then rsremunera.Close
    rsremunera.Fields.Append "descripcion", adChar, 35, adFldIsNullable
    rsremunera.Fields.Append "codigo", adChar, 2, adFldIsNullable
    rsremunera.Fields.Append "tipo", adChar, 2, adFldIsNullable
    rsremunera.Open
    Set DGrdRemunera.DataSource = rsremunera
End Sub
Public Function Graba_Promedio()
Dim Mstatus As String
Dim NroTrans As Integer
On Error GoTo ErrorTrans
NroTrans = 0

If TxtPeriodo = "" Or TxtPeriodo = "0" Then MsgBox "Debe Ingresar Periodo", vbCritical, TitMsg: TxtPeriodo.SetFocus: Exit Function
If OpcMes.Value = False And OpcSemana.Value = False Then MsgBox "Debe Indicar Tipo de Periodo", vbCritical, TitMsg: Exit Function
If OpcMes.Value = True Then Mstatus = "M"
If OpcSemana.Value = True Then Mstatus = "S"
Mgrab = MsgBox("Seguro de Grabar Promedios", vbYesNo + vbQuestion, TitMsg)
If Mgrab <> 6 Then Exit Function
Screen.MousePointer = vbArrowHourglass

cn.BeginTrans
NroTrans = 1

Dim Rq As ADODB.Recordset
Sql = "select cod_cia from cia where status<>'*' "
If fAbrRst(Rq, Sql) Then Rq.MoveFirst
Do While Not Rq.EOF
   Sql$ = "Update platasaanexo set status='*' where cia='" & Rq!cod_cia & "' and modulo='01' and tipomovimiento='" & VTipo & "' and basecalculo='" & VRemunera & "' and cargo='" & VCargo & "' and status<>'*'"
   cn.Execute Sql$

   If rsremunera.RecordCount > 0 Then rsremunera.MoveFirst
   Do While Not rsremunera.EOF
      If rsremunera!codigo <> "" Then
         Sql$ = "INSERT INTO platasaanexo values('" & Rq!cod_cia & "','" & VTipo & "','01','" & rsremunera!codigo & "','" & VRemunera & "'," & CCur(TxtPeriodo.Text) & ",'" & Mstatus & "'," & FechaSys & ",'" & VTipo & "','" & VCargo & "','" & Trim(rsremunera!tipo) & "'," & CCur(TxtFactor.Text) & ")"
         cn.Execute Sql$
      End If
      rsremunera.MoveNext
   Loop
   Rq.MoveNext
Loop
Rq.Close: Set Rq = Nothing

cn.CommitTrans
Screen.MousePointer = vbDefault
MsgBox "Se guardarón los datos correctamente", vbInformation, Me.Caption

Exit Function

ErrorTrans:
If NroTrans = 1 Then
    cn.RollbackTrans
End If
Screen.MousePointer = vbDefault
MsgBox ERR.Description, vbCritical, Me.Caption

End Function
Private Sub Procesa()
If rsremunera.RecordCount > 0 Then
   rsremunera.MoveFirst
   Do While Not rsremunera.EOF
      rsremunera.Delete
      rsremunera.MoveNext
   Loop
End If
TxtPeriodo.Text = "0"
OpcMes.Value = True
OpcSemana.Value = False

If VTipo = "05" Then
   If VTipo = "" Or VCargo = "" Then Exit Sub
Else
   If VTipo = "" Then Exit Sub
End If

If Cmbcia.ListIndex < 0 Then Exit Sub
If Cmbtipo.ListIndex < 0 Then Exit Sub
If Cmbremunera.ListIndex < 0 Then Exit Sub
Sql$ = "select a.*,b.descripcion from platasaanexo a,placonstante b where a.modulo='01' and a.status<>'*' and a.tipomovimiento='" & VTipo & "' " _
     & "and a.basecalculo='" & VRemunera & "' and a.tipotrab='" & Format(Cmbtipo.ItemData(Cmbtipo.ListIndex), "00") & "' and b.tipomovimiento='02' and a.cargo='" & VCargo & "' and b.status<>'*' and a.cia=b.cia and a.codinterno=b.codinterno AND b.cia='" & wcia & "'"

cn.CursorLocation = adUseClient
Set rs = New ADODB.Recordset
Set rs = cn.Execute(Sql$, 64)
If rs.RecordCount > 0 Then
   rs.MoveFirst
   If rs!status = "S" Then OpcSemana.Value = True Else OpcMes.Value = True
   TxtPeriodo.Text = rs!factor
   TxtFactor.Text = rs!factor_divisionario
End If
Do While Not rs.EOF
   rsremunera.AddNew
   rsremunera!Descripcion = rs!Descripcion
   rsremunera!codigo = rs!codinterno
   rsremunera!tipo = rs!tipo
   rs.MoveNext
Loop
If rs.State = 1 Then rs.Close
End Sub
