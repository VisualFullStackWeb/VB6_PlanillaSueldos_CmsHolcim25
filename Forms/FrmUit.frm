VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmUit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "» Ingreso UIT / Sueldo Minimo «"
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5745
   Icon            =   "FrmUit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   5745
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtAsigFamiliar 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4080
      TabIndex        =   9
      Top             =   1560
      Width           =   1470
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Actualizar Tablas"
      Height          =   465
      Left            =   4080
      TabIndex        =   8
      Top             =   2040
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.TextBox TxtMinimo 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4080
      TabIndex        =   6
      Top             =   960
      Width           =   1470
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   1815
      Left            =   30
      TabIndex        =   1
      Top             =   660
      Width           =   3975
      Begin VB.ListBox LstMoneda 
         Height          =   645
         Left            =   1200
         TabIndex        =   7
         Top             =   600
         Visible         =   0   'False
         Width           =   1995
      End
      Begin MSDataGridLib.DataGrid DgrdUit 
         Height          =   1695
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   2990
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
         Caption         =   "UIT"
         ColumnCount     =   3
         BeginProperty Column00 
            DataField       =   "ano"
            Caption         =   "Año"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "moneda"
            Caption         =   "moneda"
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
         BeginProperty Column02 
            DataField       =   "importe"
            Caption         =   "importe"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   764.787
            EndProperty
            BeginProperty Column01 
               Button          =   -1  'True
               Locked          =   -1  'True
               ColumnWidth     =   884.976
            EndProperty
            BeginProperty Column02 
               Alignment       =   1
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6000
      Begin VB.ComboBox Cmbcia 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   120
         Width           =   4560
      End
      Begin VB.Label Label1 
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
         TabIndex        =   2
         Top             =   120
         Width           =   825
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "% Asig. Familiar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   4080
      TabIndex        =   10
      Top             =   1320
      Width           =   1290
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sueldo Minimo S/."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   4080
      TabIndex        =   5
      Top             =   720
      Width           =   1470
   End
End
Attribute VB_Name = "FrmUit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Vcode As Integer
Dim rsuit As New Recordset

Private Sub Cmbcia_Click()
Dim wciamae As String
wcia = Trim(Right("00" & Cmbcia.ItemData(Cmbcia.ListIndex), 2))
wciamae = Determina_Maestro("01006")
Sql$ = "SELECT cod_maestro2,descrip,flag1 FROM maestros_2 where status=''"
Sql$ = Sql$ & wciamae
Set rs = cn.Execute(Sql$)
If rs.RecordCount = 0 Then Exit Sub
rs.MoveFirst
LstMoneda.Clear
Do Until rs.EOF
   LstMoneda.AddItem rs!flag1 & Space(1) & rs!DESCRIP
   rs.MoveNext
Loop
If rs.State = 1 Then rs.Close
Procesa
End Sub

Private Sub Command1_Click()
If Not IsNumeric(TxtMinimo.Text) Then Exit Sub
Sql = "select minimo,FactAsigFamiliar from cia where cod_cia='" & wcia & "' and status<>'*'"
If (fAbrRst(rs, Sql)) Then
   If rs!Minimo <> CCur(TxtMinimo.Text) Or rs!FactAsigFamiliar <> CCur(txtAsigFamiliar.Text) Then
      MsgBox "Debe Grabar Primero", vbInformation, "Sueldo Minimo"
      Exit Sub
   End If
End If
'If MsgBox("Se actualizaran las tablas de remuneraciones basicas y asignacion familiar" & Chr(13) & "de acuero al suledo Minimo", vbExclamation & vbOKCancel) = vbOK Then
If MsgBox("Se actualizaran las tablas de asignacion familiar" & Chr(13) & "de acuero al suledo Minimo", vbExclamation & vbOKCancel) = vbOK Then
   Mgrab = MsgBox("Seguro de actualizar ?", vbYesNo + vbQuestion, "Sueldo Minimo")
   If Mgrab <> 6 Then Exit Sub
End If
Actualiza_Minimo
End Sub

Private Sub DgrdUit_AfterColEdit(ByVal ColIndex As Integer)
Select Case ColIndex
       Case Is = 0
            If DgrdUit.Columns(0) <> "" Then
            If CCur(DgrdUit.Columns(0)) < 1990 Or CCur(DgrdUit.Columns(0)) > 2050 Then
               MsgBox "Ingrese Correctramente el Año", vbCritical, "UIT"
               DgrdUit.Columns(0) = ""
            End If
            End If
End Select
End Sub

Private Sub DgrdUit_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
If DgrdUit.Col = 1 Then
        KeyAscii = 0
        Cancel = True
        DgrdUit_ButtonClick (ColIndex)
End If
End Sub

Private Sub DgrdUit_ButtonClick(ByVal ColIndex As Integer)
Dim Y As Integer, xtop As Integer, xleft As Integer
Y = DgrdUit.Row
xtop = DgrdUit.Top + DgrdUit.RowTop(Y) + DgrdUit.RowHeight
Select Case ColIndex
Case 1:
       xleft = DgrdUit.Left + DgrdUit.Columns(1).Left
       With LstMoneda
       If Y < 8 Then
         .Top = xtop
       Else
         .Top = DgrdUit.Top + DgrdUit.RowTop(Y) - .Height
       End If
        .Left = xleft
        .Visible = True
        .SetFocus
        .ZOrder 0
       End With
End Select
End Sub

Private Sub DgrdUit_OnAddNew()
rsuit.AddNew
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
Me.Height = 3045
Me.Width = 5835
Crea_Rs
Call rCarCbo(Cmbcia, Carga_Cia, "C", "00")
Call rUbiIndCmbBox(Cmbcia, wcia, "00")
End Sub
Private Sub Crea_Rs()
    If rsuit.State = 1 Then rsuit.Close
    rsuit.Fields.Append "ano", adInteger, 4, adFldIsNullable
    rsuit.Fields.Append "moneda", adChar, 3, adFldIsNullable
    rsuit.Fields.Append "importe", adCurrency, 18, adFldIsNullable
    rsuit.Open
    Set DgrdUit.DataSource = rsuit
End Sub
Private Sub Procesa()
Sql$ = "Select * from plauit where cia='" & wcia & "' and status<>'*' order by ano,moneda"
cn.CursorLocation = adUseClient
Set rs = New ADODB.Recordset
Set rs = cn.Execute(Sql$, 64)
If rsuit.RecordCount > 0 Then
   rsuit.MoveFirst
   Do While Not rsuit.EOF
      rsuit.Delete
      rsuit.MoveNext
   Loop
End If
If Not rs.RecordCount > 0 Then Exit Sub
rs.MoveFirst
Do While Not rs.EOF
   rsuit.AddNew
   rsuit!ano = rs!ano
   rsuit!moneda = rs!moneda
   rsuit!importe = rs!uit
   rs.MoveNext
Loop
If rs.State = 1 Then rs.Close

Sql$ = "Select minimo,FactAsigFamiliar from cia where cod_cia='" & wcia & "' and status<>'*'"
cn.CursorLocation = adUseClient
Set rs = New ADODB.Recordset
Set rs = cn.Execute(Sql$, 64)
If rs.RecordCount > 0 Then
   TxtMinimo.Text = Format(rs!Minimo, "###,###,###.00")
   txtAsigFamiliar.Text = Format(rs!FactAsigFamiliar, "###,###,###.00")
End If
If rs.State = 1 Then rs.Close
End Sub

Public Sub Graba_Uit()
Dim NroTrans As Integer
On Error GoTo ErrorTrans
NroTrans = 0
If Not IsNumeric(TxtMinimo.Text) Then
    MsgBox "Ingrese el sueldo mínimo", vbCritical, Me.Caption
    Exit Sub
End If

If Not IsNumeric(txtAsigFamiliar.Text) Then
    MsgBox "Ingrese el Porcentaje de Asignación Familiar", vbCritical, Me.Caption
    Exit Sub
End If

If Val(txtAsigFamiliar.Text) < 0 Or Val(txtAsigFamiliar.Text) > 100 Then
    MsgBox "El Porcentaje de Asignación Familiar debe estar entre 0 a 100%", vbCritical, Me.Caption
    Exit Sub
End If

If MsgBox("Desea Grabar UIT", vbYesNo + vbQuestion) = vbNo Then Screen.MousePointer = vbDefault: Exit Sub
cn.BeginTrans
NroTrans = 1
   Dim Rq As New ADODB.Recordset
   Sql = "select cod_cia from cia where status<>'*' "
   If fAbrRst(Rq, Sql) Then Rq.MoveFirst
   Do While Not Rq.EOF
      Sql$ = "update plauit set status='*' where cia='" & Rq!cod_cia & "' and status<>'*'"
      cn.Execute Sql$
      If rsuit.RecordCount > 0 Then
         rsuit.MoveFirst
         Do While Not rsuit.EOF
            If Not IsNull(rsuit!ano) Then
               If IsNull(rsuit!moneda) = True Then
                    NroTrans = 2
                    GoTo ErrorTrans
               ElseIf Trim(rsuit!moneda) = "" Then
                    NroTrans = 2
                    GoTo ErrorTrans
               End If
               If IsNull(rsuit!importe) = True Then
                    NroTrans = 3
                    GoTo ErrorTrans
               ElseIf Trim(rsuit!importe) <= 0 Then
                    NroTrans = 3
                    GoTo ErrorTrans
               End If
               
               Sql$ = "INSERT INTO plauit values('" & Rq!cod_cia & "','" & rsuit!ano & "','" & rsuit!moneda & "'," & CCur(rsuit!importe) & ",''," & FechaSys & ",'" & wuser & "')"
               cn.Execute Sql$
            End If
            rsuit.MoveNext
         Loop
      End If
      
      porcasigfamiliar = Round((CCur(txtAsigFamiliar.Text) / 100), 2)
      sueldominimo = Round(CCur(TxtMinimo.Text), 2)
      Sql$ = "update cia set minimo=" & CCur(TxtMinimo.Text) & ",FactAsigFamiliar=" & CCur(txtAsigFamiliar.Text) & " where cod_cia='" & Rq!cod_cia & "' and status<>'*'"
      cn.Execute Sql$
      Rq.MoveNext
   Loop
   Rq.Close: Set Rq = Nothing



cn.CommitTrans
MsgBox "Se guardarón los datos correctamente", vbInformation, Me.Caption
Screen.MousePointer = vbDefault
Command1.Visible = True
Procesa
Exit Sub
ErrorTrans:
If NroTrans >= 1 Then
    cn.RollbackTrans
End If

If ERR.Number <> 0 Then
 MsgBox ERR.Description, vbCritical, Me.Caption
ElseIf NroTrans = 2 Then
    MsgBox "Ingrese la Moneda del Periodo", vbCritical, Me.Caption
ElseIf NroTrans = 3 Then
    MsgBox "Ingrese la Importe del Periodo", vbCritical, Me.Caption
End If

Screen.MousePointer = vbDefault


End Sub

Private Sub LstMoneda_Click()
If Vcode = 0 Then Vcode = 13
Call LstMoneda_KeyDown(Vcode, 0)
Vcode = 0
End Sub

Private Sub LstMoneda_KeyDown(KeyCode As Integer, Shift As Integer)
Vcode = KeyCode
If KeyCode <> 13 Then Exit Sub
If Trim(LstMoneda) <> "" Then
   DgrdUit.Columns(1) = Mid$(Trim(LstMoneda), 2, 3)
   DgrdUit.Col = DgrdUit.Col + 1
   LstMoneda.Visible = False
End If
Vcode = 0
End Sub

Private Sub LstMoneda_LostFocus()
LstMoneda.Visible = False
Vcode = 0
End Sub

Private Sub txtAsigFamiliar_LostFocus()
txtAsigFamiliar.Text = Format(txtAsigFamiliar.Text, "###,###.00")
End Sub

Private Sub TxtMinimo_LostFocus()
TxtMinimo.Text = Format(TxtMinimo.Text, "###,###.00")
End Sub
Private Sub Actualiza_Minimo()
Dim NroTrans As Integer
On Error GoTo ErrorTrans
Dim Minimo As Currency
Dim mHourMonth As Currency
Screen.MousePointer = vbArrowHourglass
NroTrans = 0

'Determinar Horas Mensuales
wciamae = Determina_Maestro("01076")
Sql$ = "Select flag2 from maestros_2 where cod_maestro2='04' and status<>'*'"
Sql$ = Sql$ & wciamae
mHourMonth = 0
If (fAbrRst(rs, Sql$)) Then mHourMonth = Val(rs!flag2)

cn.BeginTrans
NroTrans = 1

'Asignacion Familiar Mensual en Soles
Minimo = (CCur(TxtMinimo.Text) * CCur(txtAsigFamiliar.Text)) / 100
Sql = "select m.cia,m.placod,b.* from plaremunbase b,planillas m where b.cia='" & wcia & "' and concepto='02' and b.status<>'*' " _
    & "and b.moneda='" & wmoncont & "'  and b.tipo='04' and b.importe<" & Minimo & " and m.cia=b.cia and m.placod=b.placod and m.fcese is null  and m.status<>'*'"

If (fAbrRst(rs, Sql$)) Then rs.MoveFirst
Do While Not rs.EOF
   Sql = "update plaremunbase set importe=" & Minimo & " where cia='" & rs!cia & "' and placod='" & rs!PlaCod & "' and concepto='" & rs!concepto & "' and moneda='" & rs!moneda & "' and tipo='" & rs!Tipo & "' and status<>'*'"
   cn.Execute Sql$
   rs.MoveNext
Loop
rs.Close

'Asignacion Quncenal en Soles
Minimo = ((CCur(TxtMinimo.Text) * CCur(txtAsigFamiliar.Text)) / 100) / mHourMonth
Sql = "select m.cia,m.placod,b.* from plaremunbase b,planillas m where b.cia='" & wcia & "' and concepto='02' " _
    & "and b.status<>'*' and b.moneda='" & wmoncont & "'  and b.tipo='03' and b.importe< " & Minimo & " * b.factor_horas and m.cia=b.cia and m.placod=b.placod " _
    & "and m.fcese is null  and m.status<>'*'"

If (fAbrRst(rs, Sql$)) Then rs.MoveFirst
Do While Not rs.EOF
   Sql = "update plaremunbase set importe=" & Minimo & " * factor_horas where cia='" & rs!cia & "' and placod='" & rs!PlaCod & "' and concepto='" & rs!concepto & "' and moneda='" & rs!moneda & "' and tipo='" & rs!Tipo & "' and status<>'*'"
   cn.Execute Sql$
   rs.MoveNext
Loop
rs.Close

'Asignacion Familiar Semanal en Soles
Minimo = ((CCur(TxtMinimo.Text) * CCur(txtAsigFamiliar.Text)) / 100) / mHourMonth
Sql = "select m.cia,m.placod,b.* from plaremunbase b,planillas m where b.cia='" & wcia & "' and concepto='02' " _
    & "and b.status<>'*' and b.moneda='" & wmoncont & "'  and b.tipo='02' and b.importe< " & Minimo & " * b.factor_horas and m.cia=b.cia and m.placod=b.placod " _
    & "and m.fcese is null  and m.status<>'*'"

If (fAbrRst(rs, Sql$)) Then rs.MoveFirst
Do While Not rs.EOF
   Sql = "update plaremunbase set importe=" & Minimo & " * factor_horas where cia='" & rs!cia & "' and placod='" & rs!PlaCod & "' and concepto='" & rs!concepto & "' and moneda='" & rs!moneda & "' and tipo='" & rs!Tipo & "' and status<>'*'"
   cn.Execute Sql$
   rs.MoveNext
Loop
rs.Close

'Asignacion Diario en Soles
Minimo = ((CCur(TxtMinimo.Text) * CCur(txtAsigFamiliar.Text)) / 100) / mHourMonth
Sql = "select m.cia,m.placod,b.* from plaremunbase b,planillas m where b.cia='" & wcia & "' and concepto='02' " _
    & "and b.status<>'*' and b.moneda='" & wmoncont & "'  and b.tipo='01' and b.importe< " & Minimo & " * b.factor_horas and m.cia=b.cia and m.placod=b.placod " _
    & "and m.fcese is null  and m.status<>'*'"

If (fAbrRst(rs, Sql$)) Then rs.MoveFirst
Do While Not rs.EOF
   Sql = "update plaremunbase set importe=" & Minimo & " * factor_horas where cia='" & rs!cia & "' and placod='" & rs!PlaCod & "' and concepto='" & rs!concepto & "' and moneda='" & rs!moneda & "' and tipo='" & rs!Tipo & "' and status<>'*'"
   cn.Execute Sql$
   rs.MoveNext
Loop
rs.Close

'Asignacion por HORA en Soles
Minimo = ((CCur(TxtMinimo.Text) * CCur(txtAsigFamiliar.Text)) / 100) / mHourMonth
Sql = "select m.cia,m.placod,b.* from plaremunbase b,planillas m where b.cia='" & wcia & "' and concepto='02' " _
    & "and b.status<>'*' and b.moneda='" & wmoncont & "'  and b.tipo='06' and b.importe< " & Minimo & " * b.factor_horas and m.cia=b.cia and m.placod=b.placod " _
    & "and m.fcese is null  and m.status<>'*'"

If (fAbrRst(rs, Sql$)) Then rs.MoveFirst
Do While Not rs.EOF
   Sql = "update plaremunbase set importe=" & Minimo & " * factor_horas where cia='" & rs!cia & "' and placod='" & rs!PlaCod & "' and concepto='" & rs!concepto & "' and moneda='" & rs!moneda & "' and tipo='" & rs!Tipo & "' and status<>'*'"
   cn.Execute Sql$
   rs.MoveNext
Loop
rs.Close


cn.CommitTrans
MsgBox "Se guardarón los datos correctamente", vbInformation, Me.Caption
Command1.Visible = False
Screen.MousePointer = vbDefault
Exit Sub
ErrorTrans:
If NroTrans = 1 Then
    cn.RollbackTrans
End If
MsgBox ERR.Description, vbCritical, Me.Caption
Screen.MousePointer = vbDefault

End Sub
