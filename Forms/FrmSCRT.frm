VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmSCTR 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SCTR - Seguro Complementario del Trabajador en Riesgo"
   ClientHeight    =   4740
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7125
   Icon            =   "FrmSCRT.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   7125
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Height          =   3735
      Left            =   90
      TabIndex        =   4
      Top             =   840
      Width           =   6930
      Begin VB.TextBox TxtTopeMax 
         Alignment       =   1  'Right Justify
         Height          =   305
         Left            =   1200
         TabIndex        =   10
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox Txtano 
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   5880
         TabIndex        =   8
         Top             =   360
         Width           =   735
      End
      Begin VB.ComboBox Cmbmes 
         Height          =   315
         ItemData        =   "FrmSCRT.frx":030A
         Left            =   2880
         List            =   "FrmSCRT.frx":0332
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   360
         Width           =   2175
      End
      Begin MSDataGridLib.DataGrid DgrdSctr 
         Height          =   2535
         Left            =   240
         TabIndex        =   9
         Top             =   960
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   4471
         _Version        =   393216
         AllowUpdate     =   -1  'True
         Appearance      =   0
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
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
         ColumnCount     =   4
         BeginProperty Column00 
            DataField       =   "descripcion"
            Caption         =   "Descripcion"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "porsalud"
            Caption         =   "Seg. Salud(%)"
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
         BeginProperty Column02 
            DataField       =   "porpension"
            Caption         =   "Seg. Pension(%)"
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
         BeginProperty Column03 
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
               ColumnWidth     =   3704.882
            EndProperty
            BeginProperty Column01 
               Alignment       =   1
               ColumnWidth     =   1200.189
            EndProperty
            BeginProperty Column02 
               Alignment       =   1
               ColumnWidth     =   1200.189
            EndProperty
            BeginProperty Column03 
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tope Maximo"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   345
         Width           =   960
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Año"
         Height          =   195
         Left            =   5280
         TabIndex        =   7
         Top             =   360
         Width           =   285
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mes"
         Height          =   195
         Left            =   2520
         TabIndex        =   5
         Top             =   345
         Width           =   300
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8895
      Begin VB.ComboBox Cmbcia 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1035
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   120
         Width           =   3975
      End
      Begin VB.Label Lblfecha 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5160
         TabIndex        =   3
         Top             =   120
         Width           =   1815
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
         ForeColor       =   &H8000000E&
         Height          =   210
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   825
      End
   End
End
Attribute VB_Name = "FrmSCTR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsSCTR As New Recordset
Dim Vmes As String
Dim VBcoAfp As String
Dim VAreaAfp As String

Private Sub Cmbcia_Click()
wcia = Trim(Right("00" & CmbCia.ItemData(CmbCia.ListIndex), 2))
Cmbmes.ListIndex = Val(Month(Date)) - 1
Txtano.Text = Right(Lblfecha.Caption, 4)
'Call fc_Descrip_Maestros2(wcia & "007", "", CmbBcoAfp)
'Call fc_Descrip_Maestros2("01044", "", CmbAreaAfp)
Procesa
End Sub

Private Sub CmbMes_Click()
Vmes = Format(Cmbmes.ListIndex + 1, "00")
Procesa

End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
Me.Height = 5115
Me.Width = 7215
Lblfecha.Caption = Format(Date, "dd/mm/yyyy")
Crea_Rs
Call rCarCbo(CmbCia, Carga_Cia, "C", "00")
Call rUbiIndCmbBox(CmbCia, wcia, "00")
'SSTab1.Tab = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
If rs.State = 1 Then rs.Close
Set DgrdSctr.DataSource = Nothing
If rsSCTR.State = 1 Then rsSCTR.Close
Set rsafp = Nothing
End Sub


Private Sub Txtano_Change()
Procesa
End Sub

Private Sub Txtano_KeyPress(KeyAscii As Integer)
Txtano.Text = Txtano.Text + fc_ValNumeros(KeyAscii)
End Sub
Private Sub Crea_Rs()
    If rsSCTR.State = 1 Then rsSCTR.Close
    rsSCTR.Fields.Append "codsctr", adChar, 2, adFldIsNullable
    rsSCTR.Fields.Append "descripcion", adChar, 90, adFldIsNullable
    rsSCTR.Fields.Append "porsalud", adCurrency, 18, adFldIsNullable
    rsSCTR.Fields.Append "porpension", adCurrency, 18, adFldIsNullable
    rsSCTR.Fields.Append "periodo", adChar, 4, adFldIsNullable
    rsSCTR.Open
    Set DgrdSctr.DataSource = rsSCTR
End Sub
Private Sub Procesa()
Dim wciamae As String
Dim mperiodo As String
Dim rs2 As ADODB.Recordset
Dim I As Integer

mperiodo = Txtano.Text & Vmes
'TxtTopeMax.Text = "0.00"
'se cambio el maestro 01161 por 01074
wciamae = Determina_Maestro("01074")
Sql$ = "Select * from maestros_2 where cod_maestro2<>'00' AND status<>'*'"
Sql$ = Sql$ & wciamae
cn.CursorLocation = adUseClient
Set rs = New ADODB.Recordset
Set rs = cn.Execute(Sql$, 64)
If rsSCTR.RecordCount > 0 Then
   rsSCTR.MoveFirst
   Do While Not rsSCTR.EOF
      rsSCTR.Delete
      rsSCTR.MoveNext
   Loop
End If
If Not rs.RecordCount > 0 Then MsgBox "No Existen AFP's Registradas", vbCritical, TitMsg: Exit Sub
rs.MoveFirst
Do While Not rs.EOF
   rsSCTR.AddNew
   rsSCTR!codsctr = Trim(rs!cod_maestro2)
   rsSCTR!Descripcion = Trim(rs!DESCRIP)
   
   Sql$ = "Select * from plasctr where cia='" & wcia & "' and codsctr='" & rs!cod_maestro2 & "' and status<>'*' and periodo='" & mperiodo & "'"
   cn.CursorLocation = adUseClient
   Set rs2 = New ADODB.Recordset
   Set rs2 = cn.Execute(Sql$, 64)
   If rs2.RecordCount <= 0 Then
      If rs2.State = 1 Then rs2.Close
      If Vmes = "01" Then mperiodo = Format(Val(Txtano.Text) - 1, "0000") & "12" Else: mperiodo = Txtano.Text & Format(Val(Vmes) - 1, "00")
      Sql$ = "Select * from plasctr where cia='" & wcia & "' and codsctr='" & Trim(rs!cod_maestro2) & "' and status<>'*' and periodo='" & mperiodo & "'"
      cn.CursorLocation = adUseClient
      Set rs2 = New ADODB.Recordset
      Set rs2 = cn.Execute(Sql$, 64)
   End If
   If rs2.RecordCount > 0 Then
      'TxtTopeMax.Text = Format(rs2!TOPE, "###,###.00")
      rsSCTR!porsalud = rs2!porsalud
      rsSCTR!porpension = rs2!porpension
      TxtTopeMax.Text = rs2!tope
   Else
      rsSCTR!porsalud = "0.00"
      rsSCTR!porpension = "0.00"
      TxtTopeMax.Text = "0.00"
   End If
   If rs2.State = 1 Then rs2.Close
   rs.MoveNext
   
Loop

If rs.State = 1 Then rs.Close

'Sql$ = "Select * from cia where cod_cia='" & wcia & "' and status<>'*'"
'cn.CursorLocation = adUseClient
'Set rs = New ADODB.Recordset
'Set rs = cn.Execute(Sql$, 64)
'If rs.RecordCount > 0 Then
'   TxtRespAfp.Text = rs!afpresponsable
'   TxtTlfAfp.Text = rs!afpresptlf
'   Call rUbiIndCmbBox(CmbBcoAfp, rs!afpbanco, "00")
'   Call rUbiIndCmbBox(CmbAreaAfp, rs!afparearesp, "00")
'   For I = 0 To Cmbcta.ListCount - 1
'       If Left(Cmbcta.List(I), 15) = Left(rs!afpnrocta, 15) Then Cmbcta.ListIndex = I: Exit For
'   Next
'End If

End Sub
Public Function GrabarSCTR()
Dim mperiodo As String
Dim tope As String
Dim NroTrans As Integer
On Error GoTo ErrorTrans
NroTrans = 0

tope = TxtTopeMax.Text
Mgrab = MsgBox("Seguro de Grabar Seteo de SCTR", vbYesNo + vbQuestion, TitMsg)
If Mgrab <> 6 Then Exit Function
mperiodo = Txtano.Text & Vmes
'If TxtTopeMax.Text = "" Then MsgBox "Debe Registrar Tope Maximo", vbCritical, TitMsg: Exit Function
'If CCur(TxtTopeMax.Text) <= 0 Then MsgBox "Debe Registrar Tope Maximo", vbCritical, TitMsg: Exit Function
If CCur(Txtano.Text) <= 0 Then MsgBox "Debe Registrar Año", vbCritical, TitMsg: Exit Function

Screen.MousePointer = vbArrowHourglass

cn.BeginTrans
NroTrans = 1

Dim Rq As ADODB.Recordset
Sql = "select cod_cia from cia where status<>'*' "
If fAbrRst(Rq, Sql) Then Rq.MoveFirst
Do While Not Rq.EOF
   Sql$ = "Update plaSCTR set status='*' where cia='" & Rq!cod_cia & "' and periodo='" & mperiodo & "' and status<>'*'"
   cn.Execute Sql$
   If rsSCTR.RecordCount > 0 Then rsSCTR.MoveFirst
   Do While Not rsSCTR.EOF
      Sql$ = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
      Sql$ = Sql$ & "INSERT INTO plaSCTR values('" & Rq!cod_cia & "','" & rsSCTR!codsctr & "','" & Trim(rsSCTR!Descripcion) & "'," & CCur(rsSCTR!porsalud) & "," & CCur(rsSCTR!porpension) & "," & CCur(tope) & "," _
           & "''," & FechaSys & ",'" & mperiodo & "')"
      cn.Execute Sql$
      rsSCTR.MoveNext
   Loop
   Rq.MoveNext
Loop
Rq.Close: Set Rq = Nothing
cn.CommitTrans
MsgBox "Se guardarón los datos correctamente", vbInformation, Me.Caption
Screen.MousePointer = vbDefault
Exit Function

ErrorTrans:
    
    If NroTrans = 1 Then
        cn.CommitTrans
    End If
    
    MsgBox ERR.Description, vbCritical, Me.Caption
    Screen.MousePointer = vbDefault
End Function

Private Sub TxtTopeMax_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Cmbmes.SetFocus
End Sub
Private Sub TxtTopeMax_LostFocus()
TxtTopeMax.Text = Format(TxtTopeMax.Text, "###,###.00")
End Sub
