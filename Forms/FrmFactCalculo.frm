VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmFactCalculo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "» Factor de Calculo (platasaanexo) «"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6360
   Icon            =   "FrmFactCalculo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   495
      Left            =   0
      TabIndex        =   7
      Top             =   600
      Width           =   6375
      Begin VB.ComboBox Cmbcargo 
         Height          =   315
         Left            =   4080
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   120
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.ComboBox Cmbtipotrab 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   120
         Width           =   2055
      End
      Begin VB.Label lblcargo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cargo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   3480
         TabIndex        =   10
         Top             =   120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "T. Trabajador"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   1080
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   5775
      Left            =   0
      TabIndex        =   1
      Top             =   1080
      Width           =   6375
      Begin VB.ComboBox CmbTipo 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   100
         Width           =   2295
      End
      Begin MSDataGridLib.DataGrid DgrdFactor 
         Height          =   5175
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   9128
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
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "descripcion"
            Caption         =   "Descripcion"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "factor"
            Caption         =   "Factor"
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
               ColumnWidth     =   4694.74
            EndProperty
            BeginProperty Column01 
               Alignment       =   1
               ColumnWidth     =   854.929
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Tope"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   120
         TabIndex        =   6
         Top             =   90
         Width           =   1080
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
      Width           =   6375
      Begin VB.ComboBox Cmbcia 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   120
         Width           =   5175
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
End
Attribute VB_Name = "FrmFactCalculo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim VTipo As String
Dim rsfactor As New Recordset
Dim VTipotrab As String
Dim VCargo As String

Private Sub CmbCargo_Click()
VCargo = fc_CodigoComboBox(CmbCargo, 2)
Procesa
End Sub

Private Sub Cmbcia_Click()
wcia = Trim(Right("00" & Cmbcia.ItemData(Cmbcia.ListIndex), 2))
Call fc_Descrip_Maestros2("01073", "", Cmbtipo)
Call fc_Descrip_Maestros2("01055", "", CmbTipoTrab)
Cmbtipo.ListIndex = 0
Procesa
End Sub

Private Sub CmbTipo_Click()
VTipo = fc_CodigoComboBox(Cmbtipo, 2)
Procesa
End Sub

Private Sub CmbTipoTrab_Click()
VTipotrab = fc_CodigoComboBox(CmbTipoTrab, 2)
If VTipotrab = "05" Then
   VCargo = ""
   Lblcargo.Visible = True
   CmbCargo.Visible = True
   wciamae = Determina_Maestro_2("01055")
   Sql$ = "select cod_maestro3,descrip from maestros_3 where ciamaestro='" & wcia & "055" & "' and status!='*'"
   'SQL$ = SQL$ & wciamae
   Call rCarCbo(CmbCargo, Sql$, "C", "00")
Else
   VCargo = ""
   Lblcargo.Visible = False
   CmbCargo.Visible = False
End If
Procesa
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
Me.Height = 7230
Me.Width = 6480
Crea_Rs
Call rCarCbo(Cmbcia, Carga_Cia, "C", "00")
Call rUbiIndCmbBox(Cmbcia, wcia, "00")
End Sub

Private Sub Crea_Rs()
    If rsfactor.State = 1 Then rsfactor.Close
    rsfactor.Fields.Append "codigo", adChar, 2, adFldIsNullable
    rsfactor.Fields.Append "descripcion", adChar, 45, adFldIsNullable
    rsfactor.Fields.Append "factor", adCurrency, 18, adFldIsNullable
    rsfactor.Fields.Append "TIPO", adChar, 2, adFldIsNullable
    rsfactor.Fields.Append "factordivi", adCurrency, 2, adFldIsNullable
    rsfactor.Open
    Set DgrdFactor.DataSource = rsfactor
End Sub
Private Sub Procesa()
Dim rs2 As ADODB.Recordset
Dim I, pos As Integer
Sql$ = "select * from placonstante where cia='" & wcia & "' and status<>'*' and tipomovimiento='02' and calculo='S' and factor='S'"
cn.CursorLocation = adUseClient
Set rs = New ADODB.Recordset
Set rs = cn.Execute(Sql$, 64)
If rsfactor.RecordCount > 0 Then
   rsfactor.MoveFirst
   Do While Not rsfactor.EOF
      rsfactor.Delete
      rsfactor.MoveNext
   Loop
End If
If VTipotrab = "05" Then
   If VTipotrab = "" Or VCargo = "" Then Exit Sub
Else
   If VTipotrab = "" Then Exit Sub
End If
If Not rs.RecordCount > 0 Then Exit Sub
rs.MoveFirst
Do While Not rs.EOF
   rsfactor.AddNew
   rsfactor!codigo = rs!codinterno
   rsfactor!Descripcion = rs!Descripcion
   'rsfactor!tipo = rs!tipo
   If VTipotrab = "05" Then
      Sql$ = "select * from platasaanexo where cia='" & wcia & "' and modulo='02' and tipomovimiento='" & VTipo & "' and codinterno='" & rs!codinterno & "' and tipotrab='" & VTipotrab & "' and cargo='" & VCargo & "' and status<>'*'"
   Else
      Sql$ = "select * from platasaanexo where cia='" & wcia & "' and modulo='02' and tipomovimiento='" & VTipo & "' and codinterno='" & rs!codinterno & "' and tipotrab='" & VTipotrab & "' and status<>'*'"
   End If
   cn.CursorLocation = adUseClient
   Set rs2 = New ADODB.Recordset
   Set rs2 = cn.Execute(Sql$, 64)
   If rs2.RecordCount > 0 Then
      rsfactor!factor = rs2!factor
      rsfactor!tipo = rs2!tipo
      rsfactor!factordivi = rs2!factor_divisionario
   Else
      rsfactor!factor = "0.00"
   End If
   If rs2.State = 1 Then rs2.Close
   rs.MoveNext
Loop
If rs.State = 1 Then rs.Close

End Sub
Public Function Graba_FactorCalculo()
Dim NroTrans As Integer
On Error GoTo ErrorTrans
NroTrans = 0

Mgrab = MsgBox("Seguro de Grabar Factores de Calculos", vbYesNo + vbQuestion, TitMsg)
If Mgrab <> 6 Then Exit Function

Screen.MousePointer = vbArrowHourglass
cn.BeginTrans
nerotrans = 1

   Dim Rq As ADODB.Recordset
   Sql = "select cod_cia from cia where status<>'*' "
   If fAbrRst(Rq, Sql) Then Rq.MoveFirst
   Do While Not Rq.EOF
      Sql$ = "Update platasaanexo set status='*' where cia='" & Rq!cod_cia & "' and modulo='02' and tipomovimiento='" & VTipo & "' and tipotrab='" & VTipotrab & "' and cargo='" & VCargo & "' and status<>'*'"
      cn.Execute Sql$
      If rsfactor.RecordCount > 0 Then rsfactor.MoveFirst
      Do While Not rsfactor.EOF
         If CCur(rsfactor!factor) <> 0 Then
            Sql$ = "INSERT INTO platasaanexo values('" & Rq!cod_cia & "','" & VTipo & "','02','" & rsfactor!codigo & "',''," & CCur(rsfactor!factor) & ",''," & FechaSys & ",'" & VTipotrab & "','" & VCargo & "'," & IIf(IsNull(rsfactor!tipo), "NULL", "'" & rsfactor!tipo & "'") & "," & IIf(IsNull(rsfactor!factordivi), 0, rsfactor!factordivi) & ")"
            cn.Execute Sql$
         End If
         rsfactor.MoveNext
   Loop
   Rq.MoveNext
Loop
cn.CommitTrans
Screen.MousePointer = vbDefault
MsgBox "Se guardarón los datos correctamente", vbInformation, Me.Caption

Exit Function
ErrorTrans:
If NroTrans = 1 Then
    cn.RollbackTrans
End If
MsgBox ERR.Description, vbCritical, Me.Caption
End Function


