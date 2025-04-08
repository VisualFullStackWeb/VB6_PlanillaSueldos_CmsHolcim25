VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Frmhorasper 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "» Seteo de Horas por Periodo «"
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4755
   Icon            =   "Frmhorasper.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   4755
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame4 
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   1575
      Left            =   0
      TabIndex        =   6
      Top             =   3360
      Width           =   4755
      Begin MSDataGridLib.DataGrid dgturnos 
         Height          =   1335
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   4515
         _ExtentX        =   7964
         _ExtentY        =   2355
         _Version        =   393216
         Appearance      =   0
         HeadLines       =   1
         RowHeight       =   11
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
         ColumnCount     =   3
         BeginProperty Column00 
            DataField       =   "CODTURNO"
            Caption         =   "CODIGO"
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
            DataField       =   "descripcion"
            Caption         =   "Descripcion"
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
            DataField       =   "horasxdia"
            Caption         =   "Horas x Dia"
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
               ColumnWidth     =   0
            EndProperty
            BeginProperty Column01 
            EndProperty
            BeginProperty Column02 
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Frame3"
      Height          =   15
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   4575
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   2775
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   4755
      Begin MSDataGridLib.DataGrid Dgrdperiodo 
         Height          =   2535
         Left            =   120
         TabIndex        =   2
         Top             =   75
         Width           =   4515
         _ExtentX        =   7964
         _ExtentY        =   4471
         _Version        =   393216
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
         ColumnCount     =   3
         BeginProperty Column00 
            DataField       =   "periodo"
            Caption         =   "Periodo"
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
            DataField       =   "horas"
            Caption         =   "Horas"
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
         BeginProperty Column02 
            DataField       =   "codperiodo"
            Caption         =   "codperiodo"
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
               ColumnWidth     =   2789.858
            EndProperty
            BeginProperty Column01 
               Alignment       =   1
               ColumnWidth     =   1049.953
            EndProperty
            BeginProperty Column02 
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4755
      Begin VB.ComboBox Cmbcia 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   80
         Width           =   3375
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
         TabIndex        =   3
         Top             =   120
         Width           =   825
      End
   End
End
Attribute VB_Name = "Frmhorasper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsperiodo As New Recordset
Dim RX As New ADODB.Recordset
Dim rtempo As New ADODB.Recordset

Private Sub Cmbcia_Click()
wcia = Trim(Right("00" & Cmbcia.ItemData(Cmbcia.ListIndex), 2))
Procesa
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
Me.Width = 4800
Me.Height = 3600
Crea_Rs
Call rCarCbo(Cmbcia, Carga_Cia, "C", "00")
Call rUbiIndCmbBox(Cmbcia, wcia, "00")
Call Me.Llena_data
End Sub
Private Sub Crea_Rs()
    If rsperiodo.State = 1 Then rsperiodo.Close
    rsperiodo.Fields.Append "periodo", adChar, 45, adFldIsNullable
    rsperiodo.Fields.Append "horas", adInteger, 3, adFldIsNullable
    rsperiodo.Fields.Append "codperiodo", adChar, 2, adFldIsNullable
    rsperiodo.Open
    Set Dgrdperiodo.DataSource = rsperiodo
End Sub
Private Sub Procesa()
Dim wciamae As String
wciamae = Determina_Maestro("01076")
Sql$ = "Select * from maestros_2 where status<>'*'"
Sql$ = Sql$ & wciamae
cn.CursorLocation = adUseClient
Set rs = New ADODB.Recordset
Set rs = cn.Execute(Sql$, 64)
If rsperiodo.RecordCount > 0 Then
   rsperiodo.MoveFirst
   Do While Not rsperiodo.EOF
      rsperiodo.Delete
      rsperiodo.MoveNext
   Loop
End If
If Not rs.RecordCount > 0 Then Exit Sub
rs.MoveFirst
Do While Not rs.EOF
   rsperiodo.AddNew
   rsperiodo!codperiodo = Trim(rs!COD_MAESTRO2)
   rsperiodo!Periodo = rs!DESCRIP
   rsperiodo!horas = Val(rs!flag2)
   
   rs.MoveNext
Loop
If rs.State = 1 Then rs.Close
End Sub
Public Function Grabar_PerHoras()
Dim mhoras As String
Dim NroTrans As Integer
On Error GoTo ErrorTrans
NroTrans = 0

Mgrab = MsgBox("Seguro de Grabar Seteo de Horas por Periodo", vbYesNo + vbQuestion, TitMsg)
If Mgrab <> 6 Then Exit Function
Screen.MousePointer = vbArrowHourglass
cn.BeginTrans
NroTrans = 1
If rsperiodo.RecordCount > 0 Then rsperiodo.MoveFirst
Dim xciamae As String
Dim cod As String
cod = "01076"
cn.CursorLocation = adUseClient
Set rs = New ADODB.Recordset
Sql$ = "SELECT GENERAL FROM maestros where right(ciamaestro,3)='" & Right(cod, 3) & "' and status<>'*' "
If (fAbrRst(rs, Sql$)) Then
   If rs!General = "S" Then
      xciamae = " and right(ciamaestro,3)= '" & Right(cod, 3) & "'"
   Else
      xciamae = " and ciamaestro= '" & wcia + Right(cod, 3) & "'"
   End If
End If
If rs.State = 1 Then rs.Close
Do While Not rsperiodo.EOF
   If IsNull(rsperiodo!horas) Then mhoras = "" Else mhoras = rsperiodo!horas
   Sql$ = "update maestros_2 set flag2='" & mhoras & "' where cod_maestro2='" & rsperiodo!codperiodo & "' and status<>'*'"
   Sql$ = Sql$ & xciamae
   cn.Execute Sql$
   rsperiodo.MoveNext
Loop

If rsperiodo.RecordCount > 0 Then rsperiodo.MoveFirst
cod = "01055"
cn.CursorLocation = adUseClient
Set rs = New ADODB.Recordset
Sql$ = "SELECT GENERAL FROM maestros where right(ciamaestro,3)='" & Right(cod, 3) & "' and status<>'*' "
If (fAbrRst(rs, Sql$)) Then
   If rs!General = "S" Then
      xciamae = " and right(ciamaestro,3)= '" & Right(cod, 3) & "'"
   Else
      xciamae = " and ciamaestro= '" & wcia + Right(cod, 3) & "'"
   End If
End If
If rs.State = 1 Then rs.Close
Do While Not rsperiodo.EOF
   If IsNull(rsperiodo!horas) Then mhoras = "" Else mhoras = rsperiodo!horas
   Sql$ = "update maestros_2 set flag2='" & mhoras & "' where flag1='" & rsperiodo!codperiodo & "' and status<>'*'"
   Sql$ = Sql$ & xciamae
   cn.Execute Sql$
   rsperiodo.MoveNext
Loop

cn.CommitTrans
MsgBox "Se guardaron los datos correctamente", vbInformation, Me.Caption
Screen.MousePointer = vbDefault

Exit Function

ErrorTrans:
If NroTrans = 1 Then
    cn.RollbackTrans
End If
MsgBox ERR.Description, vbCritical, Me.Caption
Screen.MousePointer = vbDefault


End Function
Sub Llena_data()
    Dim con As String
    Dim cia As String
    
    Exit Sub
    On Error GoTo CORRIGE
    With rtempo.Fields
        .Append "CODTURNO", adChar, 3, adFldIsNullable
        .Append "DESCRIPCION", adChar, 30, adFldIsNullable
        .Append "HORASXDIA", adDouble, 2
    End With
    rtempo.Open
    
    cia = "0" & Cmbcia.ItemData(Cmbcia.ListIndex)
    con = "select codturno,descripcion,horasxdia from platurno where cia='" & cia & "'"

    RX.Open con, cn, adOpenStatic, adLockBatchOptimistic
    Set rtempo = RX
'    rx.MoveFirst
'    While Not rx.EOF
'          RTEMPO.AddNew
'            RTEMPO.Fields(0) = rx("cod_turno")
'            RTEMPO.Fields(1) = rx("descripcion")
'            RTEMPO.Fields(2) = rx("horasxdia")
'          rx.MoveNext
'    Loop

    Set dgturnos.DataSource = rtempo
    Exit Sub
CORRIGE:
    MsgBox "Error : " & ERR.Description, vbCritical, "Sistema De Planillas"
End Sub

Private Sub Form_Unload(Cancel As Integer)
If rtempo.State = adStateOpen Then
   rtempo.Close
   Set rtempo = Nothing
End If

If RX.State = adStateOpen Then
   RX.Close
   'RTEMPO.Close
   Set RX = Nothing
End If
End Sub

Sub Grabar()
    On Error GoTo CORRIGE
    Dim cia As String
    Exit Sub
    cia = Cmbcia.ItemData(Cmbcia.ListIndex)

    con = "DELETE FROM PLATURNO WHERE cia='" & cia & "'"
    cn.Execute con
    
    Do While Not rtempo.EOF

    Loop
CORRIGE:
   MsgBox "Error :" & ERR.Description, vbCritical, "Sistema de Planillas"
End Sub










