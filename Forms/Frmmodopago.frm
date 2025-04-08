VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Frmmodopago 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "» Periodo de Pago «"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5520
   Icon            =   "Frmmodopago.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   2535
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   5565
      Begin VB.ListBox Lsttipo 
         Height          =   1425
         Left            =   2700
         TabIndex        =   5
         Top             =   600
         Visible         =   0   'False
         Width           =   1590
      End
      Begin MSDataGridLib.DataGrid Dgrdmodo 
         Height          =   2295
         Left            =   150
         TabIndex        =   2
         Top             =   120
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   4048
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
         ColumnCount     =   5
         BeginProperty Column00 
            DataField       =   "trabajador"
            Caption         =   "Tipo de Trabajador"
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
            DataField       =   "modo"
            Caption         =   "Modo de Pago"
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
            DataField       =   "codtipo"
            Caption         =   "codtipo"
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
         BeginProperty Column03 
            DataField       =   "codmodo"
            Caption         =   "codmodo"
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
         BeginProperty Column04 
            DataField       =   "horas"
            Caption         =   "Horas"
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
               Locked          =   -1  'True
               ColumnWidth     =   2234.835
            EndProperty
            BeginProperty Column01 
               Button          =   -1  'True
               ColumnWidth     =   1574.929
            EndProperty
            BeginProperty Column02 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column03 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   915.024
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
      Width           =   5535
      Begin VB.ComboBox Cmbcia 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1050
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   80
         Width           =   4380
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
Attribute VB_Name = "Frmmodopago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsmodo As New Recordset
Private Sub Cmbcia_Click()
Dim wciamae As String
wcia = Trim(Right("00" & Cmbcia.ItemData(Cmbcia.ListIndex), 2))
LstTipo.Clear
wciamae = Determina_Maestro("01076")
Sql$ = "Select cod_maestro2,descrip from maestros_2 where  status<>'*'"
Sql$ = Sql$ & wciamae
cn.CursorLocation = adUseClient
Set rs = New ADODB.Recordset
Set rs = cn.Execute(Sql$, 64)
If rs.RecordCount > 0 Then rs.MoveFirst
Do Until rs.EOF
   LstTipo.AddItem rs!descrip & Space(100) & Trim(rs!COD_MAESTRO2)
   rs.MoveNext
Loop
Procesa
End Sub

Private Sub Dgrdmodo_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
If Dgrdmodo.Col = 1 Then
        KeyAscii = 0
        Cancel = True
        Dgrdmodo_ButtonClick (ColIndex)
End If
End Sub
Private Sub Dgrdmodo_ButtonClick(ByVal ColIndex As Integer)
Dim Y As Integer, xtop As Integer, xleft As Integer
Y = Dgrdmodo.Row
xtop = Dgrdmodo.Top + Dgrdmodo.RowTop(Y) + Dgrdmodo.RowHeight
Select Case ColIndex
Case 1:
       xleft = Dgrdmodo.Left + Dgrdmodo.Columns(1).Left
       With LstTipo
       If Y < 8 Then
         .Top = xtop
       Else
         .Top = Dgrdmodo.Top + Dgrdmodo.RowTop(Y) - .Height
       End If
        .Left = xleft
        .Visible = True
        .SetFocus
        .ZOrder 0
       End With
End Select
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
Me.Height = 3420
Me.Width = 5640
Crea_Rs
Call rCarCbo(Cmbcia, Carga_Cia, "C", "00")
Call rUbiIndCmbBox(Cmbcia, wcia, "00")
End Sub
Private Sub Crea_Rs()
    If rsmodo.State = 1 Then rsmodo.Close
    rsmodo.Fields.Append "trabajador", adChar, 200, adFldIsNullable
    rsmodo.Fields.Append "modo", adChar, 45, adFldIsNullable
    rsmodo.Fields.Append "codtipo", adChar, 2, adFldIsNullable
    rsmodo.Fields.Append "codmodo", adChar, 2, adFldIsNullable
    rsmodo.Fields.Append "horas", adInteger, 3, adFldIsNullable
    rsmodo.Open
    Set Dgrdmodo.DataSource = rsmodo
End Sub
Private Sub Procesa()
Dim wciamae As String
Dim rs2 As ADODB.Recordset
wciamae = Determina_Maestro("01055")
Sql$ = "Select * from maestros_2 where status<>'*'"
Sql$ = Sql$ & wciamae
cn.CursorLocation = adUseClient
Set rs = New ADODB.Recordset
Set rs = cn.Execute(Sql$, 64)
If rsmodo.RecordCount > 0 Then
   rsmodo.MoveFirst
   Do While Not rsmodo.EOF
      rsmodo.Delete
      rsmodo.MoveNext
   Loop
End If
If Not rs.RecordCount > 0 Then Exit Sub
rs.MoveFirst
Do While Not rs.EOF
   rsmodo.AddNew
   rsmodo!codtipo = Trim(rs!COD_MAESTRO2)
   rsmodo!trabajador = Trim(rs!descrip)
   rsmodo!codmodo = Left(rs!flag1, 2)
   If rs!flag2 = "" Or IsNull(rs!flag2) Then rsmodo!horas = 0 Else rsmodo!horas = Val(rs!flag2)
   If Trim(Left(rs!flag1, 2)) <> "" Then
      wciamae = Determina_Maestro("01076")
      Sql$ = "Select * from maestros_2 where cod_maestro2='" & Left(rs!flag1, 2) & "' and status<>'*'"
      Sql$ = Sql$ & wciamae
      cn.CursorLocation = adUseClient
      Set rs2 = New ADODB.Recordset
      Set rs2 = cn.Execute(Sql$, 64)
      If rs.RecordCount > 0 Then rsmodo!modo = rs2!descrip
      If rs2.State = 1 Then rs2.Close
   End If
   rs.MoveNext
Loop
If rs.State = 1 Then rs.Close
End Sub

Private Sub Lsttipo_Click()
Dim wciamae As String
Dim m, P As Integer
If LstTipo.ListIndex > -1 Then
   m = Len(LstTipo.Text) - 2
   Dgrdmodo.Columns(1) = Trim(Left(LstTipo.Text, m))
   Dgrdmodo.Columns(3) = Format(Right(LstTipo.Text, 2), "00")
   
   wciamae = Determina_Maestro("01076")
   Sql$ = "Select flag2 from maestros_2 where  status<>'*' and cod_maestro2='" & Format(Right(LstTipo.Text, 2), "00") & "'"
   Sql$ = Sql$ & wciamae
   cn.CursorLocation = adUseClient
   Set rs = New ADODB.Recordset
   Set rs = cn.Execute(Sql$, 64)
   If rs.RecordCount > 0 Then Dgrdmodo.Columns(4) = Trim(rs!flag2)
   rs.Close
   
   Dgrdmodo.SetFocus
   LstTipo.Visible = False
End If
End Sub

Private Sub Lsttipo_LostFocus()
LstTipo.Visible = False
End Sub
Public Function GrabarPerPago()
Dim mmodo As String
Dim mhoras As String
Mgrab = MsgBox("Seguro de Grabar Seteo de Periodos de Pago", vbYesNo + vbQuestion, TitMsg)
If Mgrab <> 6 Then Exit Function
Screen.MousePointer = vbArrowHourglass
Sql$ = wInicioTrans
cn.Execute Sql$
If rsmodo.RecordCount > 0 Then rsmodo.MoveFirst

Dim xciamae As String
Dim cod As String
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

Do While Not rsmodo.EOF
   If IsNull(rsmodo!codmodo) Then mmodo = "" Else mmodo = rsmodo!codmodo
   If IsNull(rsmodo!horas) Then mhoras = "" Else mhoras = rsmodo!horas
   Sql$ = "update maestros_2 set flag1='" & mmodo & "',flag2='" & mhoras & "' where cod_maestro2='" & rsmodo!codtipo & "' and status<>'*'"
   Sql$ = Sql$ & xciamae
   cn.Execute Sql$
   rsmodo.MoveNext
Loop

Sql$ = wFinTrans
cn.Execute Sql$
Screen.MousePointer = vbDefault
End Function


