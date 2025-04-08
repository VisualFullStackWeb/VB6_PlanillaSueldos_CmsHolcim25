VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Frmfactsenati 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Factores de Calculo para Senati"
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6390
   Icon            =   "Frmfactsenati.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   6390
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Height          =   6495
      Left            =   0
      TabIndex        =   3
      Top             =   480
      Width           =   6375
      Begin MSDataGridLib.DataGrid Dgrdsenati 
         Height          =   6255
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   11033
         _Version        =   393216
         AllowUpdate     =   -1  'True
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
            DataField       =   "area"
            Caption         =   "Area"
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
         BeginProperty Column02 
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
               Locked          =   -1  'True
               ColumnWidth     =   4680
            EndProperty
            BeginProperty Column01 
               Alignment       =   1
               ColumnWidth     =   705.26
            EndProperty
            BeginProperty Column02 
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6375
      Begin VB.ComboBox Cmbcia 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   80
         Width           =   5175
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   840
      End
   End
End
Attribute VB_Name = "Frmfactsenati"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rssenati As New Recordset
Private Sub Cmbcia_Click()
wcia = Trim(Right("00" & Cmbcia.ItemData(Cmbcia.ListIndex), 2))
Crea_Rs
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
Me.Width = 6510
Me.Height = 7320
Call rCarCbo(Cmbcia, Carga_Cia, "C", "00")
Call rUbiIndCmbBox(Cmbcia, wcia, "00")
Procesa_Senati
End Sub
Private Sub Crea_Rs()
    If rssenati.State = 1 Then rssenati.Close
    rssenati.Fields.Append "area", adChar, 60, adFldIsNullable
    rssenati.Fields.Append "factor", adCurrency, 18, adFldIsNullable
    rssenati.Fields.Append "codigo", adChar, 2, adFldIsNullable
    rssenati.Open
    Set Dgrdsenati.DataSource = rssenati
End Sub
Private Sub Procesa_Senati()
Dim wciamae As String

wciamae = Determina_Maestro("01044")
Sql$ = "Select * from maestros_2 where status<>'*'"
Sql$ = Sql$ & wciamae
cn.CursorLocation = adUseClient
Set rs = New ADODB.Recordset
Set rs = cn.Execute(Sql$, 64)
If rssenati.RecordCount > 0 Then
   rssenati.MoveFirst
   Do While Not rssenati.EOF
      rssenati.Delete
      rssenati.MoveNext
   Loop
End If
If Not rs.RecordCount > 0 Then Exit Sub
rs.MoveFirst
Do While Not rs.EOF
   rssenati.AddNew
   rssenati!codigo = Trim(rs!COD_MAESTRO2)
   rssenati!Area = rs!DESCRIP
   If IsNull(rs!flag7) Then rssenati!factor = 0
   'If RS!flag7 = "" Then rssenati!factor = 0 Else rssenati!factor = CCur(RS!flag7)
   rs.MoveNext
Loop
rssenati.MoveFirst
If rs.State = 1 Then rs.Close
End Sub
Public Function Grabar_Senati()
Dim mm As String
Dim NroTrans As Integer
NroTrans = 0
On Error GoTo ErrorTrans

Mgrab = MsgBox("Seguro de Grabar Seteo de Senati", vbYesNo + vbQuestion, "Factores Para Senati")
If Mgrab <> 6 Then Exit Function
Screen.MousePointer = vbArrowHourglass
cn.BeginTrans
NroTrans = 1
If rssenati.RecordCount > 0 Then rssenati.MoveFirst

Dim xciamae As String
Dim cod As String

cod = "01044"
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

Do While Not rssenati.EOF
   If IsNull(rssenati!factor) Then
      mm = "0.00"
   Else
      mm = rssenati!factor
   End If
   
   Sql$ = "update maestros_2 set flag7='" & mm & "' where cod_maestro2='" & rssenati!codigo & "' and status<>'*'"
   Sql$ = Sql$ & xciamae
   cn.Execute Sql$
   rssenati.MoveNext
Loop

cn.CommitTrans
MsgBox "Se guardaron los datos correctamente", vbInformation, Me.Caption
Screen.MousePointer = vbDefault
Exit Function
ErrorTrans:
If NroTrans = 1 Then
    cn.BeginTrans
End If
MsgBox ERR.Description, vbCritical, Me.Caption
Screen.MousePointer = vbDefault


End Function

