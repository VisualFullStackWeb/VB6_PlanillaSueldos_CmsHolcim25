VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form Frmcalculoscrs 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tasa de Calculo de Seguro Complementario"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6630
   Icon            =   "Frmcalculoscrs.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   6630
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox Cmbseguro 
      Height          =   315
      ItemData        =   "Frmcalculoscrs.frx":030A
      Left            =   1200
      List            =   "Frmcalculoscrs.frx":0314
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   720
      Width           =   3975
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   2175
      Left            =   0
      TabIndex        =   4
      Top             =   1080
      Width           =   6615
      Begin MSDataGridLib.DataGrid Dgrdseguro 
         Height          =   1935
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   3413
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
         ColumnCount     =   4
         BeginProperty Column00 
            DataField       =   "distribucion"
            Caption         =   "Distribucion"
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
            DataField       =   "importe"
            Caption         =   "Importe"
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
            DataField       =   "status"
            Caption         =   "Status"
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
               ColumnWidth     =   4199.811
            EndProperty
            BeginProperty Column01 
               Alignment       =   1
               ColumnWidth     =   1124.787
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   510.236
            EndProperty
            BeginProperty Column03 
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000017&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6615
      Begin VB.ComboBox Cmbcia 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   120
         Width           =   3975
      End
      Begin VB.Label Lblfecha 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Height          =   255
         Left            =   5280
         TabIndex        =   3
         Top             =   120
         Width           =   1215
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
         Left            =   240
         TabIndex        =   1
         Top             =   120
         Width           =   840
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Seguro"
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   720
      Width           =   510
   End
End
Attribute VB_Name = "Frmcalculoscrs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim VTipoSeguro As String
Dim rsscrt As New Recordset
Private Sub Cmbcia_Click()
wcia = Trim(Right("00" & Cmbcia.ItemData(Cmbcia.ListIndex), 2))
Cmbseguro.ListIndex = -1
Procesa
End Sub

Private Sub Cmbseguro_Click()
VTipoSeguro = Left(Cmbseguro.Text, 2)
Procesa
End Sub

Private Sub Dgrdseguro_AfterColEdit(ByVal ColIndex As Integer)
Select Case ColIndex
       Case Is = 2
            If Dgrdseguro.Columns(2) <> "P" And Dgrdseguro.Columns(2) <> "F" Then
               MsgBox "Status Solo Puede ser [P]orcentaje o [F]ijo", vbCritical, TitMsg
               Dgrdseguro.Columns(2) = ""
            End If
End Select
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
Me.Width = 6750
Me.Height = 3660
Lblfecha.Caption = Format(Date, "dd/mm/yyyy")
Crea_Rs
Call rCarCbo(Cmbcia, Carga_Cia, "C", "00")
Call rUbiIndCmbBox(Cmbcia, wcia, "00")
End Sub
Private Sub Crea_Rs()
    If rsscrt.State = 1 Then rsscrt.Close
    rsscrt.Fields.Append "distribucion", adChar, 45, adFldIsNullable
    rsscrt.Fields.Append "importe", adCurrency, 18, adFldIsNullable
    rsscrt.Fields.Append "status", adChar, 1, adFldIsNullable
    rsscrt.Fields.Append "codigo", adChar, 2, adFldIsNullable
    rsscrt.Open
    Set Dgrdseguro.DataSource = rsscrt
End Sub

Private Sub Procesa()
Dim wciamae As String
Dim rs2 As ADODB.Recordset
If Cmbseguro.ListIndex < 0 Then Exit Sub
wciamae = Determina_Maestro("01074")
SQL$ = "Select * from maestros_2 where status<>'*'"
SQL$ = SQL$ & wciamae
cn.CursorLocation = adUseClient
Set rs = New ADODB.Recordset
Set rs = cn.Execute(SQL$, 64)
If rsscrt.RecordCount > 0 Then
   rsscrt.MoveFirst
   Do While Not rsscrt.EOF
      rsscrt.Delete
      rsscrt.MoveNext
   Loop
End If
If Not rs.RecordCount > 0 Then MsgBox "No Existen Conceptos de Distribucion", vbCritical, TitMsg: Exit Sub
rs.MoveFirst
Do While Not rs.EOF
   rsscrt.AddNew
   rsscrt!codigo = Trim(rs!cod_maestro2)
   rsscrt!distribucion = rs!descrip
   
   SQL$ = "select * from platasaplanilla where cia='" & wcia & "' and modulo='01' and status<>'*' and tipomov='" & VTipoSeguro & "' and codinterno='" & rsscrt!codigo & "'"
   cn.CursorLocation = adUseClient
   Set rs2 = New ADODB.Recordset
   Set rs2 = cn.Execute(SQL$, 64)
   If rs2.RecordCount > 0 Then
      rsscrt!importe = rs2!tasa1
      rsscrt!status = rs2!status
   Else
      rsscrt!importe = "0.00"
      rsscrt!status = ""
   End If
   If rs2.State = 1 Then rs2.Close
   rs.MoveNext
Loop
If rs.State = 1 Then rs.Close

End Sub
Public Function Graba_Seguro()
Mgrab = MsgBox("Seguro de Grabar Seteo de Seguro Complementario", vbYesNo + vbQuestion, TitMsg)
If Mgrab <> 6 Then Exit Function
Screen.MousePointer = vbArrowHourglass
SQL$ = wInicioTrans
cn.Execute SQL$
SQL$ = "Update platasaplanilla set status='*' where cia='" & wcia & "' and modulo='01' and status<>'*' and tipomov='" & VTipoSeguro & "'"
cn.Execute SQL$
If rsscrt.RecordCount > 0 Then rsscrt.MoveFirst
Do While Not rsscrt.EOF
   SQL$ = "INSERT INTO platasaplanilla values('" & wcia & "','" & VTipoSeguro & "','01','" & rsscrt!codigo & "', " _
        & "" & CCur(rsscrt!importe) & ",'','" & rsscrt!status & "','" & wuser & "'," & FechaSys & ",'',null)"
   cn.Execute SQL$
   rsscrt.MoveNext
Loop
SQL$ = wFinTrans
cn.Execute SQL$
Screen.MousePointer = vbDefault
End Function
