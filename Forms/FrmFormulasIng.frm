VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form FrmFormulasIng 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Formulas de Calculo de Ingreso"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9450
   Icon            =   "FrmFormulasIng.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   9450
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox CmbTipoTrab 
      Height          =   315
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   600
      Width           =   2175
   End
   Begin VB.ComboBox Cmbtipo 
      Height          =   315
      ItemData        =   "FrmFormulasIng.frx":030A
      Left            =   4440
      List            =   "FrmFormulasIng.frx":0317
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   600
      Width           =   2055
   End
   Begin VB.ComboBox Cmbremunera 
      Height          =   315
      Left            =   7320
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   600
      Width           =   2055
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   2535
      Left            =   0
      TabIndex        =   3
      Top             =   960
      Width           =   9495
      Begin RichTextLib.RichTextBox TxtFormula 
         Height          =   2175
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   3836
         _Version        =   393217
         MaxLength       =   255
         TextRTF         =   $"FrmFormulasIng.frx":0349
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000017&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9495
      Begin VB.ComboBox Cmbcia 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   75
         Width           =   5775
      End
      Begin VB.Label Lblfecha 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Height          =   255
         Left            =   7440
         TabIndex        =   7
         Top             =   120
         Width           =   1815
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
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "T. Trabajador"
      Height          =   195
      Left            =   0
      TabIndex        =   10
      Top             =   600
      Width           =   960
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Tipo Concepto"
      Height          =   195
      Left            =   3360
      TabIndex        =   8
      Top             =   600
      Width           =   1050
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Concepto"
      Height          =   195
      Left            =   6600
      TabIndex        =   4
      Top             =   600
      Width           =   690
   End
End
Attribute VB_Name = "FrmFormulasIng"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim VRemunera As String
Dim VTipo As String
Dim VTipotrab As String
Private Sub Cmbcia_Click()
wcia = Trim(Right("00" & Cmbcia.ItemData(Cmbcia.ListIndex), 2))
Cmbremunera.ListIndex = -1
Cmbtipo.ListIndex = -1
VTipo = ""
VRemunera = ""
Call fc_Descrip_Maestros2("01055", "", CmbTipoTrab)
Procesa
End Sub

Private Sub Cmbremunera_Click()
VRemunera = fc_CodigoComboBox(Cmbremunera, 2)
Procesa
End Sub

Private Sub Cmbtipo_Click()
VTipo = Left(Cmbtipo.Text, 2)
Select Case VTipo
       Case Is = "02"
            SQL$ = "Select codinterno,descripcion from placonstante where cia='" & wcia & "' and tipomovimiento='02' and status<>'*' order by codinterno"
       Case Is = "01"
            SQL$ = "Select codinterno,descripcion from placonstante where cia='" & wcia & "' and tipomovimiento='03' and status<>'*' and aportacion<>0 order by codinterno"
       Case Is = "03"
            SQL$ = "Select codinterno,descripcion from placonstante where cia='" & wcia & "' and tipomovimiento='03' and status<>'*' and deduccion<>0 order by codinterno"
End Select
Cmbremunera.Clear
VRemunera = ""
If (fAbrRst(rs, SQL$)) Then
   If (Not rs.EOF) Then
      Do Until rs.EOF
         Cmbremunera.AddItem rs(1)
         Cmbremunera.ItemData(Cmbremunera.NewIndex) = rs(0)
         rs.MoveNext
       Loop
    End If
    If rs.State = 1 Then rs.Close
End If
Procesa
End Sub

Private Sub CmbTipoTrab_Click()
VTipotrab = fc_CodigoComboBox(CmbTipoTrab, 2)
Procesa
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
Me.Height = 3765
Me.Width = 9570
TxtFormula.Text = "2"
Lblfecha.Caption = Format(Date, "dd/mm/yyyy")
Call rCarCbo(Cmbcia, Carga_Cia, "C", "00")
Call rUbiIndCmbBox(Cmbcia, wcia, "00")
End Sub
Public Function Graba_FromulaIng()
Dim mformula As String
Dim i As Integer
Mgrab = MsgBox("Seguro de Grabar Formula", vbYesNo + vbQuestion, TitMsg)
If Mgrab <> 6 Then Exit Function
Screen.MousePointer = vbArrowHourglass
SQL$ = wInicioTrans
cn.Execute SQL$
SQL$ = "Update plaformulas set status='*' where cia='" & wcia & "' and tipo='" & VTipo & "' and codkey='" & VRemunera & "' and status<>'*'"
cn.Execute SQL$
mformula = ""
For i = 1 To Len(Trim(TxtFormula.Text))
    If Mid(TxtFormula.Text, i, 1) = "'" Then
       mformula = mformula & "?"
    Else
       mformula = mformula & Mid(TxtFormula.Text, i, 1)
    End If
Next i
SQL$ = "INSERT INTO plaformulas values('" & wcia & "','" & VTipo & "','" & VRemunera & "','" & Trim(mformula) & "'," & FechaSys & ",'','" & VTipotrab & "')"
cn.Execute SQL$

SQL$ = wFinTrans
cn.Execute SQL$
Screen.MousePointer = vbDefault
End Function
Private Sub Procesa()
Dim rs2 As ADODB.Recordset
Dim i, pos As Integer
Dim mformula As String
TxtFormula.Text = ""
SQL$ = "select * from plaformulas where cia='" & wcia & "' and tipo='" & VTipo & "' and codkey='" & VRemunera & "' and tipotrab='" & VTipotrab & "' and status<>'*'"
cn.CursorLocation = adUseClient
Set rs = New ADODB.Recordset
Set rs = cn.Execute(SQL$, 64)
If Not rs.RecordCount > 0 Then Exit Sub
rs.MoveFirst

mformula = ""
For i = 1 To Len(Trim(rs!Formula))
    If Mid(rs!Formula, i, 1) = "?" Then
       mformula = mformula & "'"
    Else
       mformula = mformula & Mid(rs!Formula, i, 1)
    End If
Next i

TxtFormula.Text = mformula
If rs.State = 1 Then rs.Close
End Sub
