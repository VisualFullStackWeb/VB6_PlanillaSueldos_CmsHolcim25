VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form FrmCierre 
   Caption         =   "Cierre de Planilla"
   ClientHeight    =   2100
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4515
   Icon            =   "FrmCierre.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2100
   ScaleWidth      =   4515
   Begin VB.CommandButton BtnCierre 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   4335
   End
   Begin VB.ComboBox Cmbmes 
      Height          =   315
      ItemData        =   "FrmCierre.frx":030A
      Left            =   1740
      List            =   "FrmCierre.frx":0332
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   240
      Width           =   2595
   End
   Begin VB.TextBox Txtano 
      Enabled         =   0   'False
      Height          =   285
      Left            =   735
      TabIndex        =   0
      Top             =   240
      Width           =   705
   End
   Begin VB.Label LblCierre 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   4335
   End
   Begin MSForms.SpinButton SpinButton1 
      Height          =   300
      Left            =   1455
      TabIndex        =   3
      Top             =   240
      Width           =   255
      Size            =   "450;529"
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Periodo"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   540
   End
End
Attribute VB_Name = "FrmCierre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Carga_Cierre()
Dim Rq As ADODB.Recordset
Sql$ = "SELECT Fec_Crea fROM pla_cierre where cia='" & wcia & "' and Ano=" & Txtano.Text & " and Mes=" & Cmbmes.ListIndex + 1 & " and Status<>'*'"
If fAbrRst(Rq, Sql) Then
   LblCierre.Caption = "Periodo Cerrado"
   BtnCierre.Caption = "Abrir Periodo"
   LblCierre.ForeColor = vbRed
Else
   LblCierre.Caption = "Periodo Abierto"
   BtnCierre.Caption = "Cerrar Periodo"
   LblCierre.ForeColor = vbBlue
End If
End Sub

Private Sub BtnCierre_Click()
If BtnCierre.Caption = "Abrir Periodo" Then
   Sql$ = "Update pla_cierre set status='*' where Cia='" & wcia & "' and ano=" & Txtano.Text & " and mes=" & Cmbmes.ListIndex + 1 & " and status<>'*'"
   cn.Execute Sql, 64
   LblCierre.Caption = "Periodo Abierto"
   BtnCierre.Caption = "Cerrar Periodo"
   LblCierre.ForeColor = vbBlue
Else
   Sql$ = "Insert Into pla_cierre values ('" & wcia & "'," & Txtano.Text & "," & Cmbmes.ListIndex + 1 & ",'" & wuser & "',getdate(),'')"
   cn.Execute Sql, 64
   LblCierre.Caption = "Periodo Cerrado"
   BtnCierre.Caption = "Abrir Periodo"
   LblCierre.ForeColor = vbRed
End If
End Sub

Private Sub Cmbmes_Click()
Carga_Cierre
End Sub

Private Sub Form_Load()
Me.Top = 0: Me.Left = 0
Me.Width = 4635: Me.Height = 2610
Txtano.Text = Year(Date)
Cmbmes.ListIndex = Month(Date) - 1
End Sub

Private Sub SpinButton1_SpinDown()
Txtano.Text = Txtano.Text - 1
Carga_Cierre
End Sub

Private Sub SpinButton1_SpinUp()
Txtano.Text = Txtano.Text + 1
Carga_Cierre
End Sub
