VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form Formimpri 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "» Impresora"
   ClientHeight    =   3375
   ClientLeft      =   2805
   ClientTop       =   2910
   ClientWidth     =   5160
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Formimpri.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3375
   ScaleWidth      =   5160
   Begin Threed.SSCommand SSCommand3 
      Height          =   600
      Left            =   3225
      TabIndex        =   7
      ToolTipText     =   "Salir"
      Top             =   2670
      Width           =   705
      _Version        =   65536
      _ExtentX        =   1235
      _ExtentY        =   1058
      _StockProps     =   78
      BevelWidth      =   1
      Picture         =   "Formimpri.frx":058A
   End
   Begin Threed.SSCommand SSCommand2 
      Height          =   600
      Left            =   2040
      TabIndex        =   6
      ToolTipText     =   "Imprimir"
      Top             =   2640
      Width           =   705
      _Version        =   65536
      _ExtentX        =   1235
      _ExtentY        =   1058
      _StockProps     =   78
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
      Picture         =   "Formimpri.frx":0B24
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   600
      Left            =   4335
      TabIndex        =   5
      ToolTipText     =   "Visualizar"
      Top             =   2685
      Width           =   705
      _Version        =   65536
      _ExtentX        =   1235
      _ExtentY        =   1058
      _StockProps     =   78
      BevelWidth      =   1
      Picture         =   "Formimpri.frx":10BE
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   210
      Left            =   2040
      TabIndex        =   4
      Top             =   1935
      Width           =   1005
      _Version        =   65536
      _ExtentX        =   1773
      _ExtentY        =   370
      _StockProps     =   15
      Caption         =   "Impresoras"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   0
      BevelOuter      =   1
      Autosize        =   3
   End
   Begin VB.DriveListBox Drive1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1815
   End
   Begin VB.ComboBox Cmbprint 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   2220
      Width           =   3015
   End
   Begin VB.FileListBox FILE 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Left            =   2040
      Pattern         =   "*.TXT;*.PRN"
      TabIndex        =   1
      Top             =   105
      Width           =   3015
   End
   Begin VB.DirListBox Dir 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2730
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Lblver 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   3480
      TabIndex        =   8
      Top             =   1935
      Visible         =   0   'False
      Width           =   1545
   End
End
Attribute VB_Name = "Formimpri"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RUTA As String
 
Private Sub Command1_Click()

End Sub

Private Sub Command2_Click()

End Sub

Private Sub Dir_Change()
    FILE.FileName = Dir
    RUTA = Dir.Path
End Sub

Private Sub Drive1_Change()
On Error GoTo mensaje
    Dir.Path = Drive1
    Exit Sub
mensaje:
    Call MsgBox("La unidad no esta disponible", vbCritical, "IMPRESORA")
    Exit Sub
End Sub

Private Sub FILE_Click()
Lblver = Dir & "\" & FILE
End Sub

Private Sub Form_Activate()
'Dim X As Printer
'Cmbprint.Clear
'For Each X In Printers
'    Cmbprint.AddItem X.DeviceName
'Next
'Cmbprint.Text = Printer.DeviceName
'FILE.SetFocus
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
Me.Width = 5250
Me.Height = 3750
Dim x As Printer
Dir.Path = App.Path & "\REPORTS"
If wPrintFile <> "" Then FILE.FileName = wPrintFile

If wGrupoPla = "01" Then
   Dim cItem As Integer
   Dim cLocal As String
   Dim cPuerto As String
   cLocal = "0"
   For cItem = 1 To 10
       cPuerto = Trim((Trim(aPrinloc(cItem)) & " "))
       If Len(cPuerto) > 1 Then
          cLocal = "1"
          Cmbprint.AddItem cPuerto
      End If
   Next
   If cLocal = "0" Then
       For Each x In Printers
           Cmbprint.AddItem x.DeviceName
       Next
       Cmbprint.Text = Printer.DeviceName
   Else
       Cmbprint.ListIndex = 0
   End If
   'FILE.SetFocus
Else
  
   Cmbprint.Clear
   For Each x In Printers
       Cmbprint.AddItem x.DeviceName
   Next

   Cmbprint.Text = Printer.DeviceName
   Exit Sub
End If

End Sub

Private Sub SSCommand1_Click()
Load Frmverarchivo
Frmverarchivo.Show
Frmverarchivo.ZOrder 0
Frmverarchivo.Vertexto.Navigate (Lblver.Caption)
End Sub

Private Sub SSCommand2_Click()
'Dim MB As Boolean
'MB = True
''*****VERIFICAR ESTA INFORMACION IMPRESION - NO IMPRIME*****************
'If UCase(FILE.FileName) = "BOLETAS.TXT" Or UCase(FILE.FileName) = "QUINCENA.TXT" Or UCase(FILE.FileName) = "CERTCTS.TXT" Then MB = True
'If Dir.Path = "C:\" Then
'    Call DestinoPort(Cmbprint.Text, Dir.Path + FILE.FileName, MB)
'Else
'    Call DestinoPort(Cmbprint.Text, Dir.Path + "\" + FILE.FileName, MB)
'End If
'Unload Formimpri
'***********************************************************************

If Dir.Path = "C:\" Then
    Call DestinoPort(Cmbprint.Text, Dir.Path + FILE.FileName)
Else
    Call DestinoPort(Cmbprint.Text, Dir.Path + "\" + FILE.FileName)
End If
Unload Formimpri

End Sub

Private Sub SSCommand3_Click()
    Unload Me
End Sub
