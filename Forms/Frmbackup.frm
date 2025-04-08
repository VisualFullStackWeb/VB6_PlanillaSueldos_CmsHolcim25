VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form Frmbackup 
   BackColor       =   &H8000000B&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Backup de la Base de Datos"
   ClientHeight    =   1845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5400
   Icon            =   "Frmbackup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1845
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   Begin Threed.SSCommand CmdTrans 
      Height          =   495
      Left            =   4800
      TabIndex        =   2
      Top             =   1200
      Width           =   495
      _Version        =   65536
      _ExtentX        =   873
      _ExtentY        =   873
      _StockProps     =   78
      Picture         =   "Frmbackup.frx":030A
   End
   Begin MSComDlg.CommonDialog Box 
      Left            =   5640
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   495
      Left            =   4800
      TabIndex        =   1
      Top             =   480
      Width           =   495
      _Version        =   65536
      _ExtentX        =   873
      _ExtentY        =   873
      _StockProps     =   78
      Picture         =   "Frmbackup.frx":0464
   End
   Begin VB.TextBox txtArchivos 
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   4575
   End
   Begin VB.Label Opc 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   1320
      Width           =   45
   End
   Begin VB.Label Lblprocesa 
      AutoSize        =   -1  'True
      Caption         =   "Generar Backup"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3000
      TabIndex        =   4
      Top             =   1320
      Width           =   1395
   End
   Begin VB.Label Lbltit 
      AutoSize        =   -1  'True
      Caption         =   "Indique La Ruta y el Nombre Donde Se Guardara el Backup"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   5100
   End
End
Attribute VB_Name = "Frmbackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdTrans_Click()
If Trim(txtArchivos.Text) = "" Then
   MsgBox "Debe Indicar La Ruta y el nombre", vbInformation, "Bakcup"
   Exit Sub
End If
If Opc.Caption = "B" Then
   Backup
Else
   Restore
End If
End Sub

Private Sub Form_Activate()
If Opc = "B" Then
   Lbltit.Caption = "Indique La Ruta y el Nombre Donde Se Guardara el Backup"
   Lblprocesa = "Generar Backup"
Else
   Lbltit.Caption = "Indique La Ruta y el Nombre Para Restaurar el Backup"
   Lblprocesa = "Restaurar Backup"
End If
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
Me.Height = 2250
Me.Width = 5520
End Sub

Private Sub SSCommand1_Click()
Call AbrirFile("*.sql")
End Sub

Public Sub AbrirFile(pextension As String)

If Not Cuadro_Dialogo_Abrir(pextension) Then Exit Sub

  If UCase(Right(Box.FileName, 3)) <> UCase(Right(pextension, 3)) Then
            MsgBox "La Extensión de archivo no concuerda con el formato elegido", vbCritical, "Archivo Inválido"
            Exit Sub
   End If
        
   txtArchivos.Text = Box.FileName
   txtArchivos.ToolTipText = Box.FileName
   CmdTrans.Enabled = True

End Sub
Public Function Cuadro_Dialogo_Abrir(pextension As String) As Boolean
  
 On Error GoTo ErrHandler
    Box.Filter = "All Files (*.*)|*.*|Text Files (*.sql)|*.sql|"
   Box.FilterIndex = 2
   Box.FileName = "NombreArchivo"
   Box.InitDir = App.Path & ""
   Box.ShowOpen
   Dim pos As String
   If Box.FileName = "NombreArchivo" Then
        Cuadro_Dialogo_Abrir = False
    Else
        Cuadro_Dialogo_Abrir = True
    End If
   Exit Function

ErrHandler:
   Cuadro_Dialogo_Abrir = False
   Exit Function
End Function
Private Sub Backup()
Dim Mgrab As Integer
Dim March As String
Mgrab = MsgBox("Seguro de Generar Backup", vbYesNo + vbQuestion, TitMsg)
If Mgrab <> 6 Then Exit Sub

March = Dir(txtArchivos.Text)
If Trim(March) <> "" Then
   Mgrab = MsgBox("El Archivo " & March & " ya existe" & Chr(13) & "Desea Reemplazarlo", vbYesNo + vbQuestion, "Backup")
   If Mgrab <> 6 Then Exit Sub
End If
Screen.MousePointer = vbArrowHourglass
Call Shell(App.Path & "\Logs\BACKUP.BAT " & txtArchivos.Text, 0)
Screen.MousePointer = vbDefault
End Sub
Private Sub Restore()
Dim Mgrab As Integer
Dim March As String
Mgrab = MsgBox("Seguro de Restaurar Backup " & Chr(13) & "Se eliminaran los datos actuales y se restauraran los datos del Backup indicado", vbYesNo + vbQuestion, TitMsg)
If Mgrab <> 6 Then Exit Sub
March = Dir(txtArchivos.Text)
If Trim(March) = "" Then
   MsgBox "El Archivo " & March & " no Existe", vbCritical, "Backup"
   Exit Sub
End If
Mgrab = MsgBox("Seguro de Restaurar Backup " & Chr(13) & "Se eliminaran los datos actuales y se restauraran los datos del Backup indicado", vbYesNo + vbQuestion, TitMsg)
If Mgrab <> 6 Then Exit Sub

Screen.MousePointer = vbArrowHourglass
Call Shell(App.Path & "\Logs\RESTORE.BAT " & txtArchivos.Text, 0)
Screen.MousePointer = vbDefault
MsgBox "Debe Salir del Sistema Para Terminar de Actualizar", vbInformation, "Backup Restaurado"
End Sub

