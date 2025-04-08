VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacceso 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "» Inicio de Sesión «"
   ClientHeight    =   2415
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5430
   Icon            =   "Frmacceso.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   5430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Txtclave 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   3555
      MaxLength       =   20
      PasswordChar    =   "l"
      TabIndex        =   1
      Top             =   1230
      Width           =   1470
   End
   Begin VB.TextBox Txtuser 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3555
      MaxLength       =   20
      TabIndex        =   0
      Top             =   735
      Width           =   1470
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © 2011 Roda S.A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Left            =   225
      TabIndex        =   7
      Top             =   1995
      Width           =   1845
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RODA SA - SOFTWARE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   555
      TabIndex        =   6
      Top             =   120
      Width           =   3435
   End
   Begin VB.Image Img 
      Height          =   240
      Index           =   1
      Left            =   3225
      Picture         =   "Frmacceso.frx":030A
      Top             =   1215
      Width           =   240
   End
   Begin VB.Image Img 
      Height          =   240
      Index           =   0
      Left            =   3225
      Picture         =   "Frmacceso.frx":0894
      Top             =   705
      Width           =   240
   End
   Begin VB.Shape Shape 
      BorderColor     =   &H00FFFFFF&
      Height          =   2190
      Index           =   2
      Left            =   105
      Top             =   105
      Width           =   5205
   End
   Begin VB.Image Img 
      Height          =   240
      Index           =   2
      Left            =   4935
      Picture         =   "Frmacceso.frx":0E1E
      Top             =   150
      Width           =   240
   End
   Begin VB.Shape Shape 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000001&
      Height          =   405
      Index           =   1
      Left            =   3120
      Top             =   1155
      Width           =   2040
   End
   Begin MSForms.CommandButton SSCommand4 
      Height          =   435
      Left            =   3900
      TabIndex        =   5
      Top             =   1710
      Width           =   1275
      ForeColor       =   4210752
      Caption         =   "  Cancelar"
      PicturePosition =   327683
      Size            =   "2249;767"
      Picture         =   "Frmacceso.frx":13A8
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton SSCommand3 
      Height          =   435
      Left            =   2520
      TabIndex        =   4
      Top             =   1710
      Width           =   1275
      ForeColor       =   4210752
      Caption         =   "   Aceptar"
      PicturePosition =   327683
      Size            =   "2249;767"
      Picture         =   "Frmacceso.frx":1942
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Clave"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   210
      Left            =   2325
      TabIndex        =   3
      Top             =   1095
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   210
      Left            =   2310
      TabIndex        =   2
      Top             =   600
      Width           =   630
   End
   Begin VB.Shape Shape 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000001&
      Height          =   405
      Index           =   0
      Left            =   3120
      Top             =   645
      Width           =   2040
   End
   Begin VB.Image Img 
      Height          =   2500
      Index           =   3
      Left            =   0
      Picture         =   "Frmacceso.frx":1EDC
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5500
   End
End
Attribute VB_Name = "Frmacceso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()
'Txtclave.SetFocus
End Sub
Private Sub Form_Load()
Txtuser.Text = ""
Label4.Caption = "Versión " & "[ " & App.Major & "." & App.Minor & "." & App.Revision & " ]" & "» Sistema de Planilla" & Space(40) & "Versión " & "[ " & App.Major & "." & App.Minor & "." & App.Revision & " ]"
'Txtuser.Text = wuser
End Sub

Private Sub SSCommand1_Click()
    If Trim(Txtuser.Text) = "" Then
        MsgBox "INGRESE EL USUARIO", vbInformation, ""
        Exit Sub
    End If

    LoginSucceeded = True
    wuser = Txtuser.Text
    wclave = Txtclave.Text
    Unload Me
End Sub


Private Sub SSCommand3_Click()
    LoginSucceeded = True
    If Trim(Txtuser) <> "" Then Txtuser.Text = UCase(Trim(Txtuser.Text))
    wuser = Trim(Txtuser.Text)
    wclave = Trim(Txtclave.Text)
    wCancelsis = False
    wNamePC = Trim(NamePC)
    wNamePC = Mid(wNamePC, 1, Len(wNamePC) - 1)
    Unload Me
End Sub

Private Sub SSCommand4_Click()
wCancelsis = True
LoginSucceeded = False
Unload Me
End Sub

Private Sub Txtclave_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SSCommand1_Click
End Sub

Private Sub Txtuser_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Txtclave.SetFocus
End Sub
