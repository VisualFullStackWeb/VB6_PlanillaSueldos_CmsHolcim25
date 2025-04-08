VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00EEFDFD&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1755
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5595
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00C000C0&
   FillStyle       =   6  'Cross
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1755
   ScaleWidth      =   5595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.FileListBox File1 
      Height          =   2235
      Left            =   990
      TabIndex        =   5
      Top             =   2025
      Width           =   1680
   End
   Begin ComctlLib.ProgressBar Barra 
      Height          =   555
      Left            =   270
      TabIndex        =   2
      Top             =   855
      Visible         =   0   'False
      Width           =   5325
      _ExtentX        =   9393
      _ExtentY        =   979
      _Version        =   327682
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.PictureBox picLogo 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   9375
      Left            =   0
      ScaleHeight     =   9375
      ScaleWidth      =   255
      TabIndex        =   4
      Top             =   0
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   4725
      Picture         =   "Form1.frx":030A
      Stretch         =   -1  'True
      Top             =   270
      Width           =   615
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Caption         =   "Planillas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   5580
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Espere un Momento...."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1485
      Width           =   5415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Actualizando Sistema de Planilla"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   765
      TabIndex        =   0
      Top             =   480
      Width           =   3555
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sMensaje As String

'Para Registrar Dll y Ocx  Para tu Libro
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Any, ByVal wParam As Any, ByVal lParam As Any) As Long
Private Const ERROR_SUCCESS = &H0


Dim cL As New cLogo

Dim NomExe As String
Dim DirExe As String
Dim DirExeSer As String
Dim DirSer As String
Dim DirRpt As String
Dim Mio As Date
Dim Red As Date
Dim Size As Long


Function CopyFile(Src As String, Dst As String) As Single

Static Buf$
Dim BTest!, FSize! 'declare the needed variables
Dim Chunk%, F1%, F2%

Const BUFSIZE = 1024 'set the buffer size

'On Error GoTo FileCopyError 'incase of error goto this label
F1 = FreeFile 'returns file number available
Open Src For Binary As F1 'open the source file
F2 = FreeFile 'returns file number available
Open Dst For Binary As F2 'open the destination file
 
FSize = LOF(F1)
BTest = FSize - LOF(F2)

Do
If BTest < BUFSIZE Then
   Chunk = BTest
Else
   Chunk = BUFSIZE
End If
      
Buf = String(Chunk, " ")
Get F1, , Buf
Put F2, , Buf
BTest = FSize - LOF(F2)
DoEvents
Barra.Value = (100 - Int(100 * BTest / FSize)) 'advance the progress bar as the file is copied

Loop Until BTest = 0
Close F1 'closes the source file
Close F2 'closes the destination file
'CopyFile = FSize
Barra.Value = 0 'returns the progress bar to zero
Exit Function 'exit the procedure

FileCopyError: 'file copy error label
MsgBox "Error en la actualización del Sistema...", vbCritical
Close F1 'closes the source file
Close F2 'closes the destination file
Exit Function 'exit the procedure

End Function

Private Sub Form_Activate()
Dim i As Integer
If App.PrevInstance = True Then MsgBox "Ya se esta Ejecutando el Sistema !!!", vbOKOnly + vbExclamation, "Consulte con Sistemas. ", vbOKOnly + vbCritical, "SysControl": End

On Error GoTo Errores

'Copiamos Ejecutable
NomExe = "SisPlaRedRoda.exe"
DirExe = "C:\SisPlaRedRoda\"

'Ruta principal
gsFileINI$ = "\\10.10.10.33\SISCOMACSA$\Conexiones\Configuracion Actualizadores\Log.INI"
DirSer = UCase(ReadINI("Actu", "Ruta_Actu"))

'Leemos esta ruta por si no encontramos la ruta principal
If DirSer = "" Then
    gsFileINI$ = "\\10.10.10.35\SISCOMACSA$\Conexiones\Configuracion Actualizadores\Log.INI"
    DirSer = UCase(ReadINI("Actu", "Ruta_Actu"))
End If
If DirSer = "" Then
    MsgBox "No se encuentra el archivo de configuración!!! Consulte con sistemas.", vbOKOnly + vbExclamation, "Actualizador del sistema"
    Unload Me
    Exit Sub
End If

Screen.MousePointer = vbHourglass

'If Dir(DirExe) = "" Then
'    MsgBox "No tiene acceso a la ruta: " & Chr(13) & _
'    DirExe & Chr(13) & "No se podrá iniciar el Sistema. " & Chr(13) & "Consulte con sistemas. ", vbOKOnly + vbCritical, "SysControl - Comuniquese con Sistemas "
'    End
'End If



If Dir(DirExe & NomExe) = "" Then
    FileCopy DirSer & NomExe, DirExe & NomExe
    'Barra.Value = CopyFile(DirExe & NomExe, DirExe & NomExe)
    'sMensaje = sMensaje & "No se encuentro el archivo " & NomExe & vbCrLf
Else
    If Fc_IsFileExist(DirExe & NomExe) Then
        Mio = CDate(FileSystem.FileDateTime(DirExe & NomExe))
    Else
        Mio = CDate("01/01/2000")
    End If
    Red = CDate(FileSystem.FileDateTime(DirSer & NomExe))
    
    If (Red > Mio) Then
        Barra.Visible = True
        If Fc_IsFileExist(DirExe & NomExe) Then Kill DirExe & NomExe
        Barra.Value = CopyFile(DirSer & NomExe, DirExe & NomExe)
    End If
End If


If Fc_IsFileExist(DirExe & NomExe) Then Shell DirExe & NomExe, vbNormalFocus
Unload Me
Screen.MousePointer = vbDefault

Exit Sub
Errores:
    MsgBox "Se Produjo el Siguiente Error : " & Err.Number & " - " & Err.Description, vbOKOnly + vbCritical, "SysControl - Comuniquese con Sistemas"

    Unload Me
End Sub

Private Sub Form_Load()
On Error GoTo Errores


cL.DrawingObject = picLogo
cL.Caption = "Planillas"

Exit Sub
Errores:
MsgBox "Se Produjo el Siguiente Error : " & Err.Number & " - " & Err.Description, vbOKOnly + vbCritical, "SysControl - Comuniquese con Sistemas"
End Sub

Private Sub Form_Resize()
On Error Resume Next
    picLogo.Height = Me.ScaleHeight
    On Error GoTo 0
    cL.Draw
End Sub


Private Function RegisterCOMDLL(hwnd As Long, ByVal sPath As String, ByVal Archivo As String, bRegister As Boolean)
On Error Resume Next
Dim lb As Long, pa As Long
Dim DllServerPath As String

If Mid(Archivo, InStr(1, Archivo, ".", 1) + 1, 3) = "ocx" Or Mid(Archivo, InStr(1, Archivo, ".", 1) + 1, 3) = "dll" Then
    DllServerPath = sPath & Archivo
    lb = LoadLibrary(DllServerPath)
    If bRegister Then
        pa = GetProcAddress(lb, "DllRegisterServer")
    Else
        pa = GetProcAddress(lb, "DllUnregisterServer")
    End If

    If CallWindowProc(pa, hwnd, ByVal 0&, ByVal 0&, ByVal 0&) = ERROR_SUCCESS Then
    Else
    End If
    'unmap the library's address
    FreeLibrary lb
End If

End Function


