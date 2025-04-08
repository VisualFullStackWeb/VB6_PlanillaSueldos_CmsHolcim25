Attribute VB_Name = "Module1"
Global gsFileINI As String
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long

Type SHITEMID
     cb As Long
     abID As Byte
End Type

Type ITEMIDLIST
     mkid As SHITEMID
End Type

Type BROWSEINFO
     hOwner As Long
     pidlRoot As Long
     pszDisplayName As String
     lpszTitle As String
     ulFlags As Long
     lpfn As Long
     lParam As Long
     iImage As Long
End Type

Public Const NOERROR = 0

Public Const BIF_RETURNONLYFSDIRS = &H1
Public Const BIF_DONTGOBELOWDOMAIN = &H2
Public Const BIF_STATUSTEXT = &H4
Public Const BIF_RETURNFSANCESTORS = &H8
Public Const BIF_BROWSEFORCOMPUTER = &H1000
Public Const BIF_BROWSEFORPRINTER = &H2000

'******************************************************************
' IsFileExist : Verificar la existencia de un archivo
'******************************************************************
Public Function Fc_IsFileExist(ByVal strFileName As String) As Boolean
    Dim Msg As String
    
    On Error GoTo CheckError
        Fc_IsFileExist = (Dir(strFileName) <> "")
        Exit Function

CheckError:
    Const mnErrDiskNotReady = 71, _
    mnErrDeviceUnavailable = 68
    
    If (Err.Number = mnErrDiskNotReady) Then
        Msg = "Inserte un disco en la unidad y "
        Msg = Msg & "cierre la puerta de la misma."
        If MsgBox(Msg, vbExclamation & vbOKCancel) = _
        vbOK Then
            Resume
        Else
            Resume Next
        End If
    ElseIf Err.Number = mnErrDeviceUnavailable Then
        Msg = "Esta unidad o ruta de acceso no existe: "
        Msg = Msg & strFileName
        MsgBox Msg, vbExclamation
        Resume Next
    Else
        Msg = "Error inesperado nº" & Str(Err.Number)
        Msg = Msg & " : " & Err.Description
        MsgBox Msg, vbCritical
        Stop
    End If
    Resume
End Function
Function ReadINI(cSection$, cKeyName$) As String
   'Se le quita un caracter porque el último es fin de linea
    Dim sRet As String
    Dim Longitud%
    Dim Def$
    
    sRet = String(255, " ")
    Longitud = Len(sRet)
    
   Call GetPrivateProfileString(cSection$, ByVal cKeyName, "", ByVal sRet$, ByVal Len(sRet), gsFileINI)
   
   ReadINI = left(Trim$(sRet), Len(Trim(sRet)) - 1)
    
End Function
