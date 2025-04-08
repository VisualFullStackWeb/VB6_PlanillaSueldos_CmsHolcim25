Attribute VB_Name = "ModVersionWin"
   
Public Declare Function GetVersionExA Lib "kernel32" _
(lpVersionInformation As OSVERSIONINFO) As Integer

Public Type OSVERSIONINFO
dwOSVersionInfoSize As Long
dwMajorVersion As Long
dwMinorVersion As Long
dwBuildNumber As Long
dwPlatformId As Long
szCSDVersion As String * 128
End Type

Public Function getVersion() As String
Dim osinfo As OSVERSIONINFO
Dim retvalue As Integer

osinfo.dwOSVersionInfoSize = 148
osinfo.szCSDVersion = Space$(128)
retvalue = GetVersionExA(osinfo)
With osinfo
Select Case .dwPlatformId
Case 1
    Select Case .dwMinorVersion
    Case 0
        getVersion = "Windows 95"
    Case 10
        getVersion = "Windows 98"
    Case 90
        getVersion = "Windows Mellinnium"
    End Select
Case 2
    Select Case .dwMajorVersion
    Case 3
        getVersion = "Windows NT 3.51"
    Case 4
        getVersion = "Windows NT 4.0"
    Case 5
        If .dwMinorVersion = 0 Then
            getVersion = "Windows 2000"
        Else
            getVersion = "Windows XP"
        End If
    Case 6
        getVersion = "Windows Vista"
    Case 7
        getVersion = "Windows 7"
    End Select

Case Else
    
    getVersion = "Failed"
End Select
End With
End Function



