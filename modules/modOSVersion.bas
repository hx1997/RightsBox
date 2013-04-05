Attribute VB_Name = "modOSVersion"
Option Explicit

Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Type OSVERSIONINFO
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As Long
        szCSDVersion As String * 128
End Type

Sub GetOSVer(os As OSVERSIONINFO)
    os.dwOSVersionInfoSize = Len(os)
    GetVersionEx os
End Sub

Function IsVistaOrLater() As Boolean
    Dim os As OSVERSIONINFO
    GetOSVer os
    If os.dwMajorVersion = 6 Then
        IsVistaOrLater = True
    Else
        IsVistaOrLater = False
    End If
End Function
