Attribute VB_Name = "modWindow"
Option Explicit

Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal flgs As Long) As Long
Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Declare Function GetWindowWord Lib "user32" (ByVal hWnd As Long, ByVal wIndx As Long) As Long
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal wIndx As Long) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpSting As String, ByVal nMaxCount As Long) As Long
Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal insaft As Long, ByVal x%, ByVal y%, ByVal cx%, ByVal cy%, ByVal flgs As Long) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Any) As Long
Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Declare Function GetDesktopWindow Lib "user32.dll" () As Long
Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Const WS_MINIMIZE = &H20000000
Const HWND_TOP = 0
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_SHOWWINDOW = &H40
Const GW_HWNDFIRST = 0
Const GW_HWNDNEXT = 2
Const GWL_STYLE = (-16)
Const SW_RESTORE = 9

Const WS_VISIBLE = &H10000000
Const WS_BORDER = &H800000

Const WS_CLIPSIBLINGS = &H4000000
Const WS_THICKFRAME = &H40000
Const WS_GROUP = &H20000
Const WS_TABSTOP = &H10000

Public Const LB_SETHORIZONTALEXTENT = &H194
Public Const LVM_FIRST = &H1000
Public Const LVM_SETEXTENDEDLISTVIEWSTYLE = &H1000 + 54
Public Const LVS_EX_FULLROWSELECT = &H20

Sub HxGetHWndByPID(ByVal ProcessId As Long, objObject As Collection, ByVal hWnd As Long)
    Dim hwCurr As Long
    Dim intLen As Long
    Dim strTitle As String
    
    hwCurr = GetWindow(hWnd, GW_HWNDFIRST)
    
    Do While hwCurr
        If IsWindowVisible(GetParent(hwCurr)) <> 0 Then
            Dim PID As Long
            GetWindowThreadProcessId GetParent(hwCurr), PID
            If PID = ProcessId Then
                objObject.Add GetParent(hwCurr)
            End If
        End If
        hwCurr = GetWindow(hwCurr, GW_HWNDNEXT)
    Loop
End Sub
