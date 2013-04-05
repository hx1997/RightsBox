Attribute VB_Name = "modShellWait"
Option Explicit

Private Declare Function WaitForSingleObject Lib "kernel32.dll" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32.dll" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long

Private Const SYNCHRONIZE As Long = &H100000
Private Const WAIT_TIMEOUT As Long = 258&
Private Const PROCESS_TERMINATE As Long = (&H1)

Function ShellWait(ByVal lpApplicationName As String, ByVal WindowStyle As VbAppWinStyle, Optional ByVal dwMilliseconds As Integer = 20000, Optional ByVal KillOnTimeout As Boolean = False) As Long
    Dim dwPID As Long, hProcess As Long, ret As Long
    
    dwPID = Shell(lpApplicationName, WindowStyle)
    
    If (dwPID) Then
        hProcess = OpenProcess(SYNCHRONIZE Or PROCESS_TERMINATE, 0, dwPID)
        If (hProcess) Then
            ret = WaitForSingleObject(hProcess, dwMilliseconds)
            If ret = WAIT_TIMEOUT Then
                If (KillOnTimeout) Then Call TerminateProcess(hProcess, WAIT_TIMEOUT)
            End If
            
            Call CloseHandle(hProcess)
        End If
    End If
End Function
