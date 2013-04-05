Attribute VB_Name = "modError"
Option Explicit

Private Declare Function FormatMessage Lib "kernel32.dll" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const FORMAT_MESSAGE_IGNORE_InsertS = &H200

Public Function GetLastDllErr(ByVal lErr As Long) As String
    Dim sReturn As String
    sReturn = String$(256, 32)
    FormatMessage FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_InsertS, 0&, lErr, 0&, sReturn, Len(sReturn), ByVal 0
    sReturn = Trim(sReturn)
    GetLastDllErr = sReturn
End Function

Public Function RBOX(ByVal RBOXErr As Long) As String
    RBOX = "RBOX" & RBOXErr
    Select Case RBOXErr
        Case 1
            RBOX = RBOX & " RunSafer failed."
        Case 2  'CreateProcessAsUser
            RBOX = RBOX & " Cannot start RBStart.exe."
        Case 3  'CreateJobObject
            RBOX = RBOX & " Cannot create a job object."
        Case 4  'AssignProcessToJobObject
            RBOX = RBOX & " Cannot assign the process to a job."
        Case 5  'CreateIoCompletionPort
            RBOX = RBOX & " Cannot create completion port."
        Case 6  'SetInformationJobObject
            RBOX = RBOX & " Cannot set security limits on the process."
        Case 7  'SetProcessLowIL
            RBOX = RBOX & " Cannot lower process integrity level."
        Case 8  'RestrictedToken
            RBOX = RBOX & " Cannot create a restricted token."
        Case 9  'SaferUsers
            RBOX = RBOX & " Cannot compute token from SAFER level."
        Case 10
            RBOX = RBOX & " Invalid rule action."
        Case 11
            RBOX = RBOX & " Invalid rule type."
        Case 12
            RBOX = RBOX & " Cannot restrict access to files."
        Case 13
            RBOX = RBOX & " Cannot restrict access to registry keys."
    End Select
End Function
