VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.Ocx"
Begin VB.Form frmRightsBox 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RightsBox"
   ClientHeight    =   6510
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   9975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   9975
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer5 
      Interval        =   100
      Left            =   1920
      Top             =   6120
   End
   Begin VB.Timer Timer4 
      Interval        =   1000
      Left            =   1440
      Top             =   6120
   End
   Begin VB.Timer Timer3 
      Interval        =   1000
      Left            =   960
      Top             =   6120
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   480
      Top             =   6120
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   0
      Top             =   6120
   End
   Begin VB.CommandButton cmdRightsBox 
      BackColor       =   &H000000FF&
      Caption         =   "RightsBox"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   1215
   End
   Begin ComctlLib.ListView lvRestricted 
      Height          =   5895
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   10398
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Process"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "PID"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Window Title"
         Object.Width           =   10583
      EndProperty
   End
   Begin VB.Menu mnuSandbox 
      Caption         =   "Sandbox"
      Visible         =   0   'False
      Begin VB.Menu mnuRun 
         Caption         =   "&Run"
      End
      Begin VB.Menu mnuRunAs 
         Caption         =   "Run as..."
         Begin VB.Menu mnuBasicUser 
            Caption         =   "Partially Limited"
         End
         Begin VB.Menu mnuLowRights 
            Caption         =   "Limited"
         End
         Begin VB.Menu mnuLimited 
            Caption         =   "Strictly Limited"
         End
         Begin VB.Menu mnuUntrusted 
            Caption         =   "Untrusted"
         End
      End
      Begin VB.Menu mnuTerminate 
         Caption         =   "&Terminate"
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "&Options"
      End
   End
   Begin VB.Menu mnuProcess 
      Caption         =   "Process"
      Visible         =   0   'False
      Begin VB.Menu mnuTerminateProcess 
         Caption         =   "&Terminate Process"
      End
   End
End
Attribute VB_Name = "frmRightsBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim hJob As Long
Dim hIocp As Long
Dim IsBoxOn As Boolean

Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean
Private Declare Function CreateFile Lib "kernel32.dll" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByRef lpSecurityAttributes As SECURITY_ATTRIBUTES, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function ExpandEnvironmentStrings Lib "kernel32.dll" Alias "ExpandEnvironmentStringsA" (ByVal lpSrc As String, ByVal lpDst As String, ByVal nSize As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Const ICC_USEREX_CLASSES = &H200

Private Type tagInitCommonControlsEx
        lngSize As Long
        lngICC As Long
End Type

Dim pAdminSid As Long, pLogonSid As Long, pSystemSid As Long, pEveryoneSid As Long, pUsersSid As Long, pRestrictedSid As Long, strLogonSid As String

Private Function RestrictedToken(ByVal hToken As Long, ByRef hNewToken As Long, ByVal bWriteProtected As Boolean) As Long
    
    Dim fStatus As Long
    fStatus = ERROR_SUCCESS
    
    Dim SidsToDelete() As SID_AND_ATTRIBUTES, PrivilegesToDelete As LUID_AND_ATTRIBUTES, SidsToRestrict() As SID_AND_ATTRIBUTES
    
    Dim cbBuff As Long, ret As Long
    ret = GetTokenInformation(hToken, TokenGroups, 0, 0, cbBuff)   '协商缓冲区大小
    
    If (ret <> 0) Or (cbBuff = 0) Then
        fStatus = Err.LastDllError
        Exit Function
    End If
    
    Dim pTokenGrps As TOKEN_GROUPS
    Dim InfoBuff() As Long
    
    ReDim InfoBuff((cbBuff \ 4) - 1) As Long
    ret = GetTokenInformation(hToken, TokenGroups, InfoBuff(0), cbBuff, cbBuff)
    
    If (ret <> 1) Or (cbBuff <= 0) Then
        fStatus = Err.LastDllError
        Exit Function
    End If
    
    Call CopyMemoryOrig(pTokenGrps, InfoBuff(0), Len(pTokenGrps))
    
    ReDim SidsToDelete(pTokenGrps.GroupCount - 1) As SID_AND_ATTRIBUTES
    
    '将除了 Logon SID, Everyone, Users 之外的组全部标记为 Deny-only
    Dim i As Long, nSids As Long, buff As Long, strSid As String
    nSids = 0
    For i = 0 To pTokenGrps.GroupCount - 1
        If IsValidSid(pTokenGrps.Groups(i).Sid) Then
            If ConvertSidToStringSid(ByVal pTokenGrps.Groups(i).Sid, buff) Then
                strSid = TrimNULL(StrConv(StrFromPtrW(buff, 184), vbUnicode))
                If ((pTokenGrps.Groups(i).Attributes And SE_GROUP_LOGON_ID) <> SE_GROUP_LOGON_ID) And (strSid <> "S-1-1-0") And (strSid <> "S-1-5-32-545") Then
                    SidsToDelete(nSids).Sid = pTokenGrps.Groups(i).Sid
                    nSids = nSids + 1
                ElseIf ((pTokenGrps.Groups(i).Attributes And SE_GROUP_LOGON_ID) = SE_GROUP_LOGON_ID) Then
                    pLogonSid = pTokenGrps.Groups(i).Sid
                    strLogonSid = strSid
                End If
            End If
        End If
    Next
    
    Call ConvertStringSidToSid("S-1-5-32-544", pAdminSid)
    Call ConvertStringSidToSid("S-1-5-18", pSystemSid)
    Call ConvertStringSidToSid("S-1-5-12", pRestrictedSid)
    Call ConvertStringSidToSid("S-1-5-32-545", pUsersSid)
    Call ConvertStringSidToSid("S-1-1-0", pEveryoneSid)
    
    ReDim SidsToRestrict(4) As SID_AND_ATTRIBUTES
    
    '将 Everyone, Users, RESTRICTED, Logon SID 标记为 Restricted
    SidsToRestrict(0).Sid = pEveryoneSid
    SidsToRestrict(1).Sid = pUsersSid
    SidsToRestrict(2).Sid = pRestrictedSid
    SidsToRestrict(3).Sid = pLogonSid
    
    If CreateRestrictedToken(hToken, IIf(bWriteProtected, DISABLE_MAX_PRIVILEGE Or WRITE_RESTRICTED_VISTA, DISABLE_MAX_PRIVILEGE), nSids, SidsToDelete(0), 0, PrivilegesToDelete, 4, SidsToRestrict(0), hNewToken) = 0 Then
        fStatus = Err.LastDllError
        Exit Function
    End If
    
    Dim pAcl As Long, pAcl2 As ACL, pAllowedAce As ACCESS_ALLOWED_ACE, cbAcl As Long
    Dim buf(1 To &H400) As Byte
    pAcl = VarPtr(buf(1))
    
    '计算 ACL 大小
    cbAcl = Len(pAcl2) + Len(pAllowedAce) * 3 + (GetLengthSid(pLogonSid) - 4) + (GetLengthSid(pAdminSid) - 4) + (GetLengthSid(pSystemSid) - 4) + 3
    cbAcl = cbAcl And &HFFFFFFFC
    
    '允许 System, Administrators, Logon SID 完全控制
    If InitializeAcl(pAcl, cbAcl, ACL_REVISION) = 0 Then
        fStatus = Err.LastDllError
        Exit Function
    End If
    
    If AddAccessAllowedAce(pAcl, ACL_REVISION, GENERIC_ALL, pSystemSid) = 0 Then
        fStatus = Err.LastDllError
        Exit Function
    End If
    
    If AddAccessAllowedAce(pAcl, ACL_REVISION, GENERIC_ALL, pAdminSid) = 0 Then
        fStatus = Err.LastDllError
        Exit Function
    End If
    
    If AddAccessAllowedAce(pAcl, ACL_REVISION, GENERIC_ALL, pLogonSid) = 0 Then
        fStatus = Err.LastDllError
        Exit Function
    End If
    
    Dim TokenDacl As TOKEN_DEFAULT_DACL
    TokenDacl.DefaultDacl = pAcl
    
    If SetTokenInformation(hNewToken, TokenDefaultDacl, VarPtr(TokenDacl), Len(TokenDacl)) = 0 Then
        fStatus = Err.LastDllError
        Exit Function
    End If
    
End Function

Private Function SaferUsers(ByVal hSaferLevel As Long, ByRef hToken As Long) As Long
    
    Dim fStatus As Long
    fStatus = ERROR_SUCCESS
    
    Dim hAuthzLevel As Long
    
    If SaferCreateLevel(SAFER_SCOPEID_USER, hSaferLevel, 0, hAuthzLevel, 0) = 0 Then
        fStatus = Err.LastDllError
        Exit Function
    End If
    
    If SaferComputeTokenFromLevel(hAuthzLevel, 0, hToken, 0, 0) = 0 Then
        fStatus = Err.LastDllError
        If hAuthzLevel Then SaferCloseLevel hAuthzLevel
        Exit Function
    End If
    
    If hAuthzLevel Then SaferCloseLevel hAuthzLevel
    
End Function

Private Function RunSafer(ByVal szPath As String, ByVal SaferLevel As Long, ByVal DropRights As Boolean) As Long
    
    Dim fStatus As Long
    fStatus = ERROR_SUCCESS
    
    Dim hSaferLevel As Long
    hSaferLevel = SaferLevel
    
    If Len(szPath) > MAX_PATH Then
        RunSafer = ERROR_INVALID_PARAMETER
        Exit Function
    End If
    
    Dim hToken As Long, hNewToken As Long
    Dim sa As SECURITY_ATTRIBUTES
    
    If ((hSaferLevel = SAFER_LEVELID_NORMALUSER) Or (hSaferLevel = SAFER_LEVELID_CONSTRAINED)) And DropRights Then GoTo SaferUsers
    
    If OpenProcessToken(GetCurrentProcess, TOKEN_ALL_ACCESS, hToken) = 0 Then '取得自身令牌句柄
        fStatus = Err.LastDllError
        frmRBMsg.RaiseMessage RBOX(8), fStatus
        GoTo CleanUp
    End If
    
    If Not DropRights Then GoTo RunDirectly
    
    If RestrictedToken(hToken, hNewToken, IIf(hSaferLevel, True, False)) <> ERROR_SUCCESS Then
        frmRBMsg.RaiseMessage RBOX(8), fStatus
        GoTo CleanUp
    End If
    
    GoTo RunDirectly
    
SaferUsers:

    If SaferUsers(hSaferLevel, hToken) <> ERROR_SUCCESS Then
        frmRBMsg.RaiseMessage RBOX(9), fStatus
        GoTo CleanUp
    End If
    
    GoTo RunDirectly
    
RunDirectly:

    Dim si As STARTUPINFO
    si.cb = Len(si)
    Dim pi As PROCESS_INFORMATION
    
    '设置令牌完整性级别
    If (IsVistaOrLater) Then Call SetProcessLowIL(IIf(hNewToken, hNewToken, hToken))
    
    If CreateProcessAsUser(IIf(hNewToken, hNewToken, hToken), szPath, vbNullString, 0&, 0&, False, CREATE_SUSPENDED Or CREATE_BREAKAWAY_FROM_JOB, vbNullString, Environ("SystemDrive"), si, pi) = 0 Then
        fStatus = Err.LastDllError
        frmRBMsg.RaiseMessage RBOX(2), fStatus
        GoTo CleanUp
    End If
    
    hJob = CreateJobObject(sa, vbNullString)
    
    If hJob = 0 Then
        fStatus = Err.LastDllError
        frmRBMsg.RaiseMessage RBOX(3), fStatus
        GoTo CleanUp
    End If
        
    If AssignProcessToJobObject(hJob, pi.hProcess) = 0 Then
        fStatus = Err.LastDllError
        frmRBMsg.RaiseMessage RBOX(4), fStatus
        GoTo CleanUp
    End If
    
    hIocp = CreateIoCompletionPort(-1, 0, 0, 0)
    
    If hIocp = 0 Then
        fStatus = Err.LastDllError
        frmRBMsg.RaiseMessage RBOX(5), fStatus
        GoTo CleanUp
    Else
        Dim joacp As JOBOBJECT_ASSOCIATE_COMPLETION_PORT
        joacp.CompletionKey = 1
        joacp.CompletionPort = hIocp
    End If
        
    If SetInformationJobObject(hJob, JobObjectAssociateCompletionPortInformation, VarPtr(joacp), Len(joacp)) = 0 Then
        fStatus = Err.LastDllError
        frmRBMsg.RaiseMessage RBOX(6), fStatus
        GoTo CleanUp
    End If
    
    If SetJobLimit(dwJobLimit) = 0 Then
        fStatus = Err.LastDllError
        frmRBMsg.RaiseMessage RBOX(6), fStatus
        GoTo CleanUp
    Else
        ResumeThread pi.hThread
        IsBoxOn = True
    End If
    
CleanUp:

    '清场
    If pi.hProcess Then CloseHandle pi.hProcess
    If pi.hThread Then CloseHandle pi.hThread
    If hToken Then CloseHandle hToken
    If hNewToken Then CloseHandle hNewToken
    If pSystemSid Then FreeSid pSystemSid
    If pAdminSid Then FreeSid pAdminSid
    If pLogonSid Then FreeSid pLogonSid
    If pEveryoneSid Then FreeSid pEveryoneSid
    If pRestrictedSid Then FreeSid pRestrictedSid
    If pUsersSid Then FreeSid pUsersSid
    
    RunSafer = fStatus

End Function

Private Function SetJobLimit(ByVal dwLimit As Long) As Long
    Dim JBUI As JOBOBJECT_BASIC_UI_RESTRICTIONS, JBEI As JOBOBJECT_EXTENDED_LIMIT_INFORMATION
    JBUI.UIRestrictionsClass = dwLimit
    SetJobLimit = SetInformationJobObject(hJob, JobObjectBasicUIRestrictions, VarPtr(JBUI), Len(JBUI))
    JBEI.BasicLimitInformation.LimitFlags = JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE
    If SetInformationJobObject(hJob, JobObjectExtendedLimitInformation, VarPtr(JBEI), Len(JBEI)) = 0 Then
        frmRBMsg.RaiseMessage RBOX(6), Err.LastDllError
    End If
End Function

Private Sub SetProcessLowIL(ByVal hToken As Long)
    Dim hNewToken As Long, szIntegritySid As String, pIntegritySid As Long, TIL As TOKEN_MANDATORY_LABEL, pi As PROCESS_INFORMATION, si As STARTUPINFO, b As Long, fStatus As Long
    If dwIL = 2 Then
        szIntegritySid = "S-1-16-4096"
    Else
        szIntegritySid = "S-1-16-8192"
    End If
    
    b = ConvertStringSidToSid(szIntegritySid, pIntegritySid)
    TIL.Label.Attributes = SE_GROUP_INTEGRITY
    TIL.Label.Sid = pIntegritySid
    
    If SetTokenInformation(hToken, TokenIntegrityLevel, VarPtr(TIL), Len(TIL) + GetLengthSid(pIntegritySid)) = 0 Then
        fStatus = Err.LastDllError
        frmRBMsg.RaiseMessage RBOX(7), Err.LastDllError
    End If
End Sub

Private Sub DisableWow64Redirection()
    Dim bIsWow64 As Boolean, tmp As Long
    Call IsWow64Process(-1, bIsWow64)
    If bIsWow64 Then
        Call Wow64DisableWow64FsRedirection(tmp)
    End If
End Sub

Private Sub RevertWow64Redirection()
    Dim bIsWow64 As Boolean, tmp As Long
    Call IsWow64Process(-1, bIsWow64)
    If bIsWow64 Then
        Call Wow64RevertWow64FsRedirection(tmp)
    End If
End Sub

Private Sub DisableUACElevate()
    Call DisableWow64Redirection
    ShellWait "cmd.exe /c icacls.exe %windir%\system32\consent.exe /deny Users:(X)", vbHide, 1000, True
    Call RevertWow64Redirection
End Sub

Private Sub EnableUACElevate()
    Call DisableWow64Redirection
    Shell "cmd.exe /c icacls.exe %windir%\system32\consent.exe /remove:d Users", vbHide
    Call RevertWow64Redirection
End Sub

Private Sub Form_Initialize()
    InitCommonControlsVB
    
    If (IsVistaOrLater) And (Not IsAdmin) Then
        Call ShellExecute(0, "runas", App.Path & "\" & App.EXEName & ".exe", vbNullString, vbNullString, 1)
        End
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    End
End Sub

Private Sub SetRegRights()
    Dim i As Integer
    Dim ret As Long
    
    For i = 0 To intReg - 1
        If Right$(lstReg(i), 2) = "\*" Then
            ret = SetRegistryIntegrity(Left$(lstReg(i), Len(lstReg(i)) - 2), HIGH_INTEGRITY, NO_WRITE_UP Or INHERITENCE)
            If ret <> ERROR_SUCCESS Then
                frmRBMsg.RaiseMessage RBOX(13), ret
            End If
        Else
            ret = SetRegistryIntegrity(lstReg(i), HIGH_INTEGRITY, NO_WRITE_UP)
            If ret <> ERROR_SUCCESS Then
                frmRBMsg.RaiseMessage RBOX(13), ret
            End If
        End If
    Next
    
    For i = 0 To intRegDeny - 1
        If Right$(lstRegDeny(i), 2) = "\*" Then
            ret = SetRegistryIntegrity(Left$(lstRegDeny(i), Len(lstRegDeny(i)) - 2), HIGH_INTEGRITY, NO_READ_UP Or INHERITENCE)
            If ret <> ERROR_SUCCESS Then
                frmRBMsg.RaiseMessage RBOX(13), ret
            End If
        Else
            ret = SetRegistryIntegrity(lstRegDeny(i), HIGH_INTEGRITY, NO_READ_UP)
            If ret <> ERROR_SUCCESS Then
                frmRBMsg.RaiseMessage RBOX(13), ret
            End If
        End If
    Next
End Sub

Private Sub RevertRegRights()
    Dim i As Integer
    Dim ret As Long
    
    For i = 0 To intReg - 1
        If Right$(lstReg(i), 2) = "\*" Then
            ret = SetRegistryIntegrity(Left$(lstReg(i), Len(lstReg(i)) - 2), MEDIUM_INTEGRITY, NO_WRITE_UP Or INHERITENCE)
            If ret <> ERROR_SUCCESS Then
                frmRBMsg.RaiseMessage RBOX(13), ret
            End If
        Else
            ret = SetRegistryIntegrity(lstReg(i), MEDIUM_INTEGRITY, NO_WRITE_UP)
            If ret <> ERROR_SUCCESS Then
                frmRBMsg.RaiseMessage RBOX(13), ret
            End If
        End If
    Next
    
    For i = 0 To intRegDeny - 1
        If Right$(lstRegDeny(i), 2) = "\*" Then
            ret = SetRegistryIntegrity(Left$(lstRegDeny(i), Len(lstRegDeny(i)) - 2), MEDIUM_INTEGRITY, NO_WRITE_UP Or INHERITENCE)
            If ret <> ERROR_SUCCESS Then
                frmRBMsg.RaiseMessage RBOX(13), ret
            End If
        Else
            ret = SetRegistryIntegrity(lstRegDeny(i), MEDIUM_INTEGRITY, NO_WRITE_UP)
            If ret <> ERROR_SUCCESS Then
                frmRBMsg.RaiseMessage RBOX(13), ret
            End If
        End If
    Next
End Sub

Private Sub SetFolderRights()
    Dim i As Integer
    Dim szPath As String
    Dim ret As Long
    
    For i = 0 To intFiles - 1
        szPath = Replace(EnvStr(lstFiles(i)), Chr$(0), "")
        
        If Right$(szPath, 2) = "\*" Then
            ret = SetFileIntegrity(Left$(szPath, Len(szPath) - 2), HIGH_INTEGRITY, NO_WRITE_UP Or INHERITENCE)
            If ret <> ERROR_SUCCESS Then
                frmRBMsg.RaiseMessage RBOX(12), ret
            End If
        Else
            ret = SetFileIntegrity(szPath, HIGH_INTEGRITY, NO_WRITE_UP)
            If ret <> ERROR_SUCCESS Then
                frmRBMsg.RaiseMessage RBOX(12), ret
            End If
        End If
    Next
    
    For i = 0 To intFilesDeny - 1
        szPath = Replace(EnvStr(lstFilesDeny(i)), Chr$(0), "")
        
        If Right$(szPath, 2) = "\*" Then
            ret = SetFileIntegrity(Left$(szPath, Len(szPath) - 2), HIGH_INTEGRITY, NO_READ_UP Or INHERITENCE)
            If ret <> ERROR_SUCCESS Then
                frmRBMsg.RaiseMessage RBOX(12), ret
            End If
        Else
            ret = SetFileIntegrity(szPath, HIGH_INTEGRITY, NO_READ_UP)
            If ret <> ERROR_SUCCESS Then
                frmRBMsg.RaiseMessage RBOX(12), ret
            End If
        End If
    Next
End Sub

Private Sub RevertFolderRights()
    Dim i As Integer
    Dim szPath As String
    Dim ret As Long
    
    For i = 0 To intFiles - 1
        szPath = Replace(EnvStr(lstFiles(i)), Chr$(0), "")
        
        If Right$(szPath, 2) = "\*" Then
            ret = SetFileIntegrity(Left$(szPath, Len(szPath) - 2), MEDIUM_INTEGRITY, NO_WRITE_UP Or INHERITENCE)
            If ret <> ERROR_SUCCESS Then
                frmRBMsg.RaiseMessage RBOX(12), ret
            End If
        Else
            ret = SetFileIntegrity(szPath, MEDIUM_INTEGRITY, NO_WRITE_UP)
            If ret <> ERROR_SUCCESS Then
                frmRBMsg.RaiseMessage RBOX(12), ret
            End If
        End If
    Next
    
    For i = 0 To intFilesDeny - 1
        szPath = Replace(EnvStr(lstFilesDeny(i)), Chr$(0), "")
        
        If Right$(szPath, 2) = "\*" Then
            ret = SetFileIntegrity(Left$(szPath, Len(szPath) - 2), MEDIUM_INTEGRITY, NO_WRITE_UP Or INHERITENCE)
            If ret <> ERROR_SUCCESS Then
                frmRBMsg.RaiseMessage RBOX(12), ret
            End If
        Else
            ret = SetFileIntegrity(szPath, MEDIUM_INTEGRITY, NO_WRITE_UP)
            If ret <> ERROR_SUCCESS Then
                frmRBMsg.RaiseMessage RBOX(12), ret
            End If
        End If
    Next
End Sub

Function EnvStr(ByVal szPath As String) As String
    Dim szBuf As String * 260
    Call ExpandEnvironmentStrings(szPath, szBuf, 260)
    EnvStr = szBuf
End Function

Private Sub lvRestricted_ItemClick(ByVal Item As ComctlLib.ListItem)
    mnuTerminateProcess.Enabled = True
End Sub

Private Sub lvRestricted_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then mnuTerminateProcess.Enabled = False
End Sub

Private Sub lvRestricted_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 And mnuTerminateProcess.Enabled Then PopupMenu mnuProcess
End Sub

Private Sub mnuTerminateProcess_Click()
    If MsgBox("Are you sure you want to terminate " & lvRestricted.SelectedItem.Text & "?" & vbCrLf & vbCrLf & "Terminating running process will cause it to lose unsaved data.", vbYesNo) = vbYes Then
        Dim hProcess As Long
        
        hProcess = OpenProcess(PROCESS_TERMINATE, 0, lvRestricted.SelectedItem.SubItems(1))
        
        If (hProcess) Then
            Call TerminateProcess(hProcess, 0)
            Call CloseHandle(hProcess)
        End If
    End If
End Sub

Private Sub mnuUntrusted_Click() '不信任
    dwJobLimit = JOB_OBJECT_UILIMIT_HANDLES Or JOB_OBJECT_UILIMIT_EXITWINDOWS Or JOB_OBJECT_UILIMIT_DESKTOP Or JOB_OBJECT_UILIMIT_DISPLAYSETTINGS Or JOB_OBJECT_UILIMIT_READCLIPBOARD Or JOB_OBJECT_UILIMIT_WRITECLIPBOARD Or JOB_OBJECT_UILIMIT_SYSTEMPARAMETERS Or JOB_OBJECT_UILIMIT_GLOBALATOMS
    dwSaferLevel = 0
    bDropRights = True
    DenyUACAdmin = False
    dwIL = 2
    dwLimitMode = 4
    Call mnuRun_Click
End Sub

Private Sub mnuLimited_Click() '限制性
    dwJobLimit = JOB_OBJECT_UILIMIT_HANDLES Or JOB_OBJECT_UILIMIT_EXITWINDOWS Or JOB_OBJECT_UILIMIT_DESKTOP Or JOB_OBJECT_UILIMIT_DISPLAYSETTINGS Or JOB_OBJECT_UILIMIT_WRITECLIPBOARD Or JOB_OBJECT_UILIMIT_SYSTEMPARAMETERS
    dwSaferLevel = SAFER_LEVELID_CONSTRAINED
    bDropRights = True
    DenyUACAdmin = False
    dwIL = 1
    dwLimitMode = 3
    Call mnuRun_Click
End Sub

Private Sub mnuLowRights_Click() '低权限
    dwJobLimit = JOB_OBJECT_UILIMIT_HANDLES Or JOB_OBJECT_UILIMIT_EXITWINDOWS Or JOB_OBJECT_UILIMIT_DESKTOP Or JOB_OBJECT_UILIMIT_DISPLAYSETTINGS Or JOB_OBJECT_UILIMIT_WRITECLIPBOARD Or JOB_OBJECT_UILIMIT_SYSTEMPARAMETERS
    dwSaferLevel = 1
    bDropRights = True
    DenyUACAdmin = False
    dwIL = 1
    dwLimitMode = 2
    Call mnuRun_Click
End Sub

Private Sub mnuBasicUser_Click() '基本用户
    dwJobLimit = JOB_OBJECT_UILIMIT_HANDLES Or JOB_OBJECT_UILIMIT_EXITWINDOWS Or JOB_OBJECT_UILIMIT_DESKTOP Or JOB_OBJECT_UILIMIT_SYSTEMPARAMETERS
    dwSaferLevel = SAFER_LEVELID_NORMALUSER
    bDropRights = True
    DenyUACAdmin = False
    dwIL = 1
    dwLimitMode = 1
    Call mnuRun_Click
End Sub

Private Sub mnuRun_Click()
    Dim fStatus As Long
    
    If (IsVistaOrLater) Then Call SetFolderRights
    If (IsVistaOrLater) Then Call SetRegRights
    If (IsVistaOrLater) And (DenyUACAdmin) Then Call DisableUACElevate
    
    fStatus = RunSafer(App.Path & "\RBStart.exe", dwSaferLevel, bDropRights)
    
    If fStatus = ERROR_SUCCESS Then
        Call SetWinExceptions
    Else
        frmRBMsg.RaiseMessage RBOX(1), fStatus
        If (IsVistaOrLater) Then Call RevertFolderRights
        If (IsVistaOrLater) Then Call RevertRegRights
        If (IsVistaOrLater) And (DenyUACAdmin) Then Call EnableUACElevate
    End If
End Sub

Private Sub CloseBox()
    TerminateJobObject hJob, 0
    CloseHandle hJob
    CloseHandle hIocp
    If (IsVistaOrLater) Then Call RevertFolderRights
    If (IsVistaOrLater) Then Call RevertRegRights
    If (IsVistaOrLater) And (DenyUACAdmin) Then Call EnableUACElevate
    lvRestricted.ListItems.Clear
    IsBoxOn = False
End Sub

Private Sub SetWinExceptions()
    Dim hWnd As Long
    hWnd = FindWindow("TrayNotifyWnd", vbNullString)
    Call UserHandleGrantAccess(hWnd, hJob, 1)
    hWnd = FindWindow("SystemTray_Main", vbNullString)
    Call UserHandleGrantAccess(hWnd, hJob, 1)
    hWnd = FindWindow("Connections Tray", vbNullString)
    Call UserHandleGrantAccess(hWnd, hJob, 1)
    hWnd = LoadIcon(0, IDI_APPLICATION)
    Call UserHandleGrantAccess(hWnd, hJob, 1)
    hWnd = LoadIcon(0, IDI_INFORMATION)
    Call UserHandleGrantAccess(hWnd, hJob, 1)
    hWnd = LoadIcon(0, IDI_EXCLAMATION)
    Call UserHandleGrantAccess(hWnd, hJob, 1)
    hWnd = LoadIcon(0, IDI_ERROR)
    Call UserHandleGrantAccess(hWnd, hJob, 1)
    hWnd = LoadIcon(0, IDI_QUESTION)
    Call UserHandleGrantAccess(hWnd, hJob, 1)
    hWnd = LoadIcon(0, IDI_SHIELD)
    Call UserHandleGrantAccess(hWnd, hJob, 1)
    hWnd = LoadIcon(0, IDI_WINLOGO)
    Call UserHandleGrantAccess(hWnd, hJob, 1)
    hWnd = LoadCursor(0, IDC_ARROW)
    Call UserHandleGrantAccess(hWnd, hJob, 1)
    hWnd = LoadCursor(0, IDC_HAND)
    Call UserHandleGrantAccess(hWnd, hJob, 1)
    hWnd = LoadCursor(0, IDC_IBEAM)
    Call UserHandleGrantAccess(hWnd, hJob, 1)
End Sub

Private Sub cmdRightsBox_Click()
    lvRestricted.SetFocus
    Me.PopupMenu mnuSandbox, , cmdRightsBox.Left, cmdRightsBox.Top + cmdRightsBox.Height
End Sub

Private Sub Form_Load()
    If Not IsVistaOrLater Then MsgBox "OS not recommended!! (Windows Vista at least)", vbExclamation, "WARNING"
    
    dwSaferLevel = SAFER_LEVELID_NORMALUSER
    dwIL = 1
    dwJobLimit = JOB_OBJECT_UILIMIT_HANDLES Or JOB_OBJECT_UILIMIT_EXITWINDOWS Or JOB_OBJECT_UILIMIT_DESKTOP Or JOB_OBJECT_UILIMIT_SYSTEMPARAMETERS
    dwLimitMode = 1
    bDropRights = True
    SetWinTextR = True
    Me.Icon = LoadPicture("")
    SendMessageLong lvRestricted.hWnd, LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_FULLROWSELECT, True
    'Call MiniEnablePrivilege
    
    Call AddRule(RULE_TYPE_FILE_SYSTEM, RULE_ACTION_READONLY, "%AppData%\Microsoft\Windows\Start Menu\Programs\Startup")
    Call AddRule(RULE_TYPE_REGISTRY, RULE_ACTION_READONLY, "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run")
    Call AddRule(RULE_TYPE_REGISTRY, RULE_ACTION_READONLY, "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\RunOnce")
    Call AddRule(RULE_TYPE_REGISTRY, RULE_ACTION_READONLY, "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved")
    Call AddRule(RULE_TYPE_REGISTRY, RULE_ACTION_READONLY, "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\System")
    Call AddRule(RULE_TYPE_REGISTRY, RULE_ACTION_READONLY, "HKEY_CURRENT_USER\Software\Microsoft\Windows NT\CurrentVersion\Windows")
    Call AddRule(RULE_TYPE_REGISTRY, RULE_ACTION_READONLY, "HKEY_CURRENT_USER\Software\Microsoft\Windows NT\CurrentVersion\Winlogon")
    Call AddRule(RULE_TYPE_REGISTRY, RULE_ACTION_READONLY, "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main")
    Call AddRule(RULE_TYPE_WINDOW, RULE_ACTION_ALLOW, "Shell_TrayWnd")
End Sub

Private Sub mnuOptions_Click()
    frmRBOptions.Show vbModal
End Sub

Private Sub mnuTerminate_Click()
    Call CloseBox
End Sub

Private Sub Timer1_Timer()
    If IsBoxOn = False Then
        mnuRun.Enabled = True
        mnuRunAs.Enabled = True
        mnuTerminate.Enabled = False
        cmdRightsBox.BackColor = &HFF&
    Else
        mnuRun.Enabled = False
        mnuRunAs.Enabled = False
        mnuTerminate.Enabled = True
        cmdRightsBox.BackColor = &HFF00&
    End If
End Sub

Private Sub Timer2_Timer()
    Dim lpOverlapped As Long, dwEvent As Long, dwCompKey As Long
    
    GetQueuedCompletionStatus hIocp, dwEvent, dwCompKey, lpOverlapped, 100
    
    Select Case dwEvent
        Case JOB_OBJECT_MSG_ACTIVE_PROCESS_ZERO
            Call CloseBox
            Exit Sub
        Case JOB_OBJECT_MSG_NEW_PROCESS
            Dim szPath As String, hProcess As Long
            szPath = GetProcessImagePath(lpOverlapped)
            Dim SplitStr() As String
            If szPath <> "" Then
                SplitStr = Split(szPath, "\")
                szPath = SplitStr(UBound(SplitStr))
                lvRestricted.ListItems.Add , , szPath
                lvRestricted.ListItems(lvRestricted.ListItems.Count).SubItems(1) = lpOverlapped
                Exit Sub
            End If
        Case JOB_OBJECT_MSG_EXIT_PROCESS
            Dim i As Long
            For i = 1 To lvRestricted.ListItems.Count
                If lpOverlapped = lvRestricted.ListItems.Item(i).SubItems(1) Then
                    lvRestricted.ListItems.Remove i
                    Exit Sub
                End If
            Next
        Case JOB_OBJECT_MSG_ABNORMAL_EXIT_PROCESS
            For i = 1 To lvRestricted.ListItems.Count
                If lpOverlapped = lvRestricted.ListItems.Item(i).SubItems(1) Then
                    lvRestricted.ListItems.Remove i
                    Exit Sub
                End If
            Next
    End Select
    
End Sub

Private Sub Timer3_Timer()
    Dim hWnd As New Collection, i As Long, j As Long
    If SetWinTextR Then
        If IsBoxOn Then
            For i = 1 To lvRestricted.ListItems.Count
                HxGetHWndByPID lvRestricted.ListItems.Item(i).SubItems(1), hWnd, Me.hWnd
                For j = 1 To hWnd.Count
                    Dim cbWinTextLen As Long, szWinText As String
                    cbWinTextLen = GetWindowTextLength(hWnd.Item(j)) + 1
                    szWinText = Space$(cbWinTextLen)
                    GetWindowText hWnd.Item(j), szWinText, cbWinTextLen
                    lvRestricted.ListItems.Item(i).SubItems(2) = Trim(szWinText)
                    If Left(szWinText, 3) <> "[R]" Then SetWindowText hWnd.Item(j), "[R] " & szWinText
                Next
            Next
        End If
    End If
End Sub

Private Sub Timer4_Timer()
    Dim i As Long
    If IsBoxOn Then
        For i = 1 To lvRestricted.ListItems.Count
            If lvRestricted.ListItems.Item(i).Text = "Unknown" Then
                Dim szPath As String, hProcess As Long
                szPath = GetProcessImagePath(lvRestricted.ListItems.Item(i).SubItems(1))
                Dim SplitStr() As String
                SplitStr = Split(szPath, "\")
                szPath = SplitStr(UBound(SplitStr))
                lvRestricted.ListItems.Item(i).Text = szPath
            End If
        Next
    End If
End Sub

Private Sub Timer5_Timer()
    Dim i As Long, hWnd As Long
    If IsBoxOn Then
        For i = 0 To intOpenWin - 1
            hWnd = FindWindow(lstOpenWin(i), vbNullString)
            
            If hWnd Then
                Call UserHandleGrantAccess(hWnd, hJob, 1)
            End If
        Next
    End If
End Sub

Public Function TrimNULL(ByVal str As String) As String
    If InStr(str, Chr$(0)) > 0& Then
        TrimNULL = Left$(str, InStr(str, Chr$(0)) - 1&)
    Else
        TrimNULL = str
    End If
End Function

Public Function StrFromPtrW(ByVal lpszW As Long, Optional nSize As Long = 0) As String
   Dim s As String, bTrim As Boolean
   If nSize = 0 Then
      nSize = lstrlenW(lpszW)
      bTrim = True
   End If
   s = String(nSize, Chr$(0))
   CopyMemoryOrig ByVal StrPtr(s), ByVal lpszW, nSize
   If bTrim Then s = TrimNULL(s)
   StrFromPtrW = s
End Function

Function InitCommonControlsVB() As Boolean
        On Error Resume Next
        Dim iccex As tagInitCommonControlsEx
        With iccex
           .lngSize = LenB(iccex)
           .lngICC = ICC_USEREX_CLASSES
        End With
        InitCommonControlsEx iccex
        InitCommonControlsVB = (Err.Number = 0)
        On Error GoTo 0
End Function
