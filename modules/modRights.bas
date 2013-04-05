Attribute VB_Name = "modRights"
Option Explicit

Declare Function SaferCreateLevel Lib "ADVAPI32.dll" (ByVal dwScopeId As Long, ByVal dwLevelId As Long, ByVal OpenFlags As Long, ByRef pLevelHandle As Long, ByVal lpReserved As Long) As Long
Declare Function SaferComputeTokenFromLevel Lib "ADVAPI32.dll" (ByVal LevelHandle As Long, ByVal InAccessToken As Long, ByRef OutAccessToken As Long, ByVal dwFlags As Long, ByVal lpReserved As Long) As Long
Declare Function SaferCloseLevel Lib "ADVAPI32.dll" (ByVal hLevelHandle As Long) As Long

Declare Function CreateJobObject Lib "kernel32.dll" Alias "CreateJobObjectA" (lpJobAttributes As SECURITY_ATTRIBUTES, lpName As String) As Long
Declare Function AssignProcessToJobObject Lib "kernel32.dll" (ByVal hJob As Long, ByVal hProcess As Long) As Long
Declare Function SetInformationJobObject Lib "kernel32.dll" (ByVal hJob As Long, ByVal JobObjectInformationClass As Long, ByVal lpJobObjectInformation As Long, ByVal cbJobObjectInformationLength As Long) As Long
Declare Function TerminateJobObject Lib "kernel32.dll" (ByVal hJob As Long, ByVal uExitCode As Long) As Long
Declare Function UserHandleGrantAccess Lib "user32.dll" (ByVal hUserHandle As Long, ByVal hJob As Long, ByVal bGrant As Long) As Long

Declare Function CreateIoCompletionPort Lib "kernel32.dll" (ByVal FileHandle As Long, ByVal ExistingCompletionPort As Long, ByVal CompletionKey As Long, ByVal NumberOfConcurrentThreads As Long) As Long
Declare Function GetQueuedCompletionStatus Lib "kernel32.dll" (ByVal CompletionPort As Long, ByRef lpNumberOfBytesTransferred As Long, ByRef lpCompletionKey As Long, ByRef lpOverlapped As Long, ByVal dwMilliseconds As Long) As Long

Declare Function CreateRestrictedToken Lib "ADVAPI32.dll" (ByVal ExistingTokenHandle As Long, ByVal Flags As Long, ByVal DisableSidCount As Long, ByRef SidsToDisable As SID_AND_ATTRIBUTES, ByVal DeletePrivilegeCount As Long, ByRef PrivilegesToDelete As LUID_AND_ATTRIBUTES, ByVal RestrictedSidCount As Long, ByRef SidsToRestrict As SID_AND_ATTRIBUTES, ByRef NewTokenHandle As Long) As Long
Declare Function OpenProcessToken Lib "ADVAPI32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, ByRef TokenHandle As Long) As Long
Declare Function GetTokenInformation Lib "ADVAPI32.dll" (ByVal TokenHandle As Long, ByVal TokenInformationClass As Integer, ByRef TokenInformation As Any, ByVal TokenInformationLength As Long, ByRef ReturnLength As Long) As Long
Declare Function SetTokenInformation Lib "ADVAPI32.dll" (ByVal TokenHandle As Long, ByVal TokenInformationClass As Integer, ByVal TokenInformation As Any, ByVal TokenInformationLength As Long) As Long
Declare Function CheckTokenMembership Lib "ADVAPI32.dll" (ByVal TokenHandle As Long, ByVal SidToCheck As Long, ByRef IsMember As Long) As Long

Declare Function ConvertSidToStringSid Lib "ADVAPI32.dll" Alias "ConvertSidToStringSidA" (ByVal pSid As Long, ByRef StringSid As Long) As Long
Declare Function ConvertStringSidToSid Lib "ADVAPI32.dll" Alias "ConvertStringSidToSidA" (ByVal StringSid As String, ByRef pSid As Long) As Boolean
Declare Function GetLengthSid Lib "ADVAPI32.dll" (ByVal pSid As Long) As Long
Declare Function IsValidSid Lib "ADVAPI32.dll" (ByVal pSid As Long) As Long
Declare Function InitializeAcl Lib "ADVAPI32.dll" (ByVal pAcl As Long, ByVal nAclLength As Long, ByVal dwAclRevision As Long) As Long
'Declare Function IsValidAcl Lib "advapi32.dll" (ByVal pAcl As Long) As Long
Declare Function AddAccessAllowedAce Lib "ADVAPI32.dll" (ByVal pAcl As Long, ByVal dwAceRevision As Long, ByVal AccessMask As Long, ByVal pSid As Long) As Long
Declare Sub FreeSid Lib "ADVAPI32.dll" (ByVal pSid As Long)

Declare Function ConvertStringSecurityDescriptorToSecurityDescriptorW Lib "ADVAPI32.dll" (ByVal lpStringSecurityDescriptor As Long, ByVal StringSDRevision As Long, ByRef lpSecurityDescriptor As Long, ByRef lpSecurityDescriptorSize As Long) As Long
Declare Function GetSecurityDescriptorSacl Lib "ADVAPI32.dll" (ByRef pSecurityDescriptor As Long, ByRef lpbSaclPresent As Long, ByRef pSacl As Long, ByRef lpbSaclDefaulted As Long) As Long
Declare Function SetSecurityInfo Lib "ADVAPI32.dll" (ByVal handle As Long, ByVal ObjectType As SE_OBJECT_TYPE, ByVal SecurityInfo As Long, ByRef psidOwner As Long, ByRef psidGroup As Long, ByVal pDacl As Long, ByVal pSacl As Long) As Long
Declare Function SetNamedSecurityInfoA Lib "ADVAPI32.dll" (ByVal pObjectName As String, ByVal ObjectType As SE_OBJECT_TYPE, ByVal SecurityInfo As Long, ByRef psidOwner As Long, ByRef psidGroup As Long, ByVal pDacl As Long, ByVal pSacl As Long) As Long
Declare Function LocalFree Lib "kernel32.dll" (ByVal hMem As Long) As Long

Declare Function RegCreateKeyEx Lib "ADVAPI32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByRef lpSecurityAttributes As SECURITY_ATTRIBUTES, ByRef phkResult As Long, ByRef lpdwDisposition As Long) As Long

Declare Function Wow64DisableWow64FsRedirection Lib "kernel32.dll" (ByVal OldValue As Long) As Boolean
Declare Function Wow64RevertWow64FsRedirection Lib "kernel32.dll" (ByVal OldValue As Long) As Boolean
Declare Function IsWow64Process Lib "kernel32.dll" (ByVal hProcess As Long, IsEnable As Boolean) As Boolean

Declare Function LoadIcon Lib "user32.dll" Alias "LoadIconA" (ByVal hInstance As Long, ByVal lpIconName As Long) As Long
Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long

Declare Function LoadCursor Lib "user32.dll" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Declare Function DestroyCursor Lib "user32.dll" (ByVal hCursor As Long) As Long

Declare Function GetCurrentProcess Lib "kernel32.dll" () As Long

Declare Function CreateProcessAsUser Lib "ADVAPI32.dll" Alias "CreateProcessAsUserA" (ByVal hToken As Long, ByVal lpApplicationName As String, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Boolean, ByVal dwCreationFlags As Long, ByVal lpEnvironment As String, ByVal lpCurrentDirectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Declare Function ResumeThread Lib "kernel32.dll" (ByVal hThread As Long) As Long
Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Declare Function TerminateProcess Lib "kernel32.dll" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long

Declare Sub CopyMemory Lib "ntdll.dll" Alias "RtlMoveMemory" (ByVal Dest As Long, ByVal Source As Long, ByVal Length As Long)
Declare Sub CopyMemoryOrig Lib "kernel32.dll" Alias "RtlMoveMemory" (Dest As Any, Source As Any, ByVal Length As Long)

Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long

Declare Function ZwQueryInformationProcess Lib "ntdll.dll" _
                 (ByVal ProcessHandle As Long, _
                 ByVal ProcessInformationClass As Long, _
                 ByVal ProcessInformation As Long, _
                 ByVal ProcessInformationLength As Long, _
                 ByRef ReturnLength As Long) As Long

'安全等级 ID  Safer Level ID
Public Const SAFER_LEVELID_DISALLOWED = 0
Public Const SAFER_LEVELID_UNTRUSTED = &H1000
Public Const SAFER_LEVELID_CONSTRAINED = &H10000
Public Const SAFER_LEVELID_NORMALUSER = &H20000
Public Const SAFER_LEVELID_FULLYTRUSTED = &H40000

'作用域 ID    Scope ID
Public Const SAFER_SCOPEID_MACHINE = 1
Public Const SAFER_SCOPEID_USER = 2

Public Const MAX_PATH As Integer = 260

'错误  Errors
Public Const ERROR_SUCCESS = 0&
Public Const INVALID_HANDLE_VALUE As Long = -1
Public Const ERROR_INVALID_PARAMETER = 87 '  dderror
Public Const STATUS_INFO_LENGTH_MISMATCH As Long = &HC0000004

'创建标志  Creation Flags
Public Const CREATE_BREAKAWAY_FROM_JOB = &H1000000
Public Const CREATE_DEFAULT_ERROR_MODE = &H4000000
Public Const CREATE_NEW_CONSOLE = &H10
Public Const CREATE_NEW_PROCESS_GROUP = &H200
Public Const CREATE_NO_WINDOW = &H8000000
Public Const CREATE_PROTECTED_PROCESS = &H40000
Public Const CREATE_PRESERVE_CODE_AUTHZ_LEVEL = &H2000000
Public Const CREATE_SEPARATE_WOW_VDM = &H800
Public Const CREATE_SHARED_WOW_VDM = &H1000
Public Const CREATE_SUSPENDED = &H4
Public Const CREATE_UNICODE_ENVIRONMENT = &H400
Public Const DEBUG_ONLY_THIS_PROCESS = &H2
Public Const DEBUG_PROCESS = &H1
Public Const DETACHED_PROCESS = &H8
Public Const EXTENDED_STARTUPINFO_PRESENT = &H80000
Public Const INHERIT_PARENT_AFFINITY = &H10000

'IL
Public Const SE_GROUP_INTEGRITY = &H20

'标准图标  Standard Icon
Public Const IDI_APPLICATION = 32512&
Public Const IDI_INFORMATION = 32516&
Public Const IDI_EXCLAMATION = 32515&
Public Const IDI_ERROR = 32513&
Public Const IDI_QUESTION = 32514&
Public Const IDI_SHIELD = 32518&
Public Const IDI_WINLOGO = 32517&

'标准光标  Standard Cursor
Public Const IDC_ARROW As Long = 32512&
Public Const IDC_HAND As Long = 32649&
Public Const IDC_IBEAM As Long = 32513&

Public Const SE_GROUP_LOGON_ID As Long = &HC0000000
 
Public Const ANYSIZE_ARRAY As Long = 20

Public Const DISABLE_MAX_PRIVILEGE As Long = &H1
Public Const WRITE_RESTRICTED_VISTA As Long = &H8

'访问控制列表 ACL
Public Const ACL_REVISION As Long = 2

'访问权限 Access Masks
Public Const GENERIC_READ As Long = &H80000000
Public Const GENERIC_WRITE As Long = &H40000000
Public Const GENERIC_ALL As Long = &H10000000
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const SYNCHRONIZE As Long = &H100000

Public Const FILE_SHARE_READ As Long = &H1
Public Const FILE_SHARE_WRITE As Long = &H2

Public Const OPEN_EXISTING As Long = 3
Public Const FILE_FLAG_BACKUP_SEMANTICS As Long = &H2000000

Public Const PROCESS_TERMINATE As Long = (&H1)
Public Const PROCESS_QUERY_INFORMATION As Long = (&H400)
Public Const PROCESS_ALL_ACCESS As Long = (STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &HFFF)

Public Const TOKEN_ASSIGN_PRIMARY = &H1
Public Const TOKEN_DUPLICATE = (&H2)
Public Const TOKEN_IMPERSONATE = (&H4)
Public Const TOKEN_QUERY = (&H8)
Public Const TOKEN_QUERY_SOURCE = (&H10)
Public Const TOKEN_ADJUST_PRIVILEGES = (&H20)
Public Const TOKEN_ADJUST_GROUPS = (&H40)
Public Const TOKEN_ADJUST_DEFAULT = (&H80)
Public Const TOKEN_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or TOKEN_ASSIGN_PRIMARY Or _
 TOKEN_DUPLICATE Or TOKEN_IMPERSONATE Or TOKEN_QUERY Or TOKEN_QUERY_SOURCE Or _
 TOKEN_ADJUST_PRIVILEGES Or TOKEN_ADJUST_GROUPS Or TOKEN_ADJUST_DEFAULT)

Public Const LOW_INTEGRITY_SDDL_SACL_W As String = "S:(ML;;NW;;;LW)"
Public Const MEDIUM_INTEGRITY_SDDL_SACL_W As String = "S:(ML;;NW;;;ME)"
Public Const HIGH_INTEGRITY_SDDL_SACL_W As String = "S:(ML;;NW;;;HI)"

Public Const SDDL_REVISION_1 As Long = 1

Public Const LABEL_SECURITY_INFORMATION As Long = &H10

Enum JOBOBJECTINFOCLASS
    JobObjectBasicAccountingInformation = 1
    JobObjectBasicLimitInformation = 2
    JobObjectBasicProcessIdList = 3
    JobObjectBasicUIRestrictions = 4
    JobObjectSecurityLimitInformation = 5
    JobObjectEndOfJobTimeInformation = 6
    JobObjectAssociateCompletionPortInformation = 7
    JobObjectBasicAndIoAccountingInformation = 8
    JobObjectExtendedLimitInformation = 9
    MaxJobObjectInfoClass = 10
End Enum

Enum JOBOBJECTMSGSCLASS
    JOB_OBJECT_MSG_END_OF_JOB_TIME = 1
    JOB_OBJECT_MSG_END_OF_PROCESS_TIME = 2
    JOB_OBJECT_MSG_ACTIVE_PROCESS_LIMIT = 3
    JOB_OBJECT_MSG_ACTIVE_PROCESS_ZERO = 4
    JOB_OBJECT_MSG_NEW_PROCESS = 6
    JOB_OBJECT_MSG_EXIT_PROCESS = 7
    JOB_OBJECT_MSG_ABNORMAL_EXIT_PROCESS = 8
    JOB_OBJECT_MSG_PROCESS_MEMORY_LIMIT = 9
    JOB_OBJECT_MSG_JOB_MEMORY_LIMIT = 10
End Enum

Enum JOBOBJECT_UI_LIMIT
    JOB_OBJECT_UILIMIT_DESKTOP = &H40
    JOB_OBJECT_UILIMIT_DISPLAYSETTINGS = &H10
    JOB_OBJECT_UILIMIT_EXITWINDOWS = &H80
    JOB_OBJECT_UILIMIT_GLOBALATOMS = &H20
    JOB_OBJECT_UILIMIT_HANDLES = &H1
    JOB_OBJECT_UILIMIT_READCLIPBOARD = &H2
    JOB_OBJECT_UILIMIT_SYSTEMPARAMETERS = &H8
    JOB_OBJECT_UILIMIT_WRITECLIPBOARD = &H4
End Enum

Enum JobObjectLimitFlags
    JOB_OBJECT_LIMIT_ACTIVE_PROCESS = &H8
    JOB_OBJECT_LIMIT_AFFINITY = &H10
    JOB_OBJECT_LIMIT_BREAKAWAY_OK = &H800
    JOB_OBJECT_LIMIT_DIE_ON_UNHANDLED_EXCEPTION = &H400
    JOB_OBJECT_LIMIT_JOB_MEMORY = &H200
    JOB_OBJECT_LIMIT_JOB_TIME = &H4
    JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE = &H2000
    JOB_OBJECT_LIMIT_PRESERVE_JOB_TIME = &H40
    JOB_OBJECT_LIMIT_PRIORITY_CLASS = &H20
    JOB_OBJECT_LIMIT_PROCESS_MEMORY = &H100
    JOB_OBJECT_LIMIT_PROCESS_TIME = &H2
    JOB_OBJECT_LIMIT_SCHEDULING_CLASS = &H80
    JOB_OBJECT_LIMIT_SILENT_BREAKAWAY_OK = &H1000
    JOB_OBJECT_LIMIT_WORKINGSET = &H1
End Enum

'Token枚举类  Token Enum
Enum TOKEN_INFORMATION_CLASS
    TokenUser = 1
    TokenGroups
    TokenPrivileges
    TokenOwner
    TokenPrimaryGroup
    TokenDefaultDacl
    TokenSource
    TokenType
    TokenImpersonationLevel
    TokenStatistics
    TokenRestrictedSids
    TokenSessionId
    TokenGroupsAndPrivileges
    TokenSessionReference
    TokenSandBoxInert
    TokenAuditPolicy
    TokenOrigin
    TokenElevationType
    TokenLinkedToken
    TokenElevation
    TokenHasRestrictions
    TokenAccessInformation
    TokenVirtualizationAllowed
    TokenVirtualizationEnabled
    TokenIntegrityLevel
    TokenUIAccess
    TokenMandatoryPolicy
    TokenLogonSid
    MaxTokenInfoClass
End Enum

Enum SE_OBJECT_TYPE
    SE_UNKNOWN_OBJECT_TYPE = 0
    SE_FILE_OBJECT
    SE_SERVICE
    SE_PRINTER
    SE_REGISTRY_KEY
    SE_LMSHARE
    SE_KERNEL_OBJECT
    SE_WINDOW_OBJECT
    SE_DS_OBJECT
    SE_DS_OBJECT_ALL
    SE_PROVIDER_DEFINED_OBJECT
    SE_WMIGUID_OBJECT
    SE_REGISTRY_WOW64_32KEY
End Enum

Type LARGE_INTEGER
    LowPart As Long
    HighPart As Long
End Type

Type IO_COUNTERS
    ReadOperationCount As LARGE_INTEGER
    WriteOperationCount As LARGE_INTEGER
    OtherOperationCount As LARGE_INTEGER
    ReadTransferCount As LARGE_INTEGER
    WriteTransferCount As LARGE_INTEGER
    OtherTransferCount As LARGE_INTEGER
End Type

Type JOBOBJECT_BASIC_LIMIT_INFORMATION
    PerProcessUserTimeLimit As LARGE_INTEGER
    PorJobUserTimeLimit As LARGE_INTEGER
    LimitFlags As Long
    MinimumWorkingSetSize As Long
    MaximumWorkingSetSize As Long
    ActiveProcessLimit As Long
    Affinity As LARGE_INTEGER
    PriorityClass As Long
    SchedulingClass As Long
End Type

Type JOBOBJECT_EXTENDED_LIMIT_INFORMATION
    BasicLimitInformation As JOBOBJECT_BASIC_LIMIT_INFORMATION
    IoInfo As IO_COUNTERS
    ProcessMemoryLimit As Long
    JobMemoryLimit As Long
    PeakProcessMemoryUsed As Long
    PeakJobMemoryUsed As Long
End Type

Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadId As Long
End Type

Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Type LUID
    LowPart As Long
    HighPart As Long
End Type

Type LUID_AND_ATTRIBUTES
    pLuid As LUID
    Attributes As Long
End Type

Type SID_AND_ATTRIBUTES
    Sid As Long
    Attributes As Long
End Type

Type TOKEN_MANDATORY_LABEL
    Label As SID_AND_ATTRIBUTES
End Type

Type TOKEN_GROUPS
    GroupCount As Long
    Groups(ANYSIZE_ARRAY) As SID_AND_ATTRIBUTES
End Type

Type JOBOBJECT_ASSOCIATE_COMPLETION_PORT
    CompletionKey As Long
    CompletionPort As Long
End Type

Type OVERLAPPED
    Internal As Long
    InternalHigh As Long
    offset As Long
    OffsetHigh As Long
    hEvent As Long
End Type

Type JOBOBJECT_BASIC_UI_RESTRICTIONS
    UIRestrictionsClass As Long
End Type

Type ACL
    AclRevision As Byte
    Sbz1 As Byte
    AclSize As Integer
    AceCount As Integer
    Sbz2 As Integer
End Type

Type ACE_HEADER
    AceType As Byte
    AceFlags As Byte
    AceSize As Long
End Type

Type ACCESS_ALLOWED_ACE
    Header As ACE_HEADER
    Mask As Long
    SidStart As Long
End Type

Type TOKEN_DEFAULT_DACL
    DefaultDacl As Long
End Type

Type SECURITY_DESCRIPTOR
   Revision As Byte
   Sbz1 As Byte
   Control As Long
   Owner As Long
   Group As Long
   Sacl As ACL
   Dacl As ACL
End Type

'Custom
Enum SandboxRuleType
    RULE_TYPE_FILE_SYSTEM
    RULE_TYPE_REGISTRY
    RULE_TYPE_WINDOW
End Enum

Enum SandboxRuleAction
    RULE_ACTION_ALLOW
    RULE_ACTION_DENY
    RULE_ACTION_READONLY
End Enum

Enum SE_INTEGRITY_LEVEL
    LOW_INTEGRITY = 1
    MEDIUM_INTEGRITY
    HIGH_INTEGRITY
End Enum

Enum SE_INTEGRITY_POLICY
    NO_READ_UP = 1
    NO_WRITE_UP
    INHERITENCE
End Enum


Public bDropRights As Boolean
Public dwSaferLevel As Long  'Possible Values: SAFER_LEVELID_NORMALUSER, SAFER_LEVELID_CONSTRAINED, 1 - Custom Safer Level with WRITE_RESTRICTED on, 0 - Custom Safer Level
Public SetWinTextR As Boolean
Public DenyUACAdmin As Boolean
Public dwIL As Integer  '1 - Medium IL, 2 - Low IL
Public dwJobLimit As Long
Public dwLimitMode As Integer '1 - BasicUser, 2 - LowRights, 3 - Limited, 4 - Untrusted
Public lstFiles(50) As String  '只读文件列表
Public intFiles As Integer  'lstFiles计数
Public lstFilesDeny(50) As String  '阻止文件列表
Public intFilesDeny As Integer
Public lstReg(50) As String  '只读注册表列表
Public intReg As Integer
Public lstRegDeny(50) As String  '阻止注册表列表
Public intRegDeny As Integer
Public lstOpenWin(50) As String  '允许访问窗口列表
Public intOpenWin As Integer

Sub AddRule(RuleType As SandboxRuleType, RuleAction As SandboxRuleAction, ByVal RuleObject As String)
    Select Case RuleType
        Case RULE_TYPE_FILE_SYSTEM
            Select Case RuleAction
                Case RULE_ACTION_DENY
                    lstFilesDeny(intFilesDeny) = RuleObject
                    intFilesDeny = intFilesDeny + 1
                Case RULE_ACTION_READONLY
                    lstFiles(intFiles) = RuleObject
                    intFiles = intFiles + 1
                Case Else
                    frmRBMsg.RaiseMessage RBOX(10), 87
            End Select
        Case RULE_TYPE_REGISTRY
            Select Case RuleAction
                Case RULE_ACTION_DENY
                    lstRegDeny(intRegDeny) = RuleObject
                    intRegDeny = intRegDeny + 1
                Case RULE_ACTION_READONLY
                    lstReg(intReg) = RuleObject
                    intReg = intReg + 1
                Case Else
                    frmRBMsg.RaiseMessage RBOX(10), 87
            End Select
        Case RULE_TYPE_WINDOW
            Select Case RuleAction
                Case RULE_ACTION_ALLOW
                    lstOpenWin(intOpenWin) = RuleObject
                    intOpenWin = intOpenWin + 1
                Case Else
                    frmRBMsg.RaiseMessage RBOX(10), 87
            End Select
        Case Else
            frmRBMsg.RaiseMessage RBOX(11), 87
    End Select
End Sub

Function IsAdmin() As Boolean
    Dim pAdminSid As Long, b As Long
    
    Call ConvertStringSidToSid("S-1-5-32-544", pAdminSid)
    
    If CheckTokenMembership(0&, pAdminSid, b) Then
        IsAdmin = b
    End If
    
    Call FreeSid(pAdminSid)
End Function

Function SetObjectIntegrity(ByVal szObject As String, ByVal se_type As SE_OBJECT_TYPE, ByVal se_level As SE_INTEGRITY_LEVEL, ByVal se_policy As SE_INTEGRITY_POLICY) As Long
    Dim fStatus As Long, ret As Long
    fStatus = ERROR_SUCCESS
    
    Dim pSD As Long, szSD As String, pSacl As Long
    Dim fSaclPresent As Long, fSaclDefaulted As Long
    
    fSaclPresent = False: fSaclDefaulted = False
    
    szSD = "S:(ML;"
    
    If (se_policy And INHERITENCE) Then
        szSD = szSD & "CIOI;"
    Else
        szSD = szSD & ";"
    End If
    
    If (se_policy And NO_READ_UP) Then
        szSD = szSD & "NRNWNX;;;"
    ElseIf (se_policy And NO_WRITE_UP) Then
        szSD = szSD & "NW;;;"
    End If
    
    Select Case se_level
    Case LOW_INTEGRITY
        szSD = szSD & "LW)"
    Case MEDIUM_INTEGRITY
        szSD = szSD & "ME)"
    Case HIGH_INTEGRITY
        szSD = szSD & "HI)"
    End Select
    
    If ConvertStringSecurityDescriptorToSecurityDescriptorW(StrPtr(szSD), SDDL_REVISION_1, pSD, 0) Then
        If GetSecurityDescriptorSacl(ByVal pSD, fSaclPresent, pSacl, fSaclDefaulted) Then
            ret = SetNamedSecurityInfoA(szObject, se_type, LABEL_SECURITY_INFORMATION, 0, 0, 0, ByVal pSacl)
            
            If ret <> ERROR_SUCCESS Then
                fStatus = ret
            End If
            
            Call LocalFree(pSD)
        Else
            fStatus = Err.LastDllError
        End If
    Else
        fStatus = Err.LastDllError
    End If
    
    SetObjectIntegrity = fStatus
End Function

Function SetFileIntegrity(ByVal szFile As String, ByVal se_level As SE_INTEGRITY_LEVEL, ByVal se_policy As SE_INTEGRITY_POLICY) As Long
    Dim fStatus As Long
    fStatus = ERROR_SUCCESS
    
    fStatus = SetObjectIntegrity(szFile, SE_FILE_OBJECT, se_level, se_policy)
    
    SetFileIntegrity = fStatus
End Function

Function SetRegistryIntegrity(ByVal szKey As String, ByVal se_level As SE_INTEGRITY_LEVEL, ByVal se_policy As SE_INTEGRITY_POLICY) As Long
    Dim fStatus As Long, tmp As String
    fStatus = ERROR_SUCCESS
    
    tmp = szKey
    
    tmp = Replace(tmp, "HKEY_CLASSES_ROOT", "CLASSES_ROOT")
    tmp = Replace(tmp, "HKEY_CURRENT_USER", "CURRENT_USER")
    tmp = Replace(tmp, "HKEY_LOCAL_MACHINE", "MACHINE")
    tmp = Replace(tmp, "HKEY_USERS", "USERS")
    tmp = Replace(tmp, "HKCR", "CLASSES_ROOT")
    tmp = Replace(tmp, "HKCU", "CURRENT_USER")
    tmp = Replace(tmp, "HKLM", "MACHINE")
    tmp = Replace(tmp, "HKUS", "USERS")
    
    fStatus = SetObjectIntegrity(tmp, SE_REGISTRY_KEY, se_level, se_policy)
    
    SetRegistryIntegrity = fStatus
End Function
