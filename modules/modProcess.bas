Attribute VB_Name = "modProcess"
Option Explicit

Private Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Declare Function GetLastError Lib "kernel32.dll" () As Long
Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Declare Function ReadProcessMemory Lib "kernel32.dll" (ByVal hProcess As Long, ByRef lpBaseAddress As Any, ByRef lpBuffer As Any, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Declare Function NtQueryInformationProcess Lib "ntdll.dll" (ByVal ProcessHandle As Long, ByVal ProcessInformationClass As PROCESSINFOCLASS, ByVal ProcessInformation As Long, ByVal ProcessInformationLength As Long, ByRef ReturnLength As Long) As Long
Declare Function TerminateProcess Lib "kernel32.dll" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Declare Function GetProcessImageFileName Lib "psapi.dll" Alias "GetProcessImageFileNameA" (ByVal hProcess As Long, lpImageName As Any, ByVal nSize As Long) As Long
Declare Function NtOpenProcess Lib "ntdll.dll" (ByRef ProcessHandle As Long, ByVal AccessMask As Long, ByRef ObjectAttributes As OBJECT_ATTRIBUTES, ByRef ClientID As CLIENT_ID) As Long
Declare Function NtSuspendProcess Lib "ntdll.dll" (ByVal ProcessHandle As Long) As Long
Declare Function NtResumeProcess Lib "ntdll.dll" (ByVal ProcessHandle As Long) As Long
Declare Function CreateFile Lib "kernel32.dll" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Declare Function WriteFile Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, ByRef lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
Declare Function CreateProcess Lib "kernel32.dll" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, ByRef lpProcessAttributes As SECURITY_ATTRIBUTES, ByRef lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDriectory As String, ByRef lpStartupInfo As STARTUPINFO, ByRef lpProcessInformation As PROCESS_INFORMATION) As Long

Const ERROR_SUCCESS As Long = 0&

Const PROCESS_TERMINATE As Long = (&H1)
Const PROCESS_VM_OPERATION As Long = (&H8)
Const PROCESS_VM_READ As Long = (&H10)
Const PROCESS_VM_WRITE As Long = (&H20)
Const PROCESS_QUERY_INFORMATION As Long = (&H400)
Const PROCESS_SUSPEND_RESUME As Long = (&H800)
Const PROCESS_QUERY_LIMITED_INFORMATION As Long = (&H1000)
Const PROCESS_ALL_ACCESS As Long = (&H1F0FFF)

Const PAGE_SIZE As Long = 4096

Const GENERIC_READ As Long = &H80000000
Const GENERIC_WRITE As Long = &H40000000
Const CREATE_NEW As Long = 1
Const FILE_SHARE_READ As Long = &H1

Const INVALID_HANDLE_VALUE As Long = -1

Enum PROCESSINFOCLASS
    ProcessBasicInformation
    ProcessQuotaLimits
    ProcessIoCounters
    ProcessVmCounters
    ProcessTimes
    ProcessBasePriority
    ProcessRaisePriority
    ProcessDebugPort
    ProcessExceptionPort
    ProcessAccessToken
    ProcessLdtInformation
    ProcessLdtSize
    ProcessDefaultHardErrorMode
    ProcessIoPortHandlers '// Note: this is kernel mode only
    ProcessPooledUsageAndLimits
    ProcessWorkingSetWatch
    ProcessUserModeIOPL
    ProcessEnableAlignmentFaultFixup
    ProcessPriorityClass
    ProcessWx86Information
    ProcessHandleCount
    ProcessAffinityMask
    ProcessPriorityBoost
    ProcessDeviceMap
    ProcessSessionInformation
    ProcessForegroundInformation
    ProcessWow64Information
    ProcessImageFileName
    ProcessLUIDDeviceMapsEnabled
    ProcessBreakOnTermination
    ProcessDebugObjectHandle
    ProcessDebugFlags
    ProcessHandleTracing
    ProcessIoPriority
    ProcessExecuteFlags
    ProcessResourceManagement
    ProcessCookie
    ProcessImageInformation
    MaxProcessInfoClass '// MaxProcessInfoClass should always be the last enum
End Enum

Type PROCESS_BASIC_INFORMATION
    ExitStatus As Long
    PebBaseAddress As Long
    AffinityMask As Long
    BasePriority As Long
    UniqueProcessId As Long
    InheritedFromUniqueProcessId As Long
End Type

Type LIST_ENTRY
    Flink As Long
    Blink As Long
End Type

Type PEB_LDR_DATA
    Length As Long
    Initialized As Long
    SsHandle As Long
    InLoadOrderModuleList As LIST_ENTRY
    InMemoryOrderModuleList As LIST_ENTRY
    InInitializationOrderModuleList As LIST_ENTRY
 End Type
 
Type UNICODE_STRING
    Length As Integer
    MaximumLength As Integer
    buffer As Long
 End Type
 
Type LDR_MODULE
    a_InLoadOrderModuleList As LIST_ENTRY
    b_InMemoryOrderModuleList As LIST_ENTRY
    c_InInitializationOrderModuleList As LIST_ENTRY
    d_BaseAddress As Long
    e_vEntryPoint As Long
    f_SizeOfImage As Long
    g_FullDllName As UNICODE_STRING
    h_BaseDllName As UNICODE_STRING
    i_Flags As Long
    l_LoadCount As Integer
    m_TlsIndex As Integer
    n_HashTableEntry As LIST_ENTRY
    o_TimeDateStamp As Long
End Type

Type RTL_DRIVE_LETTER_CURDIR
    Flags As Integer
    Length As Integer
    TimeStamp As Long
    DosPath As UNICODE_STRING
End Type

Type RTL_USER_PROCESS_PARAMETERS
    a_MaximumLength As Long
    b_Length As Long
    a_Flags As Long
    a_DebugFlags As Long
    c_ConsoleHandle As Long
    d_ConsoleFlags As Long
    e_StdInputHandle As Long
    f_StdOutputHandle As Long
    g_StdErrorHandle As Long
    h_CurrentDirectoryPath As UNICODE_STRING
    i_CurrentDirectoryHandle As Long
    l_DllPath As UNICODE_STRING
    m_ImagePathName As UNICODE_STRING
    n_CommandLine As UNICODE_STRING
    o_Environment As Long
    p_StartingPositionLeft As Long
    q_StartingPositionTop As Long
    r_Width As Long
    s_Height As Long
    t_CharWidth As Long
    u_CharHeight As Long
    v_ConsoleTextAttributes As Long
    z_WindowFlags As Long
    z1_ShowWindowFlags As Long
    z2_WindowTitle As UNICODE_STRING
    z3_DesktopName As UNICODE_STRING
    z4_ShellInfo As Long
    z5_RuntimeData As Long
    z6_DLCurrentDirectory(&H20) As RTL_DRIVE_LETTER_CURDIR
End Type

Type PEB
    InheritedAddressSpace As Byte
    ReadImageFileExecOptions As Byte
    BeingDebugged As Byte
    Reserved1 As Byte
    Mutant As Long
    SectionBaseAddress As Long
    ProcessModuleInfo As Long
    ProcessParameters As Long
    SubSystemData As Long
    ProcessHeap As Long
    FastPebLock As Long
    AcquireFastPebLock As Long
    ReleaseFastPebLock As Long
    EnvironmentUpdateCount As Long
    User32Dispatch As Long
    EventLogSection As Long
    EventLog As Long
    ExecuteOptions As Long
    'FreeList As Long ' // PEB_FREE_BLOCK
    TlsBitMapSize As Long
    TlsBitMap As Long
    TlsBitMapData(1 To 2) As Long
    ReadOnlySharedMemoryBase As Long
    ReadOnlySharedMemoryHeap As Long
    ReadOnlyStaticServerData As Long
    InitAnsiCodePageData As Long
    InitOemCodePageData As Long
    InitUnicodeCaseTableData As Long
    KeNumberProcessors As Long
    NtGlobalFlag As Long
    Reserved9 As Long
    MmCriticalSectionTimeout As Currency
    MmHeapSegmentReserve As Long
    MmHeapSegmentCommit As Long
    MmHeapDeCommitTotalFreeThreshold As Long
    MmHeapDeCommitFreeBlockThreshold As Long
    NumberOfHeaps As Long
    AvailableHeaps As Long
    ProcessHeapsListBuffer As Long
    GdiSharedHandleTable As Long
    ProcessStarterHelper As Long
    GdiInitialBatchLimit As Long
    LoaderLock As Long
    NtMajorVersion As Long
    NtMinorVersion As Long
    NtBuildNumber As Integer
    NtCSDVersion As Integer
    PlatformId As Long
    Subsystem As Long
    MajorSubsystemVersion As Long
    MinorSubsystemVersion As Long
    AffinityMask As Long
    GdiHandleBuffer(33) As Long
    PostProcessInitRoutine As Long
    TlsExpansionBitmap As Long
    TlsExpansionBitmapBits(127) As Byte
    SessionId As Long
    AppCompatFlags(1 To 2) As Long
    AppCompatFlagsUser(1 To 2) As Long
    ShimData As Long
    AppCompatInfo As Long
    CSDVersion As UNICODE_STRING
    ActivationContextData As Long
    ProcessAssemblyStorageMap As Long
    SystemDefaultActivationData As Long
    SystemAssemblyStorageMap As Long
    MinimumStackCommit As Long
    FlsCallBack As Long
    FlsListHead As Long
    FlsBitmap As Long
    FlsBitmapBits(3) As Long
    FlsHighIndex As Long
End Type

Type VM_COUNTERS
    PeakVirtualSize As Long
    VirtualSize As Long
    PageFaultCount As Long
    PeakWorkingSetSize As Long
    WorkingSetSize As Long
    QuotaPeakPagedPoolUsage As Long
    QuotaPagedPoolUsage As Long
    QuotaPeakNonPagedPoolUsage As Long
    QuotaNonPagedPoolUsage As Long
    PagefileUsage As Long
    PeakPagefileUsage As Long
End Type

Type IO_COUNTERS
    ReadOperationCount As Currency
    WriteOperationCount As Currency
    OtherOperationCount As Currency
    ReadTransferCount As Currency
    WriteTransferCount As Currency
    OtherTransferCount As Currency
End Type

Type CLIENT_ID
    UniqueProcess As Long
    UniqueThread As Long
End Type

Type OBJECT_ATTRIBUTES
    Length As Long
    RootDirectory As Long
    ObjectName As Long
    Attributes As Long
    SecurityDescriptor As Long
    SecurityQualityOfService As Long
End Type

Function NT_SUCCESS(ByVal nStatus As Long) As Boolean
    NT_SUCCESS = (nStatus >= 0)
End Function

Function HXOpenProcess(ByVal dwPID As Long, ByVal MinimalAccess As Boolean, Optional ByVal AccessMask As Long) As Long
    Dim hProcess As Long, ci As CLIENT_ID, oa As OBJECT_ATTRIBUTES
    
    If (MinimalAccess) Then
        If (IsVistaOrLater) Then
            hProcess = OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, 0, dwPID)
            
            If (hProcess = 0) Then
                oa.Length = Len(oa)
                ci.UniqueProcess = dwPID
                
                Call NtOpenProcess(hProcess, PROCESS_QUERY_LIMITED_INFORMATION, oa, ci)
            End If
        Else
            hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, 0, dwPID)
            
            If (hProcess = 0) Then
                oa.Length = Len(oa)
                ci.UniqueProcess = dwPID
                
                Call NtOpenProcess(hProcess, PROCESS_QUERY_INFORMATION, oa, ci)
            End If
        End If
    Else
        hProcess = OpenProcess(AccessMask, 0, dwPID)
        
        If (hProcess = 0) Then
            oa.Length = Len(oa)
            ci.UniqueProcess = dwPID
            
            Call NtOpenProcess(hProcess, AccessMask, oa, ci)
        End If
    End If
    
    HXOpenProcess = hProcess
End Function

Function QueryPEBAddress(ByVal hProcess As Long) As Long
    Dim pbi As PROCESS_BASIC_INFORMATION
    
    If (hProcess) Then
        If NT_SUCCESS(NtQueryInformationProcess(hProcess, ProcessBasicInformation, VarPtr(pbi), Len(pbi), 0)) Then
            QueryPEBAddress = pbi.PebBaseAddress
        End If
    End If
End Function

Function GetProcessImagePath(ByVal dwPID As Long) As String
    Dim hProcess As Long, sPEB As PEB, RtlName As RTL_USER_PROCESS_PARAMETERS, bytName(260 * 2 - 1) As Byte, sName As String, sBuff As String * 260
    
    If (dwPID = 4) Then
        GetProcessImagePath = "System"
        Exit Function
    End If
    
    hProcess = HXOpenProcess(dwPID, False, PROCESS_VM_READ)
    
    If (hProcess) Then
        If ReadProcessMemory(hProcess, ByVal QueryPEBAddress(hProcess), sPEB, Len(sPEB), 0) Then
            If ReadProcessMemory(hProcess, ByVal sPEB.ProcessParameters, RtlName, Len(RtlName), 0) Then
                If ReadProcessMemory(hProcess, ByVal RtlName.m_ImagePathName.buffer, bytName(0), 260 * 2, 0) Then
                    sName = bytName
                    If InStr(1, sName, Chr$(0)) Then sName = Left$(sName, Len(sName) - Len(Mid$(sName, InStr(1, sName, Chr$(0)))))
                    GetProcessImagePath = sName
                End If
            End If
        End If
    End If
    
    Call CloseHandle(hProcess)
    
    If sName = "" Then
        hProcess = HXOpenProcess(dwPID, True)
        
        If (hProcess) Then
            If GetProcessImageFileName(hProcess, bytName(0), 260 * 2) Then
                sName = StrConv(bytName, vbUnicode)
                If InStr(1, sName, Chr$(0)) Then sName = Left$(sName, Len(sName) - Len(Mid$(sName, InStr(1, sName, Chr$(0)))))
                GetProcessImagePath = sName
            End If
            
            Call CloseHandle(hProcess)
        End If
    End If
End Function

Function GetProcessCurrentDirectory(ByVal dwPID As Long) As String
    Dim hProcess As Long, sPEB As PEB, RtlName As RTL_USER_PROCESS_PARAMETERS, bytName(260 * 2 - 1) As Byte, sName As String
    
    hProcess = HXOpenProcess(dwPID, False, PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ)
    
    If (hProcess) Then
        If ReadProcessMemory(hProcess, ByVal QueryPEBAddress(hProcess), sPEB, Len(sPEB), 0) Then
            If ReadProcessMemory(hProcess, ByVal sPEB.ProcessParameters, RtlName, Len(RtlName), 0) Then
                If ReadProcessMemory(hProcess, ByVal RtlName.h_CurrentDirectoryPath.buffer, bytName(0), 260 * 2, 0) Then
                    sName = bytName
                    If InStr(1, sName, Chr$(0)) Then sName = Left$(sName, Len(sName) - Len(Mid(sName, InStr(1, sName, Chr$(0)))))
                    GetProcessCurrentDirectory = sName
                End If
            End If
        End If
        
        Call CloseHandle(hProcess)
    End If
End Function

Function GetProcessCommandLine(ByVal dwPID As Long) As String
    Dim hProcess As Long, sPEB As PEB, RtlName As RTL_USER_PROCESS_PARAMETERS, bytName(260 * 2 - 1) As Byte, sName As String
    
    hProcess = HXOpenProcess(dwPID, False, PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ)
    
    If (hProcess) Then
        If ReadProcessMemory(hProcess, ByVal QueryPEBAddress(hProcess), sPEB, Len(sPEB), 0) Then
            If ReadProcessMemory(hProcess, ByVal sPEB.ProcessParameters, RtlName, Len(RtlName), 0) Then
                If ReadProcessMemory(hProcess, ByVal RtlName.n_CommandLine.buffer, bytName(0), 260 * 2, 0) Then
                    sName = bytName
                    If InStr(1, sName, Chr$(0)) Then sName = Left$(sName, Len(sName) - Len(Mid(sName, InStr(1, sName, Chr$(0)))))
                    GetProcessCommandLine = sName
                End If
            End If
        End If
        
        Call CloseHandle(hProcess)
    End If
End Function

Function GetProcessDesktopName(ByVal dwPID As Long) As String
    Dim hProcess As Long, sPEB As PEB, RtlName As RTL_USER_PROCESS_PARAMETERS, bytName(260 * 2 - 1) As Byte, sName As String
    
    hProcess = HXOpenProcess(dwPID, False, PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ)
    
    If (hProcess) Then
        If ReadProcessMemory(hProcess, ByVal QueryPEBAddress(hProcess), sPEB, Len(sPEB), 0) Then
            If ReadProcessMemory(hProcess, ByVal sPEB.ProcessParameters, RtlName, Len(RtlName), 0) Then
                If ReadProcessMemory(hProcess, ByVal RtlName.z3_DesktopName.buffer, bytName(0), 260 * 2, 0) Then
                    sName = bytName
                    If InStr(1, sName, Chr$(0)) Then sName = Left$(sName, Len(sName) - Len(Mid(sName, InStr(1, sName, Chr$(0)))))
                    GetProcessDesktopName = sName
                End If
            End If
        End If
        
        Call CloseHandle(hProcess)
    End If
End Function

Function IsProcessBeingDebugged(ByVal dwPID As Long) As Boolean
    Dim hProcess As Long, sPEB As PEB, BeingDebugged As Long
    
    hProcess = HXOpenProcess(dwPID, False, PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ)
    
    If (hProcess) Then
        If ReadProcessMemory(hProcess, ByVal QueryPEBAddress(hProcess), sPEB, LenB(sPEB), 0) Then
            IsProcessBeingDebugged = CBool(sPEB.BeingDebugged)
        End If
        
        Call CloseHandle(hProcess)
    End If
End Function

Function GetProcessParentPID(ByVal dwPID As Long) As Long
    Dim hProcess As Long, pbi As PROCESS_BASIC_INFORMATION
    
    hProcess = HXOpenProcess(dwPID, True)
    
    If (hProcess) Then
        If NT_SUCCESS(NtQueryInformationProcess(hProcess, ProcessBasicInformation, VarPtr(pbi), Len(pbi), 0)) Then
            GetProcessParentPID = pbi.InheritedFromUniqueProcessId
        End If
        
        Call CloseHandle(hProcess)
    End If
End Function

Function IsProcessAlive(ByVal dwPID As Long) As Boolean
    Dim hProcess As Long, pbi As PROCESS_BASIC_INFORMATION
    
    If (dwPID = 4) Then
        IsProcessAlive = True
        Exit Function
    End If
    
    hProcess = HXOpenProcess(dwPID, True)
    
    If (hProcess) Then
        If NT_SUCCESS(NtQueryInformationProcess(hProcess, ProcessBasicInformation, VarPtr(pbi), Len(pbi), 0)) Then
            IsProcessAlive = (pbi.ExitStatus = 259)
        End If
        
        Call CloseHandle(hProcess)
    End If
End Function

Function EndProcess(ByVal dwPID As Long) As Boolean
    Dim hProcess As Long
    
    hProcess = HXOpenProcess(dwPID, False, PROCESS_TERMINATE)
    
    If (hProcess) Then
        EndProcess = TerminateProcess(hProcess, 0)
        
        Call CloseHandle(hProcess)
    End If
End Function

Function SuspendProcess(ByVal dwPID As Long) As Boolean
    Dim hProcess As Long
    
    hProcess = HXOpenProcess(dwPID, False, PROCESS_SUSPEND_RESUME)
    
    If (hProcess) Then
        SuspendProcess = NT_SUCCESS(NtSuspendProcess(hProcess))
        
        Call CloseHandle(hProcess)
    End If
End Function

Function ResumeProcess(ByVal dwPID As Long) As Boolean
    Dim hProcess As Long
    
    hProcess = HXOpenProcess(dwPID, False, PROCESS_SUSPEND_RESUME)
    
    If (hProcess) Then
        ResumeProcess = NT_SUCCESS(NtResumeProcess(hProcess))
        
        Call CloseHandle(hProcess)
    End If
End Function

