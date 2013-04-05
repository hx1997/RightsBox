VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmRBStart 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Run - RightsBox"
   ClientHeight    =   2070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7530
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   7530
   StartUpPosition =   1  '所有者中心
   Begin VB.CheckBox chkUAC 
      Caption         =   "Run as UAC Admin"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   1440
      Width           =   2775
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "可执行程序 (*.exe)|*.exe|所有文件 (*.*)|*.*"
      InitDir         =   "%SystemDrive%"
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
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
      Left            =   5640
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
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
      Left            =   4200
      TabIndex        =   3
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   375
      Left            =   6840
      TabIndex        =   2
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox txtPath 
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
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   6495
   End
   Begin VB.Label Label1 
      Caption         =   "Type the program path, and RightsBox will open it for you."
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
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   5415
   End
End
Attribute VB_Name = "frmRBStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, lpProcessAttributes As SECURITY_ATTRIBUTES, lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDriectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function IsProcessInJob Lib "kernel32" (ByVal hProcess As Long, ByVal JobHandle As Long, ByRef Result As Boolean) As Boolean
Private Declare Function MessageBox Lib "user32.dll" Alias "MessageBoxA" (ByVal hWnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Const ICC_USEREX_CLASSES = &H200
Private Const SW_SHOWNORMAL = 1
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const FORMAT_MESSAGE_IGNORE_InsertS = &H200
Private Const ERROR_FILE_NOT_FOUND = 2&
Private Const ERROR_PATH_NOT_FOUND = 3&
Private Const ERROR_BAD_FORMAT = 11&
Private Const SE_ERR_ACCESSDENIED = 5
Private Const SE_ERR_ASSOCINCOMPLETE = 27
Private Const SE_ERR_DDEBUSY = 30
Private Const SE_ERR_DDEFAIL = 29
Private Const SE_ERR_DDETIMEOUT = 28
Private Const SE_ERR_DLLNOTFOUND = 32
Private Const SE_ERR_FNF = 2
Private Const SE_ERR_NOASSOC = 31
Private Const SE_ERR_OOM = 8
Private Const SE_ERR_PNF = 3
Private Const SE_ERR_SHARE = 26

Private Type tagInitCommonControlsEx
        lngSize As Long
        lngICC As Long
End Type

Dim strOperation As String

Private Sub chkUAC_Click()
    strOperation = IIf(chkUAC, "runas", "")
End Sub

Private Sub cmdBrowse_Click()
    CommonDialog1.Flags = cdlOFNHideReadOnly
    CommonDialog1.ShowOpen
    txtPath = CommonDialog1.FileName
End Sub

Private Sub cmdCancel_Click()
    End
End Sub

Private Sub cmdOK_Click()
    Dim ret As Long, sa As SECURITY_ATTRIBUTES, pi As PROCESS_INFORMATION, si As STARTUPINFO
    Me.Hide
    ret = ShellExecute(0, strOperation, txtPath, vbNullString, Environ("SystemDrive"), SW_SHOWNORMAL)
    If (ret <= 32) Then
        Dim LastError As Long
        LastError = Err.LastDllError
        frmErrorMsg!Label1.Caption = "Failed starting process " & vbCrLf & txtPath & vbCrLf & vbCrLf & "Error message" & vbCrLf & FormatErrorMessage(LastError) & " (" & LastError & ")"
        frmErrorMsg.Show vbModal
        Me.Show
        txtPath.SetFocus
    Else
        End
    End If
End Sub

Private Sub Form_Initialize()
    InitCommonControlsVB
End Sub

Private Sub Form_Load()
    Dim IsInJob As Boolean, ret As Long
    
    ret = IsProcessInJob(-1, 0, IsInJob)
    
    If (ret) Then
        If IsInJob = False Then
            MessageBox Me.hWnd, "Not started by RightsBox, exiting.", "Starter for RightsBox", 0
            End
        End If
    Else
        MessageBox Me.hWnd, "Unable to get some info, exiting.", "Starter for RightsBox", 0
        End
    End If
    
    chkUAC.Visible = IIf(IsVistaOrLater And (Not IsAdmin), True, False)
End Sub

Private Sub txtPath_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        cmdOK_Click
    End If
End Sub

Private Function FormatErrorMessage(ByVal ErrID As Long) As String
    Dim astr As String
    Dim bstr As String
    Dim l As Long
     
    astr = String$(256, 20)
    l = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or _
    FORMAT_MESSAGE_IGNORE_InsertS, 0&, ErrID, 0&, _
    astr, Len(astr), ByVal 0)
    If l Then
        bstr = Left$(astr, InStr(astr, Chr(10)) - 2)
        FormatErrorMessage = bstr
    End If
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
