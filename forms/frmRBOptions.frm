VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.Ocx"
Begin VB.Form frmRBOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "RightsBox Options"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11715
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   11715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'ËùÓÐÕßÖÐÐÄ
   Begin ComctlLib.TreeView TV 
      Height          =   5415
      Left            =   240
      TabIndex        =   23
      Top             =   240
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   9551
      _Version        =   327682
      HideSelection   =   0   'False
      Indentation     =   617
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
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
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Left            =   10440
      TabIndex        =   4
      Top             =   5880
      Width           =   1095
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
      Height          =   320
      Left            =   9240
      TabIndex        =   3
      Top             =   5880
      Width           =   1095
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
      Height          =   320
      Left            =   8040
      TabIndex        =   2
      Top             =   5880
      Width           =   1095
   End
   Begin VB.Frame fmeSecurity 
      Caption         =   "Drop Rights"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   2760
      TabIndex        =   0
      Top             =   240
      Width           =   8775
      Begin VB.CheckBox chkDropRights 
         Caption         =   "Drop rights from Administrators and Power Users groups"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         TabIndex        =   36
         Top             =   1440
         Width           =   5535
      End
      Begin VB.CheckBox chkDenyUACAdmin 
         Caption         =   "Prevent programs from gaining UAC Admin rights"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         TabIndex        =   30
         Top             =   1920
         Width           =   4935
      End
      Begin VB.Label Label2 
         Caption         =   "To improve the security and isolation, RightsBox can strip some rights from the programs inside the Box."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   1
         Top             =   600
         Width           =   7815
      End
   End
   Begin VB.Frame fmeIL 
      Caption         =   "Integrity Level"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   2760
      TabIndex        =   8
      Top             =   240
      Visible         =   0   'False
      Width           =   8775
      Begin VB.OptionButton optMed 
         Caption         =   "Medium Integrity Level"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   10
         Top             =   1920
         Width           =   3015
      End
      Begin VB.OptionButton optLow 
         Caption         =   "Low Integrity Level"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   9
         Top             =   2280
         Value           =   -1  'True
         Width           =   3255
      End
      Begin VB.Label Label7 
         Caption         =   "Integrity Level was introduced in Windows Vista, designed to improve the security."
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   12
         Top             =   600
         Width           =   7455
      End
      Begin VB.Label Label8 
         Caption         =   "Lowering integrity level can reduce the risk of malware infection. Which integrity level would you want to lower to?"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   11
         Top             =   1080
         Width           =   6495
      End
   End
   Begin VB.Frame fmeUI 
      Caption         =   "User Interface"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   2760
      TabIndex        =   5
      Top             =   240
      Visible         =   0   'False
      Width           =   8775
      Begin VB.CheckBox chkEnableR 
         Caption         =   "Put [R] before window title."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   6
         Top             =   2160
         Value           =   1  'Checked
         Width           =   3735
      End
      Begin VB.Label Label5 
         Caption         =   $"frmRBOptions.frx":0000
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   480
         TabIndex        =   7
         Top             =   600
         Width           =   7455
      End
   End
   Begin VB.Frame fmeOpenWindow 
      Caption         =   "OpenWinClass"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   2760
      TabIndex        =   31
      Top             =   240
      Visible         =   0   'False
      Width           =   8775
      Begin VB.CommandButton cmdRmvOpenWindow 
         Caption         =   "Remove"
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
         Left            =   7320
         TabIndex        =   35
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CommandButton cmdAddOpenWindow 
         Caption         =   "Add"
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
         Left            =   7320
         TabIndex        =   34
         Top             =   1080
         Width           =   1095
      End
      Begin VB.ListBox lstOpenWindow 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3630
         Left            =   480
         TabIndex        =   33
         Top             =   1080
         Width           =   6615
      End
      Begin VB.Label Label12 
         Caption         =   $"frmRBOptions.frx":00C8
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   32
         Top             =   480
         Width           =   6615
      End
   End
   Begin VB.Frame fmeRegDeny 
      Caption         =   "ClosedRegKey"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   2760
      TabIndex        =   24
      Top             =   240
      Visible         =   0   'False
      Width           =   8775
      Begin VB.CommandButton cmdRmvRegDeny 
         Caption         =   "Remove"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7320
         TabIndex        =   27
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CommandButton cmdAddRegDeny 
         Caption         =   "Add"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7320
         TabIndex        =   26
         Top             =   1080
         Width           =   1095
      End
      Begin VB.ListBox lstRegistryDeny 
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3630
         Left            =   480
         TabIndex        =   25
         Top             =   1080
         Width           =   6615
      End
      Begin VB.Label Label11 
         Caption         =   "Note: Blocked > Read-only."
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   29
         Top             =   720
         Width           =   2535
      End
      Begin VB.Label Label10 
         Caption         =   "Programs under the supervision of RightsBox cannot access the following registry keys."
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   28
         Top             =   480
         Width           =   7935
      End
   End
   Begin VB.Frame fmeRegRO 
      Caption         =   "ReadRegKey"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   2760
      TabIndex        =   18
      Top             =   240
      Visible         =   0   'False
      Width           =   8775
      Begin VB.ListBox lstRegistryRO 
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3630
         Left            =   480
         TabIndex        =   21
         Top             =   1080
         Width           =   6615
      End
      Begin VB.CommandButton cmdAddRegRO 
         Caption         =   "Add"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7320
         TabIndex        =   20
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CommandButton cmdRmvRegRO 
         Caption         =   "Remove"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7320
         TabIndex        =   19
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Programs under the supervision of RightsBox cannot write to the following registry keys."
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   22
         Top             =   600
         Width           =   7935
      End
   End
   Begin VB.Frame fmeFileDenied 
      Caption         =   "ClosedFilePath"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   2760
      TabIndex        =   37
      Top             =   240
      Width           =   8775
      Begin VB.CommandButton cmdRmvFileDenied 
         Caption         =   "Remove"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7320
         TabIndex        =   41
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CommandButton cmdAddFileDenied 
         Caption         =   "Add"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7320
         TabIndex        =   40
         Top             =   1080
         Width           =   1095
      End
      Begin VB.ListBox lstFileDenied 
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3630
         Left            =   480
         TabIndex        =   39
         Top             =   1080
         Width           =   6615
      End
      Begin VB.Label Label4 
         Caption         =   "Note: Blocked > Read-only."
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   42
         Top             =   720
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "Programs under the supervision of RightsBox cannot access the following files/folders."
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   38
         Top             =   480
         Width           =   7935
      End
   End
   Begin VB.Frame fmeFileRst 
      Caption         =   "ReadFilePath"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   2760
      TabIndex        =   13
      Top             =   240
      Visible         =   0   'False
      Width           =   8775
      Begin VB.CommandButton cmdRmvFolder 
         Caption         =   "Remove"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7320
         TabIndex        =   17
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CommandButton cmdAddFolder 
         Caption         =   "Add"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7320
         TabIndex        =   16
         Top             =   1080
         Width           =   1095
      End
      Begin VB.ListBox lstFolder 
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3630
         Left            =   480
         TabIndex        =   15
         Top             =   1080
         Width           =   6615
      End
      Begin VB.Label Label9 
         Caption         =   "Programs under the supervision of RightsBox cannot write to the following files/folders."
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   14
         Top             =   600
         Width           =   7935
      End
   End
End
Attribute VB_Name = "frmRBOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAddFileDenied_Click()
    If intFilesDeny < 50 Then
        Dim szPath As String
        szPath = InputBox("Enter a file (or folder) path")
        If szPath <> "" Then
            lstFileDenied.AddItem szPath
            Call AddRule(RULE_TYPE_FILE_SYSTEM, RULE_ACTION_DENY, szPath)
        End If
    Else
        MsgBox "50 files (or folders) at most!"
    End If
End Sub

Private Sub cmdAddFolder_Click()
    If intFiles < 50 Then
        Dim szPath As String
        szPath = InputBox("Enter a file (or folder) path")
        If szPath <> "" Then
            lstFolder.AddItem szPath
            Call AddRule(RULE_TYPE_FILE_SYSTEM, RULE_ACTION_READONLY, szPath)
        End If
    Else
        MsgBox "50 files (or folders) at most!"
    End If
End Sub

Private Sub cmdAddOpenWindow_Click()
    If intOpenWin < 50 Then
        Dim szWinCls As String
        szWinCls = InputBox("Enter a window class name")
        If szWinCls <> "" Then
            lstOpenWindow.AddItem szWinCls
            Call AddRule(RULE_TYPE_WINDOW, RULE_ACTION_ALLOW, szWinCls)
        End If
    Else
        MsgBox "50 window class names at most!"
    End If
End Sub

Private Sub cmdAddRegDeny_Click()
    If intRegDeny < 50 Then
        Dim szPath As String
        szPath = InputBox("Enter a registry path")
        If szPath <> "" Then
            lstRegistryDeny.AddItem szPath
            Call AddRule(RULE_TYPE_REGISTRY, RULE_ACTION_DENY, szPath)
        End If
    Else
        MsgBox "50 registry keys at most!"
    End If
End Sub

Private Sub cmdAddRegRO_Click()
    If intReg < 50 Then
        Dim szPath As String
        szPath = InputBox("Enter a registry path")
        If szPath <> "" Then
            lstRegistryRO.AddItem szPath
            Call AddRule(RULE_TYPE_REGISTRY, RULE_ACTION_READONLY, szPath)
        End If
    Else
        MsgBox "50 registry keys at most!"
    End If
End Sub

Private Sub cmdApply_Click()
    bDropRights = IIf(chkDropRights.Value, 1, 0)
    SetWinTextR = IIf(chkEnableR.Value, 1, 0)
    DenyUACAdmin = IIf(chkDenyUACAdmin.Value, 1, 0)
    If optMed.Value Then dwIL = 1
    If optLow.Value Then dwIL = 2
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    cmdApply_Click
    Unload Me
End Sub

Private Sub cmdRmvFileDenied_Click()
    On Error GoTo ErrHandler
    Dim i As Long
    
    lstFileDenied.RemoveItem lstFileDenied.ListIndex
    
    For i = 0 To intFilesDeny - 1
        lstFilesDeny(i) = ""
    Next
    
    intFilesDeny = lstFileDenied.ListCount
    
    For i = 0 To intFilesDeny - 1
        lstFilesDeny(i) = lstFileDenied.List(i)
    Next
    
ErrHandler:
    Exit Sub
End Sub

Private Sub cmdRmvFolder_Click()
    On Error GoTo ErrHandler
    Dim i As Long
    
    lstFolder.RemoveItem lstFolder.ListIndex
    
    For i = 0 To intFiles - 1
        lstFiles(i) = ""
    Next
    
    intFiles = lstFolder.ListCount
    
    For i = 0 To intFiles - 1
        lstFiles(i) = lstFolder.List(i)
    Next
    
ErrHandler:
    Exit Sub
End Sub

Private Sub cmdRmvOpenWindow_Click()
    On Error GoTo ErrHandler
    Dim i As Long
    
    lstOpenWindow.RemoveItem lstOpenWindow.ListIndex
    
    For i = 0 To intOpenWin - 1
        lstOpenWin(i) = ""
    Next
    
    intOpenWin = lstOpenWindow.ListCount
    
    For i = 0 To intOpenWin - 1
        lstOpenWin(i) = lstOpenWindow.List(i)
    Next
    
ErrHandler:
    Exit Sub
End Sub

Private Sub cmdRmvRegDeny_Click()
    On Error GoTo ErrHandler
    Dim i As Long
    
    lstRegistryDeny.RemoveItem lstRegistryDeny.ListIndex
    
    For i = 0 To intRegDeny - 1
        lstRegDeny(i) = ""
    Next
    
    intRegDeny = lstRegistryDeny.ListCount
    
    For i = 0 To intRegDeny - 1
        lstRegDeny(i) = lstRegistryDeny.List(i)
    Next
    
ErrHandler:
    Exit Sub
End Sub

Private Sub cmdRmvRegRO_Click()
    On Error GoTo ErrHandler
    Dim i As Long
    
    lstRegistryRO.RemoveItem lstRegistryRO.ListIndex
    
    For i = 0 To intReg - 1
        lstReg(i) = ""
    Next
    
    intReg = lstRegistryRO.ListCount
    
    For i = 0 To intReg - 1
        lstReg(i) = lstRegistryRO.List(i)
    Next
    
ErrHandler:
    Exit Sub
End Sub

Private Sub InitTree()
    TV.Nodes.Add , , "Security", "Security"
    TV.Nodes.Add "Security", tvwChild, "LowerRights", "Drop Rights"
    TV.Nodes.Add , , "UI", "User Interface"
    TV.Nodes.Add , , "ResAccess", "Resource Access"
    If (IsVistaOrLater) Then
        TV.Nodes.Add "Security", tvwChild, "IL", "Integrity Level"
        TV.Nodes.Add "ResAccess", tvwChild, "FileAccess", "File Access"
        TV.Nodes.Add "FileAccess", tvwChild, "FileDenyAccess", "Blocked Access"
        TV.Nodes.Add "FileAccess", tvwChild, "FileReadOnlyAccess", "Read-only Access"
        TV.Nodes.Add "ResAccess", tvwChild, "RegAccess", "Registry Access"
        TV.Nodes.Add "RegAccess", tvwChild, "RegDenyAccess", "Blocked Access"
        TV.Nodes.Add "RegAccess", tvwChild, "RegReadOnlyAccess", "Read-only Access"
    End If
    TV.Nodes.Add "ResAccess", tvwChild, "WinAccess", "Window Access"
End Sub

Private Sub Form_Load()
    Call InitTree
    Me.Icon = LoadPicture("")
    chkEnableR.Value = IIf(SetWinTextR, 1, 0)
    chkDenyUACAdmin.Value = IIf(DenyUACAdmin, 1, 0)
    chkDenyUACAdmin.Visible = IIf(IsVistaOrLater, True, False)
    chkDropRights.Value = IIf(bDropRights, 1, 0)
    If IsAdmin Then chkDenyUACAdmin.Visible = False: chkDenyUACAdmin.Value = 0
    
    If dwIL = 1 Then
        optMed.Value = 1
    Else
        optLow.Value = 1
    End If
    
    Dim i As Integer
    For i = 0 To intFiles - 1
        lstFolder.AddItem lstFiles(i)
    Next
    For i = 0 To intFilesDeny - 1
        lstFileDenied.AddItem lstFilesDeny(i)
    Next
    For i = 0 To intReg - 1
        lstRegistryRO.AddItem lstReg(i)
    Next
    For i = 0 To intRegDeny - 1
        lstRegistryDeny.AddItem lstRegDeny(i)
    Next
    For i = 0 To intOpenWin - 1
        lstOpenWindow.AddItem lstOpenWin(i)
    Next
End Sub

Private Sub TV_NodeClick(ByVal Node As ComctlLib.Node)
    Select Case Node.Key
        Case "LowerRights"
            fmeSecurity.Visible = True
            fmeUI.Visible = False
            fmeIL.Visible = False
            fmeFileRst.Visible = False
            fmeFileDenied.Visible = False
            fmeRegRO.Visible = False
            fmeRegDeny.Visible = False
            fmeOpenWindow.Visible = False
        Case "IL"
            fmeSecurity.Visible = False
            fmeUI.Visible = False
            fmeIL.Visible = True
            fmeFileRst.Visible = False
            fmeFileDenied.Visible = False
            fmeRegRO.Visible = False
            fmeRegDeny.Visible = False
            fmeOpenWindow.Visible = False
        Case "UI"
            fmeSecurity.Visible = False
            fmeUI.Visible = True
            fmeIL.Visible = False
            fmeFileRst.Visible = False
            fmeFileDenied.Visible = False
            fmeRegRO.Visible = False
            fmeRegDeny.Visible = False
            fmeOpenWindow.Visible = False
        Case "FileDenyAccess"
            fmeSecurity.Visible = False
            fmeUI.Visible = False
            fmeIL.Visible = False
            fmeFileRst.Visible = False
            fmeFileDenied.Visible = True
            fmeRegRO.Visible = False
            fmeRegDeny.Visible = False
            fmeOpenWindow.Visible = False
        Case "FileReadOnlyAccess"
            fmeSecurity.Visible = False
            fmeUI.Visible = False
            fmeIL.Visible = False
            fmeFileRst.Visible = True
            fmeFileDenied.Visible = False
            fmeRegRO.Visible = False
            fmeRegDeny.Visible = False
            fmeOpenWindow.Visible = False
        Case "RegReadOnlyAccess"
            fmeSecurity.Visible = False
            fmeUI.Visible = False
            fmeIL.Visible = False
            fmeFileRst.Visible = False
            fmeFileDenied.Visible = False
            fmeRegRO.Visible = True
            fmeRegDeny.Visible = False
            fmeOpenWindow.Visible = False
        Case "RegDenyAccess"
            fmeSecurity.Visible = False
            fmeUI.Visible = False
            fmeIL.Visible = False
            fmeFileRst.Visible = False
            fmeFileDenied.Visible = False
            fmeRegRO.Visible = False
            fmeRegDeny.Visible = True
            fmeOpenWindow.Visible = False
        Case "WinAccess"
            fmeSecurity.Visible = False
            fmeUI.Visible = False
            fmeIL.Visible = False
            fmeFileRst.Visible = False
            fmeFileDenied.Visible = False
            fmeRegRO.Visible = False
            fmeRegDeny.Visible = False
            fmeOpenWindow.Visible = True
    End Select
End Sub
