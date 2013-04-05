VERSION 5.00
Begin VB.Form frmRBMsg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "RightsBox Message"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   8820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Timer Timer1 
      Interval        =   800
      Left            =   0
      Top             =   3720
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
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
      Left            =   3840
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3240
      Width           =   1095
   End
   Begin VB.ListBox lstMsg 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2760
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   8295
   End
End
Attribute VB_Name = "frmRBMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub

Sub RaiseMessage(ByVal szMsg As String, ByVal dwError As Long)
    lstMsg.AddItem szMsg
    lstMsg.AddItem "Error code: " & dwError & " | " & GetLastDllErr(dwError)
    Me.Show
End Sub

Private Sub Form_Load()
    Me.Icon = LoadPicture("")
End Sub

Private Sub Timer1_Timer()
    If cmdClose.BackColor = &H8000000F Then
        cmdClose.BackColor = &H0&
    Else
        cmdClose.BackColor = &H8000000F
    End If
End Sub
