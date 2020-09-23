VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Blackout by Slayer - Options"
   ClientHeight    =   2925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4080
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   4080
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAbout 
      Caption         =   "A&bout"
      Height          =   375
      Left            =   713
      TabIndex        =   6
      Top             =   2175
      Width           =   2655
   End
   Begin VB.CommandButton cmdForce 
      Caption         =   "&Force"
      Height          =   495
      Left            =   713
      TabIndex        =   4
      Top             =   1575
      Width           =   1215
   End
   Begin VB.Timer tmrUnload 
      Left            =   3600
      Top             =   120
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "&Remove"
      Height          =   495
      Left            =   2146
      TabIndex        =   1
      Top             =   375
      Width           =   1215
   End
   Begin VB.CommandButton cmdQuickLock 
      Caption         =   "&Quick Lock"
      Height          =   495
      Left            =   744
      TabIndex        =   2
      Top             =   975
      Width           =   1215
   End
   Begin VB.CommandButton cmdIdle 
      Caption         =   "I&dle Detect"
      Height          =   495
      Left            =   2146
      TabIndex        =   3
      Top             =   975
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Default         =   -1  'True
      Height          =   495
      Left            =   2153
      TabIndex        =   5
      Top             =   1575
      Width           =   1215
   End
   Begin VB.CommandButton cmdChangePass 
      Caption         =   "&Password"
      Height          =   495
      Left            =   744
      TabIndex        =   0
      Top             =   375
      Width           =   1215
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAbout_Click()
  frmAbout.Show
End Sub

Private Sub cmdChangePass_Click()
  frmNewPass.Show
End Sub
Private Sub cmdExit_Click()
  Dim DetectEnabled As String
  Dim QuickLock As String
  Dim ForceBlackout As String
  DetectEnabled = GetString(HKEY_LOCAL_MACHINE, "Software\Blackout", "EnableIdleDetect")
  QuickLock = GetString(HKEY_LOCAL_MACHINE, "Software\Blackout", "EnableQuickLock")
  ForceBlackout = GetString(HKEY_LOCAL_MACHINE, "Software\Blackout", "EnableForceBlackout")
  If UCase$(DetectEnabled) = "FALSE" And UCase$(QuickLock) = "FALSE" And UCase$(ForceBlackout) = "FALSE" Then
    Unload Me
    End
  End If
  frmTimers.Visible = False
  frmTimers.Show
End Sub
Private Sub cmdForce_Click()
  frmForce.Show
End Sub
Private Sub cmdIdle_Click()
  frmIdleDetect.Show
End Sub
Private Sub cmdQuickLock_Click()
  frmQuickLock.Show
End Sub
Private Sub cmdRemove_Click()
  Dim MessageReply As Single
  MessageReply = MsgBox("This will remove Blackout from your computer. Proceed?", vbYesNo Or vbQuestion, "Blackout by Slayer")
  If MessageReply = vbYes Then
    DeleteValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "Blackout"
    DeleteKey HKEY_LOCAL_MACHINE, "Software\Blackout"
    Kill PasswordFileLocation
    MsgBox "Blackout has been removed from your system. To re-install this program, execute Blackout.", vbOKOnly Or vbInformation, "Blackout by Slayer"
    Unload Me
    End
  End If
End Sub
Private Sub Form_Load()
  tmrUnload.Interval = 1
  Unload frmNewPass
  Unload frmIdleDetect
  Unload frmQuickLock
  Mouse False
End Sub
Private Sub tmrUnload_Timer()
  Unload frmLockoutScreen
  Mouse False
  StartBar False
End Sub
