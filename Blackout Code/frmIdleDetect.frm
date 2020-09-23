VERSION 5.00
Begin VB.Form frmIdleDetect 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Black by Slayer - Idle Detect Options"
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmIdleDetect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.VScrollBar ScrollBar 
      Height          =   375
      Left            =   480
      Max             =   0
      Min             =   -45
      TabIndex        =   4
      Top             =   1080
      Value           =   -1
      Width           =   255
   End
   Begin VB.Frame fraIdleDetect 
      Caption         =   "I&dle Detect Settings"
      Height          =   1455
      Left            =   233
      TabIndex        =   0
      Top             =   195
      Width           =   4215
      Begin VB.CommandButton cmdOk 
         Caption         =   "O&k"
         Height          =   375
         Left            =   3240
         TabIndex        =   5
         Top             =   960
         Width           =   855
      End
      Begin VB.CommandButton cmdQuestion 
         Caption         =   "&What is Idle Detect?"
         Height          =   375
         Left            =   2280
         TabIndex        =   2
         Top             =   360
         Width           =   1815
      End
      Begin VB.CheckBox chkEnableIdleDetect 
         Caption         =   "E&nable Idle Detect"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   420
         Width           =   1815
      End
      Begin VB.Label lblExecute 
         Caption         =   "E&xecute Blackout after"
         Height          =   255
         Left            =   600
         TabIndex        =   3
         Top             =   960
         Width           =   3135
      End
   End
End
Attribute VB_Name = "frmIdleDetect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub chkEnableIdleDetect_Click()
  If chkEnableIdleDetect.Value = 0 Then
    ScrollBar.Enabled = False
    cmdQuestion.Enabled = False
  ElseIf chkEnableIdleDetect.Value = 1 Then
    ScrollBar.Enabled = True
    cmdQuestion.Enabled = True
  End If
End Sub
Private Sub cmdOK_Click()
  Dim ForceVal As String
  Dim MessageReply As Single
  If ScrollBar.Value = 0 And chkEnableIdleDetect.Value = 1 Then
    MsgBox "0 is an invalid Idle time. You must select a value greater than 0", vbOKOnly Or vbExclamation, "Blackout by Slayer"
    Exit Sub
  End If
  ForceVal = GetString(HKEY_LOCAL_MACHINE, "Software\Blackout", "EnableForceBlackout")
  If UCase$(ForceVal) = "TRUE" Then
    MessageReply = MsgBox("Force Blackout is currently enabled. Idle Detect cannot be enabled unless Force Blackout is disabled. Disable Force Blackout?", vbYesNo Or vbQuestion, "Blackout by Slayer")
    If MessageReply = vbYes Then
      Call SaveString(HKEY_LOCAL_MACHINE, "Software\Blackout", "EnableIdleDetect", "True")
      Call SaveString(HKEY_LOCAL_MACHINE, "Software\Blackout", "EnableForceBlackout", "False")
      MsgBox "Force Blackout disabled", vbInformation, "Blackout by Slayer"
    Else
      Exit Sub
    End If
  End If
  Select Case chkEnableIdleDetect.Value
    Case 0
      Call SaveString(HKEY_LOCAL_MACHINE, "Software\Blackout", "EnableIdleDetect", "False")
      Call SaveString(HKEY_LOCAL_MACHINE, "Software\Blackout", "IdleDetectTime", Abs(ScrollBar.Value))
    Case 1
      Call SaveString(HKEY_LOCAL_MACHINE, "Software\Blackout", "EnableIdleDetect", "True")
      Call SaveString(HKEY_LOCAL_MACHINE, "Software\Blackout", "IdleDetectTime", Abs(ScrollBar.Value))
  End Select
  frmOptions.Show
End Sub
Private Sub cmdQuestion_Click()
  MsgBox "Idle Detect is a function that executes Blackout after" & vbCr & "a specified time of non-activity on your computer", vbOKOnly Or vbInformation, "Blackout by Slayer"
End Sub
Private Sub Form_Load()
  Unload frmOptions
  Dim Valx As Variant
  Dim NewVal As Integer
  Valx = GetString(HKEY_LOCAL_MACHINE, "Software\Blackout", "EnableIdleDetect")
  If UCase$(Valx) = "TRUE" Then
    chkEnableIdleDetect.Value = 1
    chkEnableIdleDetect_Click
  Else
    chkEnableIdleDetect.Value = 0
    chkEnableIdleDetect_Click
  End If
  Valx = GetString(HKEY_LOCAL_MACHINE, "Software\Blackout", "IdleDetectTime")
  Valx = "-" & Valx
  NewVal = Val(Valx)
  ScrollBar.Value = NewVal
  ScrollBar_Change
End Sub
Private Sub ScrollBar_Change()
  lblExecute.Caption = "E&xecute Blackout after " & Abs(ScrollBar.Value) & " minutes"
End Sub
