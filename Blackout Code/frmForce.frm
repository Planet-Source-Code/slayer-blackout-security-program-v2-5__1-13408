VERSION 5.00
Begin VB.Form frmForce 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Blackout by Slayer - Force Blackout"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4830
   Icon            =   "frmForce.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   4830
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraForce 
      Caption         =   "&Force Blackout Options"
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      Begin VB.Frame lblForceByTime 
         Caption         =   "&Enter Time to Force Blackout:"
         Height          =   855
         Left            =   600
         TabIndex        =   9
         Top             =   1680
         Width           =   2415
         Begin VB.OptionButton optPM 
            Caption         =   "P&M"
            Height          =   195
            Left            =   960
            TabIndex        =   12
            Top             =   405
            Value           =   -1  'True
            Width           =   615
         End
         Begin VB.OptionButton optAM 
            Caption         =   "&AM"
            Height          =   195
            Left            =   1560
            TabIndex        =   11
            Top             =   405
            Width           =   735
         End
         Begin VB.TextBox txtForceTime 
            Height          =   285
            Left            =   240
            TabIndex        =   10
            Text            =   "12:00"
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.Frame fraForceProperties 
         Caption         =   "&Properties"
         Height          =   855
         Left            =   600
         TabIndex        =   6
         Top             =   720
         Width           =   1695
         Begin VB.OptionButton optUseTime 
            Caption         =   "Use &Time"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton optUseInterval 
            Caption         =   "&Use Interval"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   480
            Width           =   1335
         End
      End
      Begin VB.VScrollBar ScrollBar 
         Height          =   375
         Left            =   600
         Max             =   0
         Min             =   -45
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1680
         Width           =   255
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "O&k"
         Default         =   -1  'True
         Height          =   375
         Left            =   3480
         TabIndex        =   3
         Top             =   1680
         Width           =   855
      End
      Begin VB.CommandButton cmdQuestion 
         Caption         =   "&What is Force Blackout?"
         Height          =   375
         Left            =   2280
         TabIndex        =   2
         Top             =   360
         Width           =   2175
      End
      Begin VB.CheckBox chkEnableForce 
         Caption         =   "E&nable Force Blackout"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   420
         Width           =   1935
      End
      Begin VB.Label lblForceTime 
         Caption         =   "Force &Blackout after 0 minutes"
         Height          =   255
         Left            =   960
         TabIndex        =   4
         Top             =   1740
         Width           =   2535
      End
   End
End
Attribute VB_Name = "frmForce"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ForceByTime As String
Dim ForceByInterval As String
Private Sub chkEnableForce_Click()
  Select Case chkEnableForce.Value
    Case 0
      cmdQuestion.Enabled = False
    Case 1
      cmdQuestion.Enabled = True
  End Select
End Sub
Private Sub cmdOK_Click()
  Dim IdleVal As String
  Dim MessageReply As Single
  Dim ForceTime As String
  If optUseInterval.Value = True Then
    If Abs(ScrollBar.Value) = 0 Then
      MsgBox "0 is an invalid Force Time. Please enter a new time to Force Blackout.", vbOKOnly Or vbExclamation, "Blackout by Slayer"
      Exit Sub
    End If
  End If
  If optUseTime.Value = True Then
    If optAM.Value = True Then
      ForceTime = txtForceTime.Text & ":00" & " AM"
    ElseIf optPM.Value = True Then
      ForceTime = txtForceTime.Text & ":00" & " PM"
    End If
  ElseIf optUseInterval.Value = True Then
    ForceTime = Abs(ScrollBar.Value)
  End If
  If chkEnableForce.Value = 1 Then
    IdleVal = GetString(HKEY_LOCAL_MACHINE, "Software\Blackout", "EnableIdleDetect")
    If UCase$(IdleVal) = "TRUE" Then
      MessageReply = MsgBox("You currently have Idle Detect enabled. Force Blackout can only be enabled while Idle Detect is disabled. Disable Idle Detect?", vbYesNo Or vbQuestion, "Blackout by Slayer")
      If MessageReply = vbYes Then
        Call SaveString(HKEY_LOCAL_MACHINE, "Software\Blackout", "EnableIdleDetect", "False")
        Call SaveString(HKEY_LOCAL_MACHINE, "Software\Blackout", "EnableForceBlackout", "True")
        Call SaveString(HKEY_LOCAL_MACHINE, "Software\Blackout", "ForceBlackoutTime", ForceTime)
        If optUseTime.Value = True Then
          Call SaveString(HKEY_LOCAL_MACHINE, "Software\Blackout", "ForceByTime", "True")
          Call SaveString(HKEY_LOCAL_MACHINE, "Sofware\Blackout", "ForceByInterval", "False")
        ElseIf optUseTime.Value = False Then
          Call SaveString(HKEY_LOCAL_MACHINE, "Software\Blackout", "ForceByTime", "False")
          Call SaveString(HKEY_LOCAL_MACHINE, "Software\Blackout", "ForceByInterval", "True")
        End If
        MsgBox "Idle Detect disabled", vbOKOnly Or vbInformation, "Blackout by Slayer"
        Unload Me
      End If
    ElseIf UCase$(IdleVal) = "FALSE" Then
      Call SaveString(HKEY_LOCAL_MACHINE, "Software\Blackout", "EnableForceBlackout", "True")
      Call SaveString(HKEY_LOCAL_MACHINE, "Software\Blackout", "ForceBlackoutTime", ForceTime)
      If optUseTime.Value = True Then
        Call SaveString(HKEY_LOCAL_MACHINE, "Software\Blackout", "ForceByTime", "True")
        Call SaveString(HKEY_LOCAL_MACHINE, "Software\Blackout", "ForceByInterval", "False")
      ElseIf optUseTime.Value = True Then
        Call SaveString(HKEY_LOCAL_MACHINE, "Software\Blackout", "ForceByTime", "False")
        Call SaveString(HKEY_LOCAL_MACHINE, "Software\Blackout", "ForceByInterval", "True")
      End If
    End If
  ElseIf chkEnableForce.Value = 0 Then
    Call SaveString(HKEY_LOCAL_MACHINE, "Software\Blackout", "EnableForceBlackout", "False")
  End If
  If optUseTime.Value = True Then
    Call SaveString(HKEY_LOCAL_MACHINE, "Software\Blackout", "ForceByTime", "True")
    Call SaveString(HKEY_LOCAL_MACHINE, "Software\Blackout", "ForceByInterval", "False")
  ElseIf optUseInterval.Value = True Then
    Call SaveString(HKEY_LOCAL_MACHINE, "Software\Blackout", "ForcebyTime", "False")
    Call SaveString(HKEY_LOCAL_MACHINE, "Software\Blackout", "ForceByInterval", "True")
  End If
  Unload Me
End Sub
Private Sub cmdQuestion_Click()
  MsgBox "Force Blackout is a function that executes Blackout after a specified amount of time", vbOKOnly Or vbInformation, "Blackout by Slayer"
End Sub
Private Sub Form_Load()
  Dim EnableForceVal As String
  Dim TimeVal As Variant
  Dim TheTime As String
  Dim Counting As Long
  Dim ForceOp As String
  EnableForceVal = GetString(HKEY_LOCAL_MACHINE, "Software\Blackout", "EnableForceBlackout")
  Select Case UCase$(EnableForceVal)
    Case "TRUE"
      chkEnableForce.Value = 1
    Case "FALSE"
      chkEnableForce.Value = 0
  End Select
  TimeVal = GetString(HKEY_LOCAL_MACHINE, "Software\Blackout", "ForceBlackoutTime")
  Select Case UCase$(Right$(TimeVal, 1))
    Case "M"
      For Counting = 1 To Len(TimeVal)
        If Mid$(TimeVal, Counting, 1) <> ":" Then
          TheTime = TheTime & Mid$(TimeVal, Counting, 1)
        ElseIf Mid$(TimeVal, Counting, 1) = ":" Then
          Exit For
        End If
      Next Counting
      TheTime = Mid$(TimeVal, 1, Counting + 2)
      txtForceTime.Text = TheTime
      If UCase$(Right$(TimeVal, 2)) = "AM" Then
        optAM.Value = True
      ElseIf UCase$(Right$(TimeVal, 2)) = "PM" Then
        optPM.Value = True
      End If
    Case Is <> "M"
      ScrollBar.Value = Val(TimeVal * -1)
      ScrollBar_Change
  End Select
  ForceByInterval = GetString(HKEY_LOCAL_MACHINE, "Software\Blackout", "ForceByTime")
  If UCase$(ForceByInterval) = "FALSE" Then
    optUseInterval.Value = True
    optUseTime.Value = False
  ElseIf UCase$(ForceByInterval) = "TRUE" Then
    optUseInterval.Value = False
    optUseTime.Value = True
  End If
  optUseInterval_Click
  optUseTime_Click
  Unload frmOptions
End Sub
Private Sub Form_Unload(Cancel As Integer)
  frmOptions.Show
End Sub
Private Sub optUseInterval_Click()
  If optUseInterval.Value = True Then
    ScrollBar.Visible = True
    lblForceTime.Visible = True
    lblForceByTime.Visible = False
  End If
End Sub
Private Sub optUseTime_Click()
  If optUseTime.Value = True Then
    ScrollBar.Visible = False
    lblForceTime.Visible = False
    lblForceByTime.Visible = True
  End If
End Sub
Private Sub ScrollBar_Change()
  lblForceTime.Caption = "Force &Blackout after " & Abs(ScrollBar.Value) & " minutes"
End Sub
