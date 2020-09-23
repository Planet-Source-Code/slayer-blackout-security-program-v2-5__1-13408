VERSION 5.00
Begin VB.Form frmTimers 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   90
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   90
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   90
   ScaleWidth      =   90
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer tmrQuickLock 
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer tmrIdleDetect 
      Left            =   525
      Top             =   0
   End
   Begin VB.Timer tmrForceBlackout 
      Left            =   1440
      Top             =   240
   End
End
Attribute VB_Name = "frmTimers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim OldX, OldY As Single
Dim IdleTime As String
Dim ForceTime As String
Dim ForceByInterval As String
Dim Minutes As Integer
Dim Seconds As Integer
Dim ForceByTime As String
Private Sub Form_Load()
  Dim IdleDetect As String
  Dim ForceBlackout As String
  Dim QuickLock As String
  Dim EnableForceBlackout
  Me.Visible = False
  Unload frmOptions
  IdleDetect = GetString(HKEY_LOCAL_MACHINE, "Software\Blackout", "EnableIdleDetect")
  If UCase$(IdleDetect) = "TRUE" Then
    IdleTime = GetString(HKEY_LOCAL_MACHINE, "Software\Blackout", "IdleDetectTime")
    IdleTime = Val(IdleTime)
    IdleCount = 0
    tmrIdleDetect.Interval = 1000
  ElseIf UCase$(IdleDetect) = "FALSE" Then
    EnableForceBlackout = GetString(HKEY_LOCAL_MACHINE, "Software\Blackout", "EnableForceBlackout")
    Debug.Print "EnableForceBlackout = " & EnableForceBlackout
    If UCase$(EnableForceBlackout) = "TRUE" Then
      ForceByTime = GetString(HKEY_LOCAL_MACHINE, "Software\Blackout", "ForceByTime")
      Debug.Print "xx ForceByTime = " & ForceByTime
      If UCase$(ForceByTime) = "TRUE" Then
        ForceByTime = "True"
        Debug.Print "ForceByTime = True"
        ForceTime = GetString(HKEY_LOCAL_MACHINE, "Software\Blackout", "ForceBlackoutTime")
        tmrForceBlackout.Interval = 100
      ElseIf UCase$(ForceByTime) = "FALSE" Then
        ForceByTime = "False"
        ForceTime = GetString(HKEY_LOCAL_MACHINE, "Software\Blackout", "ForceBlackoutTime")
        tmrForceBlackout.Interval = 1000
      End If
    End If
  End If
  QuickLock = GetString(HKEY_LOCAL_MACHINE, "Software\Blackout", "EnableQuickLock")
  If UCase$(QuickLock) = "TRUE" Then
    tmrQuickLock.Interval = 100
    OldX = 0
    OldY = 0
  End If
  Me.Visible = False
End Sub
Private Sub tmrForceBlackout_Timer()
  Me.Visible = False
  Select Case UCase$(ForceByTime)
    Case "TRUE"
      If Time = ForceTime Then
        tmrForceBlackout.Enabled = False
        tmrIdleDetect.Enabled = False
        tmrQuickLock.Enabled = False
        Unload Me
        frmLockoutScreen.Show
      End If
    Case "FALSE"
      Seconds = Seconds + 1
      If Seconds = 60 Then
        Seconds = 0
        Minutes = Minutes + 1
        If Minutes = ForceTime Then
          tmrForceBlackout.Enabled = False
          tmrQuickLock.Enabled = False
          tmrIdleDetect.Enabled = False
          Unload Me
          frmLockoutScreen.Show
        End If
      End If
  End Select
End Sub
Private Sub tmrIdleDetect_Timer()
  Me.Visible = False
  Dim CursorPos As POINTAPI
  GetCursorPos CursorPos
  If OldX = 0 And OldY = 0 Then
    OldX = CursorPos.x
    OldY = CursorPos.y
  End If
  If CursorPos.x = OldX And CursorPos.y = OldY Then
    IdleCount = IdleCount + 1
    OldX = CursorPos.x
    OldY = CursorPos.y
  ElseIf CursorPos.x <> OldX Or CursorPos.y <> OldY Then
    OldX = 0
    OldY = 0
    IdleCount = 0
  End If
  If IdleCount = IdleTime * 60 Then
    tmrIdleDetect.Enabled = False
    tmrQuickLock.Enabled = False
    Unload Me
    frmLockoutScreen.Show
  End If
End Sub
Private Sub tmrQuickLock_Timer()
  Me.Visible = False
  If GetAsyncKeyState(VK_CONTROL) And GetAsyncKeyState(VK_F12) Then
    tmrQuickLock.Enabled = False
    tmrIdleDetect.Enabled = False
    Unload Me
    frmLockoutScreen.Show
  End If
End Sub
