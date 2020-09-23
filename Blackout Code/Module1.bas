Attribute VB_Name = "Module1"
Option Explicit
Sub StartBar(Hide As Boolean)
  Select Case Hide
    Case True
      Dim Handle As Long
      Handle = FindWindow("Shell_TrayWnd", vbNullString)
      ShowWindow Handle, 0
    Case False
      Dim TrayHandle As Long
      TrayHandle = FindWindow("Shell_TrayWnd", vbNullString)
      ShowWindow TrayHandle, 1
  End Select
End Sub
Sub Mouse(Hide As Boolean)
  Select Case Hide
    Case True
      ShowCursor 0
    Case False
      ShowCursor 1
  End Select
End Sub
Sub CheckForPasswordFile()
  Dim FileExists As String
  Dim fn As Integer
  CheckedForPass = 212
  FileExists = ""
  FileExists = Dir$(PasswordFileLocation)
  If FileExists = "" Then
    NoPassFile = True
    frmNewPass.Show
    CreateRegSettings
  ElseIf FileExists <> "" Then
    fn = FreeFile
    Open PasswordFileLocation For Input As #fn
      Input #fn, Password
    Close #fn
    NoPassFile = False
    StartupStuff
  End If
End Sub
Sub CreateRegSettings()
  Dim AppName As String
  RegCreateKey HKEY_LOCAL_MACHINE, "Software\Blackout", 0
  Call SaveString(HKEY_LOCAL_MACHINE, "Software\Blackout", "EnableQuickLock", "False")
  Call SaveString(HKEY_LOCAL_MACHINE, "Software\Blackout", "EnableIdleDetect", "False")
  Call SaveString(HKEY_LOCAL_MACHINE, "Software\Blackout", "EnableForceBlackout", "False")
  Call SaveString(HKEY_LOCAL_MACHINE, "Software\Blackout", "ForceBlackoutTime", "0")
  Call SaveString(HKEY_LOCAL_MACHINE, "Software\Blackout", "IdleDetectTime", "0")
  Call SaveString(HKEY_LOCAL_MACHINE, "Software\Blackout", "ForceByTime", "False")
  Call SaveString(HKEY_LOCAL_MACHINE, "Software\Blackout", "ForceByInterval", "True")
  AppName = App.Path
  If Right(App.Path, 1) <> "\" Then
    AppName = AppName & "\"
  End If
  AppName = AppName & App.EXEName & ".exe"
  Call SaveString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "Blackout", AppName)
End Sub
Sub PutFormOnTop()
  With frmLockoutScreen
    SetWindowPos .hwnd, HWND_TOPMOST, .Left / 15, .Top / 15, .Width / 15, .Height / 15, SWP_NOACTIVATE Or SWP_SHOWWINDOW
  End With
End Sub
Sub DoCAD(DisableCAD As Boolean)
  Dim RetVal As Integer
  Dim pOld As Boolean
  Select Case DisableCAD
    Case True
      RetVal = SystemParametersInfo(SPI_SCREENSAVERRUNNING, True, pOld, 0)
    Case False
     RetVal = SystemParametersInfo(SPI_SCREENSAVERRUNNING, False, pOld, 0)
  End Select
End Sub
Sub StartupStuff()
  With frmLockoutScreen
    .Show
    .Move 0, 0, Screen.Width, Screen.Height
    .txtPassword.Move Screen.Height / 2, Screen.Width / 2 - 2000
    .tmrSetFocus.Interval = 100
  End With
  StartBar True
  Mouse True
  PutFormOnTop
  DoCAD True
End Sub
Sub CheckStartupSettings()
  Dim AppPath As String
  Dim RegString As String
  If Right(App.Path, 1) <> "\" Then
    AppPath = App.Path & "\" & App.EXEName & ".exe"
  ElseIf Right(App.Path, 1) = "\" Then
    AppPath = App.Path & App.EXEName & ".exe"
  End If
  RegString = GetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "Blackout")
  If UCase$(RegString) <> UCase$(AppPath) Then
    Call SaveString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "Blackout", AppPath)
  End If
  
End Sub
