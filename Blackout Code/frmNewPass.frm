VERSION 5.00
Begin VB.Form frmNewPass 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Blackout - New Password"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3630
   Icon            =   "frmNewPass.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   3630
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   1560
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   1560
      Width           =   855
   End
   Begin VB.Frame fraNewPass 
      Caption         =   "&Enter New Password:"
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3255
      Begin VB.TextBox txtPassword2 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   360
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   840
         Width           =   2535
      End
      Begin VB.TextBox txtPassword1 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   360
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   480
         Width           =   2535
      End
   End
End
Attribute VB_Name = "frmNewPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdCancel_Click()
  frmOptions.Show
End Sub
Private Sub cmdOK_Click()
  Dim fn As Integer
  If Trim(txtPassword1.Text) = "" Or Trim(txtPassword2.Text) = "" Then
    MsgBox "You must enter a password, dumbass", vbOKOnly Or vbExclamation, "Blackout by Slayer"
    Exit Sub
  End If
  If UCase$(txtPassword1.Text) = UCase$(txtPassword2.Text) Then
    fn = FreeFile
    Open PasswordFileLocation For Output As #fn
      Print #fn, UCase$(txtPassword1.Text)
    Close #fn
    MsgBox "Your password has successfully been changed", vbOKOnly Or vbInformation, "Blackout by Slayer"
    Password = txtPassword1.Text
    StartupStuff
  ElseIf UCase$(txtPassword1.Text) <> UCase$(txtPassword2.Text) Then
    MsgBox "Your passwords did not match. Please re-type your passwords", vbOKOnly Or vbCritical, "Blackout by Slayer"
    txtPassword1.Text = ""
    txtPassword2.Text = ""
    txtPassword1.SetFocus
    Exit Sub
  End If
End Sub
Private Sub Form_Load()
  Dim Reply As Single
  Unload frmLockoutScreen
  Unload frmOptions
  Select Case NoPassFile
    Case True
      Reply = MsgBox("Blackout was unable to locate the password file. You must enter " & vbCr & " a new password. Proceed?", vbYesNo Or vbQuestion, "Blackout by Slayer")
      Select Case Reply
        Case vbNo
          Unload Me
          End
        Case Else
          cmdCancel.Enabled = False
      End Select
    Case False
      cmdCancel.Enabled = True
  End Select
End Sub
Private Sub Form_Unload(Cancel As Integer)
  Select Case NoPassFile
    Case True
      frmLockoutScreen.Show
  End Select
End Sub
Private Sub txtPassword1_GotFocus()
  With txtPassword1
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub
Private Sub txtPassword2_GotFocus()
  With txtPassword2
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub
