VERSION 5.00
Begin VB.Form frmLockoutScreen 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Lockout by Slayer"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   Icon            =   "frmLockoutScreen.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrEgg 
      Left            =   240
      Top             =   1440
   End
   Begin VB.Timer tmrSetFocus 
      Left            =   240
      Top             =   720
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1440
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   960
      Width           =   2775
   End
   Begin VB.Image imgEgg 
      Height          =   495
      Left            =   0
      Top             =   0
      Width           =   495
   End
End
Attribute VB_Name = "frmLockoutScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
  CheckForPasswordFile
  CheckStartupSettings
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Mouse False
End Sub

Private Sub imgEgg_Click()
  tmrEgg.Interval = 500
End Sub
Private Sub tmrEgg_Timer()
  If GetAsyncKeyState(VK_S) And GetAsyncKeyState(VK_L) And GetAsyncKeyState(VK_A) And GetAsyncKeyState(VK_Y) And GetAsyncKeyState(VK_E) And GetAsyncKeyState(VK_R) Then
    Mouse False
    StartBar False
    DoCAD False
    frmOptions.Show
  End If
End Sub
Private Sub tmrSetFocus_Timer()
  BringWindowToTop Me.hwnd
  txtPassword.SetFocus
  Mouse True
  StartBar True
  
End Sub
Private Sub txtPassword_Change()
  With txtPassword
    If UCase$(.Text) = Password Then
      tmrSetFocus.Enabled = False
      StartBar False
      Mouse False
      DoCAD False
      frmOptions.Show
    End If
  End With
End Sub
