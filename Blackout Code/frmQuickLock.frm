VERSION 5.00
Begin VB.Form frmQuickLock 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Blackout by Slayer - Quick Lock"
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4560
   Icon            =   "frmQuickLock.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   4560
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraQuickLock 
      Caption         =   "&Quick Lock Settings"
      Height          =   1335
      Left            =   173
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      Begin VB.CommandButton cmdOK 
         Caption         =   "O&k"
         Default         =   -1  'True
         Height          =   375
         Left            =   3240
         TabIndex        =   3
         Top             =   840
         Width           =   855
      End
      Begin VB.CommandButton cmdQuestion 
         Caption         =   "&What is Quick Lock?"
         Height          =   375
         Left            =   2280
         TabIndex        =   2
         Top             =   300
         Width           =   1815
      End
      Begin VB.CheckBox chkQuickLock 
         Caption         =   "E&nable Quick Lock"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1935
      End
   End
End
Attribute VB_Name = "frmQuickLock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdOK_Click()
  If chkQuickLock.Value = 1 Then
    Call SaveString(HKEY_LOCAL_MACHINE, "Software\Blackout", "EnableQuickLock", "True")
  ElseIf chkQuickLock.Value = 0 Then
    Call SaveString(HKEY_LOCAL_MACHINE, "Software\Blackout", "EnableQuickLock", "False")
  End If
  frmOptions.Show
End Sub
Private Sub cmdQuestion_Click()
  MsgBox "Quick Lock is a function that executes Blackout when the user pushes " & vbCr & "Control + F12 at the same time", vbOKOnly Or vbInformation, "Blackout by Slayer"
End Sub
Private Sub Form_Load()
  Unload frmOptions
  Dim StringVal As String
  StringVal = GetString(HKEY_LOCAL_MACHINE, "Software\Blackout", "EnableQuickLock")
  If UCase$(StringVal) = "TRUE" Then
    chkQuickLock.Value = 1
  ElseIf UCase$(StringVal) = "FALSE" Then
    chkQuickLock.Value = 0
  End If
End Sub
