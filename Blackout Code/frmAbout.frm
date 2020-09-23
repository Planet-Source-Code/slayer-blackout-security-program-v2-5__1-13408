VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Blackout by Slayer - About"
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3300
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   3300
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image1 
      Height          =   255
      Left            =   3120
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label lblVersion 
      Caption         =   "&Version Info:  v2.0"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label lblCompiled 
      Caption         =   "&Compiled:  November 22, 2000"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label lblLanguage 
      Caption         =   "&Language:  Visual Basic 6.0"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label lblAbout 
      Caption         =   "Blackout was coded by Dustin Leslie"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   2895
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
  Unload frmOptions
End Sub
Private Sub Form_Unload(Cancel As Integer)
  frmOptions.Show
End Sub

Private Sub Image1_Click()
  MsgBox "Bah!! Leave this program alone, quit buggin it!!" & vbCr & vbCr & vbTab & vbTab & " -Slayer", vbOKOnly Or vbExclamation, "Blackout by Slayer"
End Sub
