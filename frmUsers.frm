VERSION 5.00
Begin VB.Form frmUsers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Player Config"
   ClientHeight    =   1995
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3915
   Icon            =   "frmUsers.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1995
   ScaleWidth      =   3915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox checkOption 
      Caption         =   "Ranged DMG"
      Height          =   240
      Index           =   9
      Left            =   135
      TabIndex        =   14
      Top             =   945
      Width           =   1320
   End
   Begin VB.CheckBox checkOption 
      Caption         =   "Comments"
      Height          =   240
      Index           =   8
      Left            =   1485
      TabIndex        =   13
      Top             =   1485
      Width           =   1140
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1395
      TabIndex        =   10
      Text            =   "Player1"
      Top             =   45
      Width           =   2445
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   330
      Left            =   2700
      TabIndex        =   9
      Top             =   1365
      Width           =   1140
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   330
      Left            =   2700
      TabIndex        =   8
      Top             =   1005
      Width           =   1140
   End
   Begin VB.CheckBox checkOption 
      Caption         =   "Accuracy"
      Height          =   240
      Index           =   7
      Left            =   1485
      TabIndex        =   7
      Top             =   1215
      Width           =   1410
   End
   Begin VB.CheckBox checkOption 
      Caption         =   "Hits/Misses"
      Height          =   240
      Index           =   6
      Left            =   1485
      TabIndex        =   6
      Top             =   945
      Width           =   1410
   End
   Begin VB.CheckBox checkOption 
      Caption         =   "Critical Hits"
      Height          =   240
      Index           =   5
      Left            =   1485
      TabIndex        =   5
      Top             =   675
      Width           =   1410
   End
   Begin VB.CheckBox checkOption 
      Caption         =   "Spell DMG"
      Height          =   240
      Index           =   4
      Left            =   1485
      TabIndex        =   4
      Top             =   405
      Width           =   1410
   End
   Begin VB.CheckBox checkOption 
      Caption         =   "Skill DMG"
      Height          =   240
      Index           =   3
      Left            =   135
      TabIndex        =   3
      Top             =   1485
      Width           =   1410
   End
   Begin VB.CheckBox checkOption 
      Caption         =   "Critical DMG"
      Height          =   240
      Index           =   2
      Left            =   135
      TabIndex        =   2
      Top             =   1215
      Width           =   1410
   End
   Begin VB.CheckBox checkOption 
      Caption         =   "Melee DMG"
      Height          =   240
      Index           =   1
      Left            =   135
      TabIndex        =   1
      Top             =   675
      Width           =   1410
   End
   Begin VB.CheckBox checkOption 
      Caption         =   "Total DMG"
      Height          =   240
      Index           =   0
      Left            =   135
      TabIndex        =   0
      Top             =   405
      Width           =   1410
   End
   Begin VB.Label lblNote 
      Alignment       =   2  'Center
      Caption         =   "Note: Changing options will clear current results!"
      Height          =   240
      Left            =   45
      TabIndex        =   12
      Top             =   1755
      Width           =   3795
   End
   Begin VB.Label Label1 
      Caption         =   "Player Name:"
      Height          =   240
      Left            =   135
      TabIndex        =   11
      Top             =   90
      Width           =   1320
   End
End
Attribute VB_Name = "frmUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Sub cmdCancel_Click()
Unload Me
Set frmUsers = Nothing
End Sub


Private Sub cmdOK_Click()
frmRead.comboUser.List(frmRead.comboUser.ListIndex) = txtName.Text

UserLog(frmRead.comboUser.ListIndex, 1) = ""
UserLog(frmRead.comboUser.ListIndex, 2) = "0"
UserLog(frmRead.comboUser.ListIndex, 3) = "0"
UserLog(frmRead.comboUser.ListIndex, 4) = "0"
UserLog(frmRead.comboUser.ListIndex, 5) = "0"
UserLog(frmRead.comboUser.ListIndex, 6) = "0"
UserLog(frmRead.comboUser.ListIndex, 8) = "0"
UserLog(frmRead.comboUser.ListIndex, 9) = "0"
UserLog(frmRead.comboUser.ListIndex, 10) = "0"
UserLog(frmRead.comboUser.ListIndex, 11) = "0"
UserLog(frmRead.comboUser.ListIndex, 12) = "0"
UserLog(frmRead.comboUser.ListIndex, 13) = "0"
    
UserLog(frmRead.comboUser.ListIndex, 0) = frmRead.comboUser.Text
UserLog(frmRead.comboUser.ListIndex, 7) = checkOption(0).Value & "," & checkOption(1).Value & "," & checkOption(2).Value & "," & checkOption(3).Value & "," & checkOption(4).Value & "," & checkOption(5).Value & "," & checkOption(6).Value & "," & checkOption(7).Value & "," & checkOption(8).Value & "," & checkOption(9).Value
SaveSetting App.Title, "Settings", "Player" & frmRead.comboUser.ListIndex + 1, frmRead.comboUser.Text
SaveSetting App.Title, "Settings", "PlayerOptions" & frmRead.comboUser.ListIndex + 1, UserLog(frmRead.comboUser.ListIndex, 7)
frmRead.comboUser_Click
Dim i As Integer
For i = 0 To frmRead.comboUser.ListCount - 1
    frmRead.mnuViewPlayer(i).Caption = "&" & i & ". " & frmRead.comboUser.List(i)
Next
Unload Me
Set frmUsers = Nothing
End Sub


Private Sub Form_Load()
frmRead.Enabled = False
Me.Left = frmRead.Left + 100
Me.Top = frmRead.Top + 100
txtName.Text = UserLog(frmRead.comboUser.ListIndex, 0)
checkOption(0).Value = Mid$(UserLog(frmRead.comboUser.ListIndex, 7), 1, 1)
checkOption(1).Value = Mid$(UserLog(frmRead.comboUser.ListIndex, 7), 3, 1)
checkOption(2).Value = Mid$(UserLog(frmRead.comboUser.ListIndex, 7), 5, 1)
checkOption(3).Value = Mid$(UserLog(frmRead.comboUser.ListIndex, 7), 7, 1)
checkOption(4).Value = Mid$(UserLog(frmRead.comboUser.ListIndex, 7), 9, 1)
checkOption(5).Value = Mid$(UserLog(frmRead.comboUser.ListIndex, 7), 11, 1)
checkOption(6).Value = Mid$(UserLog(frmRead.comboUser.ListIndex, 7), 13, 1)
checkOption(7).Value = Mid$(UserLog(frmRead.comboUser.ListIndex, 7), 15, 1)
checkOption(8).Value = Mid$(UserLog(frmRead.comboUser.ListIndex, 7), 17, 1)
checkOption(9).Value = Mid$(UserLog(frmRead.comboUser.ListIndex, 7), 19, 1)
End Sub


Private Sub Form_Unload(Cancel As Integer)
frmRead.Enabled = True
End Sub


