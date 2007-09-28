VERSION 5.00
Begin VB.Form frmSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Setup"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5550
   Icon            =   "frmSetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   5550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add Player"
      Height          =   315
      Left            =   3120
      TabIndex        =   35
      Top             =   1500
      Width           =   2295
   End
   Begin VB.Frame Frame2 
      Caption         =   "Crafting Skill Levels"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1395
      Left            =   60
      TabIndex        =   16
      Top             =   2220
      Width           =   5415
      Begin VB.TextBox txtCraft 
         Height          =   285
         Index           =   8
         Left            =   3765
         TabIndex        =   33
         Text            =   "0"
         Top             =   900
         Width           =   375
      End
      Begin VB.TextBox txtCraft 
         Height          =   285
         Index           =   7
         Left            =   3765
         TabIndex        =   31
         Text            =   "0"
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox txtCraft 
         Height          =   285
         Index           =   6
         Left            =   3765
         TabIndex        =   29
         Text            =   "0"
         Top             =   300
         Width           =   375
      End
      Begin VB.TextBox txtCraft 
         Height          =   285
         Index           =   5
         Left            =   1958
         TabIndex        =   27
         Text            =   "0"
         Top             =   900
         Width           =   375
      End
      Begin VB.TextBox txtCraft 
         Height          =   285
         Index           =   4
         Left            =   1958
         TabIndex        =   25
         Text            =   "0"
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox txtCraft 
         Height          =   285
         Index           =   3
         Left            =   1958
         TabIndex        =   23
         Text            =   "0"
         Top             =   300
         Width           =   375
      End
      Begin VB.TextBox txtCraft 
         Height          =   285
         Index           =   2
         Left            =   225
         TabIndex        =   21
         Text            =   "0"
         Top             =   900
         Width           =   375
      End
      Begin VB.TextBox txtCraft 
         Height          =   285
         Index           =   1
         Left            =   225
         TabIndex        =   19
         Text            =   "0"
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox txtCraft 
         Height          =   285
         Index           =   0
         Left            =   225
         TabIndex        =   17
         Text            =   "0"
         Top             =   300
         Width           =   375
      End
      Begin VB.Label lblCraft 
         Caption         =   "Woodworking"
         Height          =   195
         Index           =   8
         Left            =   4245
         TabIndex        =   34
         Top             =   960
         Width           =   1035
      End
      Begin VB.Label lblCraft 
         Caption         =   "Smithing"
         Height          =   195
         Index           =   7
         Left            =   4245
         TabIndex        =   32
         Top             =   660
         Width           =   675
      End
      Begin VB.Label lblCraft 
         Caption         =   "Leathercraft"
         Height          =   195
         Index           =   6
         Left            =   4245
         TabIndex        =   30
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblCraft 
         Caption         =   "Goldsmithing"
         Height          =   195
         Index           =   5
         Left            =   2438
         TabIndex        =   28
         Top             =   960
         Width           =   975
      End
      Begin VB.Label lblCraft 
         Caption         =   "Fishing"
         Height          =   195
         Index           =   4
         Left            =   2438
         TabIndex        =   26
         Top             =   660
         Width           =   555
      End
      Begin VB.Label lblCraft 
         Caption         =   "Cooking"
         Height          =   195
         Index           =   3
         Left            =   2438
         TabIndex        =   24
         Top             =   360
         Width           =   675
      End
      Begin VB.Label lblCraft 
         Caption         =   "Clothcraft"
         Height          =   195
         Index           =   2
         Left            =   705
         TabIndex        =   22
         Top             =   960
         Width           =   735
      End
      Begin VB.Label lblCraft 
         Caption         =   "Bonecraft"
         Height          =   195
         Index           =   1
         Left            =   705
         TabIndex        =   20
         Top             =   660
         Width           =   735
      End
      Begin VB.Label lblCraft 
         Caption         =   "Alchemy"
         Height          =   195
         Index           =   0
         Left            =   705
         TabIndex        =   18
         Top             =   360
         Width           =   675
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   315
      Left            =   4320
      TabIndex        =   7
      Top             =   1860
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   315
      Left            =   3120
      TabIndex        =   6
      Top             =   1860
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Character Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2115
      Left            =   60
      TabIndex        =   8
      Top             =   60
      Width           =   3015
      Begin VB.ComboBox comboServer 
         Height          =   315
         ItemData        =   "frmSetup.frx":000C
         Left            =   840
         List            =   "frmSetup.frx":0070
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1740
         Width           =   2055
      End
      Begin VB.TextBox txtLevel 
         Height          =   315
         Left            =   840
         TabIndex        =   4
         Top             =   1440
         Width           =   2055
      End
      Begin VB.ComboBox comboSub 
         Height          =   315
         ItemData        =   "frmSetup.frx":0195
         Left            =   840
         List            =   "frmSetup.frx":01C6
         Sorted          =   -1  'True
         TabIndex        =   3
         Text            =   "comboSub"
         Top             =   1140
         Width           =   2055
      End
      Begin VB.ComboBox comboJob 
         Height          =   315
         ItemData        =   "frmSetup.frx":0257
         Left            =   840
         List            =   "frmSetup.frx":0288
         Sorted          =   -1  'True
         TabIndex        =   2
         Text            =   "comboJob"
         Top             =   840
         Width           =   2055
      End
      Begin VB.TextBox txtLS 
         Height          =   315
         Left            =   840
         TabIndex        =   1
         Top             =   540
         Width           =   2055
      End
      Begin VB.ComboBox comboPlayer 
         Height          =   315
         ItemData        =   "frmSetup.frx":0319
         Left            =   840
         List            =   "frmSetup.frx":031B
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "LinkShell:"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   15
         Top             =   570
         Width           =   675
      End
      Begin VB.Label Label1 
         Caption         =   "Server:"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   13
         Top             =   1770
         Width           =   675
      End
      Begin VB.Label Label1 
         Caption         =   "Level:"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   12
         Top             =   1470
         Width           =   675
      End
      Begin VB.Label Label1 
         Caption         =   "Sub-Job:"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   11
         Top             =   1170
         Width           =   675
      End
      Begin VB.Label Label1 
         Caption         =   "Job:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   870
         Width           =   675
      End
      Begin VB.Label Label1 
         Caption         =   "Player:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   0
         Left            =   120
         MouseIcon       =   "frmSetup.frx":031D
         MousePointer    =   99  'Custom
         TabIndex        =   9
         Top             =   270
         Width           =   675
      End
   End
   Begin VB.Label Label2 
      Caption         =   $"frmSetup.frx":046F
      ForeColor       =   &H00FF0000&
      Height          =   855
      Left            =   3180
      TabIndex        =   14
      Top             =   420
      Width           =   2355
   End
End
Attribute VB_Name = "frmSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
Label1_Click 0
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
Dim i, o, UserName As String, FoundPlayer As Boolean, PlayerCount As Long
PlayerCount = GetSetting(App.Title, "Online_Setup", "Count", "5")
For i = 0 To PlayerCount
    UserName = GetSetting(App.Title, "Online_Setup", "Name" & i, "")
    If UserName = comboPlayer.Text Then
        SaveSetting App.Title, "Online_Setup", "LS" & i, txtLS.Text
        SaveSetting App.Title, "Online_Setup", "Job" & i, comboJob.Text
        SaveSetting App.Title, "Online_Setup", "SubJob" & i, comboSub.Text
        SaveSetting App.Title, "Online_Setup", "Level" & i, txtLevel.Text
        SaveSetting App.Title, "Online_Setup", "Server" & i, comboServer.Text
        For o = 0 To txtCraft.UBound
            SaveSetting App.Title, "Online_Setup", "Skill" & i & "-" & o, txtCraft(o).Text
        Next
        FoundPlayer = True
        Exit For
    Else
        FoundPlayer = False
    End If
Next

If FoundPlayer = False Then
    For i = 0 To PlayerCount
        UserName = GetSetting(App.Title, "Online_Setup", "Name" & i, "")
        If UserName = "" Then
            SaveSetting App.Title, "Online_Setup", "Name" & i, comboPlayer.Text
            SaveSetting App.Title, "Online_Setup", "LS" & i, txtLS.Text
            SaveSetting App.Title, "Online_Setup", "Job" & i, comboJob.Text
            SaveSetting App.Title, "Online_Setup", "SubJob" & i, comboSub.Text
            SaveSetting App.Title, "Online_Setup", "Level" & i, txtLevel.Text
            SaveSetting App.Title, "Online_Setup", "Server" & i, comboServer.Text
            For o = 0 To txtCraft.UBound
                SaveSetting App.Title, "Online_Setup", "Skill" & i & "-" & o, txtCraft(o).Text
            Next
            FoundPlayer = True
            Exit For
        Else
            FoundPlayer = False
        End If
    Next
End If

If FoundPlayer = False Then
    SaveSetting App.Title, "Online_Setup", "Count", PlayerCount + 1
    SaveSetting App.Title, "Online_Setup", "Name" & PlayerCount + 1, comboPlayer.Text
    SaveSetting App.Title, "Online_Setup", "LS" & PlayerCount + 1, txtLS.Text
    SaveSetting App.Title, "Online_Setup", "Job" & PlayerCount + 1, comboJob.Text
    SaveSetting App.Title, "Online_Setup", "SubJob" & PlayerCount + 1, comboSub.Text
    SaveSetting App.Title, "Online_Setup", "Level" & PlayerCount + 1, txtLevel.Text
    SaveSetting App.Title, "Online_Setup", "Server" & PlayerCount + 1, comboServer.Text
    For o = 0 To txtCraft.UBound
        SaveSetting App.Title, "Online_Setup", "Skill" & PlayerCount + 1 & "-" & o, txtCraft(o).Text
    Next
End If
End Sub


Private Sub comboPlayer_Click()
Dim i, o, UserName As String, PlayerCount As Long
Dim Job, SubJob, Server
PlayerCount = GetSetting(App.Title, "Online_Setup", "Count", "5")
For i = 0 To PlayerCount
    UserName = GetSetting(App.Title, "Online_Setup", "Name" & i, "")
    If UserName = comboPlayer.Text Then
        Job = GetSetting(App.Title, "Online_Setup", "Job" & i, "")
        If Job <> "" Then
            comboJob.Text = GetSetting(App.Title, "Online_Setup", "Job" & i, "")
        End If
        SubJob = GetSetting(App.Title, "Online_Setup", "SubJob" & i, "")
        If SubJob <> "" Then
            comboSub.Text = GetSetting(App.Title, "Online_Setup", "SubJob" & i, "")
        End If
        txtLevel.Text = GetSetting(App.Title, "Online_Setup", "Level" & i, "")

        Server = GetSetting(App.Title, "Online_Setup", "Server" & i, "")
        If Server <> "" Then
           comboServer.Text = GetSetting(App.Title, "Online_Setup", "Server" & i, "")
        End If
        txtLS.Text = GetSetting(App.Title, "Online_Setup", "LS" & i, "")
        For o = 0 To txtCraft.UBound
           txtCraft(o).Text = GetSetting(App.Title, "Online_Setup", "Skill" & i & "-" & o, "0")
        Next
        Exit For
    End If
Next
End Sub


Private Sub Form_Load()
On Error Resume Next
If frmRead.listPlayers.ListCount <> 0 Then
    For i = 0 To frmRead.listPlayers.ListCount
        If frmRead.listPlayers.List(i) <> "" Then
            comboPlayer.AddItem frmRead.listPlayers.List(i)
        End If
    Next
Else
    PlayerCount = GetSetting(App.Title, "Online_Setup", "Count", "5")
    For i = 0 To PlayerCount
        UserName = GetSetting(App.Title, "Online_Setup", "Name" & i, "")
        If UserName <> "" Then
            comboPlayer.AddItem UserName
        End If
    Next
End If
Me.Left = frmRead.Left + 100
Me.Top = frmRead.Top + 100
If comboPlayer.ListCount <> 0 Then
    comboPlayer.ListIndex = 0
End If
comboPlayer_Click
End Sub


Private Sub Label1_Click(Index As Integer)
Dim AddPlayer As String, FoundIt As Boolean
AddPlayer = InputBox("Player Name:", "Add Player")
If AddPlayer <> "" And AddPlayer <> "0" Then
    For i = 0 To comboPlayer.ListCount
        If LCase(Trim(comboPlayer.List(i))) = LCase(Trim(AddPlayer)) Then FoundIt = True
    Next
    If FoundIt Then
        MsgBox "Player already added.", vbInformation, "Add Player"
    Else
        comboPlayer.AddItem AddPlayer
        MsgBox "Player added.", vbInformation, "Add Player"
    End If
End If
frmSetup.SetFocus
End Sub


Private Sub txtCraft_Validate(Index As Integer, Cancel As Boolean)
If IsNumeric(txtCraft(Index)) = False Then
    txtCraft(Index).Text = "0"
End If
End Sub


