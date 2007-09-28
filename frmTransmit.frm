VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmTransmit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Transmit"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6015
   Icon            =   "frmTransmit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Send Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   60
      TabIndex        =   13
      Top             =   60
      Width           =   2295
      Begin VB.OptionButton optionSend 
         Caption         =   "Crafting Data"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   22
         Top             =   960
         Width           =   1815
      End
      Begin VB.ComboBox comboMOB 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmTransmit.frx":000C
         Left            =   120
         List            =   "frmTransmit.frx":000E
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   1440
         Width           =   2055
      End
      Begin VB.ComboBox comboPlayer 
         Height          =   315
         ItemData        =   "frmTransmit.frx":0010
         Left            =   120
         List            =   "frmTransmit.frx":0012
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   1980
         Width           =   2055
      End
      Begin VB.OptionButton optionSend 
         Caption         =   "Highest Hits Only"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Width           =   1815
      End
      Begin VB.OptionButton optionSend 
         Caption         =   "Selected MOB Type"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   480
         Width           =   1815
      End
      Begin VB.OptionButton optionSend 
         Caption         =   "All MOB Types"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Send stats for MOB type:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   1230
         Width           =   1995
      End
      Begin VB.Label Label1 
         Caption         =   "Send stats for player:"
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
         Height          =   195
         Left            =   120
         MouseIcon       =   "frmTransmit.frx":0014
         MousePointer    =   99  'Custom
         TabIndex        =   18
         Top             =   1770
         Width           =   1575
      End
   End
   Begin InetCtlsObjects.Inet inet 
      Left            =   3960
      Top             =   2460
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   315
      Left            =   4860
      TabIndex        =   10
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send"
      Height          =   315
      Left            =   2400
      TabIndex        =   9
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   2400
      TabIndex        =   0
      Top             =   60
      Width           =   3555
      Begin VB.TextBox txtStatus 
         BackColor       =   &H00FF0000&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   0
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Text            =   "Waiting..."
         Top             =   240
         Width           =   1155
      End
      Begin VB.TextBox txtStatus 
         BackColor       =   &H00FF0000&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   4
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   "Waiting..."
         Top             =   1440
         Width           =   1155
      End
      Begin VB.TextBox txtStatus 
         BackColor       =   &H00FF0000&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   3
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Text            =   "Waiting..."
         Top             =   1140
         Width           =   1155
      End
      Begin VB.TextBox txtStatus 
         BackColor       =   &H00FF0000&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   2
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Text            =   "Waiting..."
         Top             =   840
         Width           =   1155
      End
      Begin VB.TextBox txtStatus 
         BackColor       =   &H00FF0000&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   1
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   "Waiting..."
         Top             =   540
         Width           =   1155
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   21
         Top             =   1800
         Width           =   3315
      End
      Begin VB.Label lblStatus 
         Caption         =   "Scanning Statistics:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   255
         Width           =   2175
      End
      Begin VB.Label lblStatus 
         Caption         =   "Magic Damage Data:"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   4
         Top             =   1455
         Width           =   2175
      End
      Begin VB.Label lblStatus 
         Caption         =   "Ranged Damage Data:"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   3
         Top             =   1155
         Width           =   2175
      End
      Begin VB.Label lblStatus 
         Caption         =   "Weapon Skill Damage Data:"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   2
         Top             =   855
         Width           =   2175
      End
      Begin VB.Label lblStatus 
         Caption         =   "Basic Damage Data:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Top             =   555
         Width           =   2175
      End
   End
End
Attribute VB_Name = "frmTransmit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub checkOption_Click()
If checkOption.Value = 1 Then
    If frmRead.listResults.SelCount > 1 Then
        MsgBox "Please select only 1 MOB from the Results list and try again.", vbInformation, "Error"
        checkOption.Value = 0
    End If
End If
End Sub

Private Sub SendCraft()
Screen.MousePointer = vbHourglass
Dim UserName As String, Server As String, Level As String, Job As String, SubJob As String, ls
Dim CurrentType As String, HighHit As String
Dim URL_Part1 As String
Dim URL_Part2 As String
Dim Result As String, Include As Boolean, FoundPlayer As Boolean, PlayerCount As Long

Dim CurrentBasicHigh As Long, CurrentBasicMob As String, CurrentBasicAcc As String
Dim CurrentSkillHigh As Long, CurrentSkillMob As String, CurrentSkill As String
Dim CurrentRangedHigh As Long, CurrentRangedMob As String, CurrentRangedAcc As String
Dim CurrentSpellHigh As Long, CurrentSpellMob As String, CurrentSpell As String

Dim myCraft(8) As String



PlayerCount = GetSetting(App.Title, "Online_Setup", "Count", "5")
For i = 0 To PlayerCount
    UserName = GetSetting(App.Title, "Online_Setup", "Name" & i, "")
    If UserName = comboPlayer.Text Then
        UserName = GetSetting(App.Title, "Online_Setup", "Name" & i, "")
        Job = GetSetting(App.Title, "Online_Setup", "Job" & i, "")
        SubJob = GetSetting(App.Title, "Online_Setup", "SubJob" & i, "")
        Level = GetSetting(App.Title, "Online_Setup", "Level" & i, "")
        Server = GetSetting(App.Title, "Online_Setup", "Server" & i, "")
        ls = GetSetting(App.Title, "Online_Setup", "LS" & i, "")
        For o = 0 To UBound(myCraft)
           myCraft(o) = GetSetting(App.Title, "Online_Setup", "Skill" & i & "-" & o, "0")
        Next
        FoundPlayer = True
        Exit For
    Else
        FoundPlayer = False
    End If
Next
If FoundPlayer = False Then
    MsgBox "Player is not configured.", vbInformation, "Send"
    Screen.MousePointer = vbDefault
    Exit Sub
End If

URL_Part1 = "http://ffxi.mmorpgparsers.com/craft/update_craft.php?user=" & UserName & "&server=" & Server & "&level=" & Level & "&job=" & Job & "&subjob=" & SubJob & "&ls=" & ls & "&alchemy=" & myCraft(0) & "&bone=" & myCraft(1) & "&cloth=" & myCraft(2) & "&cook=" & myCraft(3) & "&fish=" & myCraft(4) & "&gold=" & myCraft(5) & "&leather=" & myCraft(6) & "&smith=" & myCraft(7) & "&wood=" & myCraft(8)
'alchemy=1&bone=2&cloth=3&cook=4&fish=5&gold=6&leather=7&smith=8&wood=9&moon=New Moon&day=Lightsday&perc=5%&hour=8:31&result=a flask of holy water&count=1

    txtStatus(1).Text = "Skipping."
    txtStatus(1).BackColor = &HC000C0
    txtStatus(2).Text = "Skipping."
    txtStatus(2).BackColor = &HC000C0
    txtStatus(3).Text = "Skipping."
    txtStatus(3).BackColor = &HC000C0
    txtStatus(4).Text = "Skipping."
    txtStatus(4).BackColor = &HC000C0
    
    txtStatus(0).Text = "Sending..."
    txtStatus(0).BackColor = &HFF&
    For i = 0 To UBound(CraftingCSV)
        lblStatus(5).Caption = "Sending " & i + 1 & " of " & UBound(CraftingCSV) + 1
        URL_Part2 = "&moon=" & CraftingCSV(i).MoonPhase
        URL_Part2 = URL_Part2 & "&day=" & CraftingCSV(i).DayType
        URL_Part2 = URL_Part2 & "&perc=" & CraftingCSV(i).MoonPerc
        URL_Part2 = URL_Part2 & "&hour=" & CraftingCSV(i).CurrentTime
        URL_Part2 = URL_Part2 & "&result=" & CraftingCSV(i).Result
        URL_Part2 = URL_Part2 & "&count=" & CraftingCSV(i).Count
        URL_Part2 = URL_Part2 & "&direction=" & CraftingCSV(i).Direction
        Result = inet.OpenURL(URL_Part1 & URL_Part2)
        If Result = "error" Then
            txtStatus(0).Text = "Failed."
            txtStatus(0).BackColor = vbBlack
            Exit For
        Else
        End If
    Next
    txtStatus(0).Text = "Done."
    txtStatus(0).BackColor = &HC000&
    

MsgBox "Finished!", vbInformation, "Send"
Screen.MousePointer = vbDefault
frmTransmit.SetFocus
End Sub

Private Sub SendMOB()
Screen.MousePointer = vbHourglass
Dim UserName As String, Server As String, Level As String, Job As String, SubJob As String, ls
Dim CurrentType As String, HighHit As String
Dim URL_Part1 As String
Dim URL_Part2 As String
Dim Result As String, Include As Boolean, FoundPlayer As Boolean, PlayerCount As Long

Dim CurrentBasicHigh As Long, CurrentBasicMob As String, CurrentBasicAcc As String
Dim CurrentSkillHigh As Long, CurrentSkillMob As String, CurrentSkill As String
Dim CurrentRangedHigh As Long, CurrentRangedMob As String, CurrentRangedAcc As String
Dim CurrentSpellHigh As Long, CurrentSpellMob As String, CurrentSpell As String

PlayerCount = GetSetting(App.Title, "Online_Setup", "Count", "5")
For i = 0 To PlayerCount
    UserName = GetSetting(App.Title, "Online_Setup", "Name" & i, "")
    If UserName = comboPlayer.Text Then
        UserName = GetSetting(App.Title, "Online_Setup", "Name" & i, "")
        Job = GetSetting(App.Title, "Online_Setup", "Job" & i, "")
        SubJob = GetSetting(App.Title, "Online_Setup", "SubJob" & i, "")
        Level = GetSetting(App.Title, "Online_Setup", "Level" & i, "")
        Server = GetSetting(App.Title, "Online_Setup", "Server" & i, "")
        ls = GetSetting(App.Title, "Online_Setup", "LS" & i, "")
        FoundPlayer = True
        Exit For
    Else
        FoundPlayer = False
    End If
Next
If FoundPlayer = False Then
    MsgBox "Player is not configured.", vbInformation, "Send"
    Screen.MousePointer = vbDefault
    Exit Sub
End If

URL_Part1 = "http://ffxi.mmorpgparsers.com/update_data.php?user=" & UserName & "&server=" & Server & "&level=" & Level & "&job=" & Job & "&subjob=" & SubJob & "&ls=" & ls

If ReportError <> "" Then
    Result = inet.OpenURL("http://ffxi.mmorpgparsers.com/update_error.php?error_text=" & ReportError & "&version=" & App.Major & "." & App.Minor & "." & App.Revision)
    ReportError = ""
End If


If optionSend(0).Value = True Then
    For o = 0 To comboMOB.ListCount - 1
        lblStatus(5).Caption = "Sending: " & comboMOB.List(o) & " (" & o + 1 & " of " & comboMOB.ListCount & ")"
        For i = 0 To txtStatus.UBound
            txtStatus(i).Text = "Waiting..."
            txtStatus(i).BackColor = &HFF0000
        Next
        txtStatus(0).BackColor = &HFF&
        For i = 0 To UBound(FullStats)
            txtStatus(0).Text = i & " of " & UBound(FullStats)
            DoEvents
            If FullStats(i).Attacker = UserName And FullStats(i).Defender = comboMOB.List(o) Then
                If FullStats(i).Basic.High > CurrentBasicHigh Then
                    CurrentBasicHigh = FullStats(i).Basic.High
                    CurrentBasicMob = FullStats(i).Defender
                    CurrentBasicAcc = Round((FullStats(i).Basic.Hit / (FullStats(i).Basic.Hit + FullStats(i).Basic.Miss)) * 100, 2)
                End If
                If FullStats(i).Skill.High > CurrentSkillHigh Then
                    CurrentSkillHigh = FullStats(i).Skill.High
                    CurrentSkillMob = FullStats(i).Defender
                    CurrentSkill = FullStats(i).Skill.HighSkillType
                End If
                If FullStats(i).Ranged.High > CurrentRangedHigh Then
                    CurrentRangedHigh = FullStats(i).Ranged.High
                    CurrentRangedMob = FullStats(i).Defender
                    CurrentRangedAcc = Round((FullStats(i).Ranged.Hit / (FullStats(i).Ranged.Hit + FullStats(i).Ranged.Miss)) * 100, 2)
                End If
                If FullStats(i).Spell.High > CurrentSpellHigh Then
                    CurrentSpellHigh = FullStats(i).Spell.High
                    CurrentSpellMob = FullStats(i).Defender
                    CurrentSpell = FullStats(i).Spell.HighSkillType
                End If
            End If
        Next
        txtStatus(0).Text = "Done."
        txtStatus(0).BackColor = &HC000&
        
        txtStatus(1).BackColor = &HFF&
        txtStatus(1).Text = "Sending..."
        If CurrentBasicHigh <> 0 Then
            URL_Part2 = "&type=Basic&high=" & CurrentBasicHigh & "&mob=" & CurrentBasicMob & "&acc=" & CurrentBasicAcc
            Result = inet.OpenURL(URL_Part1 & URL_Part2)
            If Result = "error" Then
                txtStatus(1).Text = "Failed."
                txtStatus(1).BackColor = vbBlack
            Else
                txtStatus(1).Text = "Done."
                txtStatus(1).BackColor = &HC000&
            End If
        Else
            txtStatus(1).Text = "Skipped."
            txtStatus(1).BackColor = &HC000C0
        End If
        DoEvents
        
        txtStatus(2).BackColor = &HFF&
        txtStatus(2).Text = "Sending..."
        If CurrentSkillHigh <> 0 Then
            URL_Part2 = "&type=Skill&high=" & CurrentSkillHigh & "&mob=" & CurrentSkillMob & "&skilltype=" & CurrentSkill
            Result = inet.OpenURL(URL_Part1 & URL_Part2)
            If Result = "error" Then
                txtStatus(2).Text = "Failed."
                txtStatus(2).BackColor = vbBlack
            Else
                txtStatus(2).Text = "Done."
                txtStatus(2).BackColor = &HC000&
            End If
        Else
            txtStatus(2).Text = "Skipped."
            txtStatus(2).BackColor = &HC000C0
        End If
        
        txtStatus(3).BackColor = &HFF&
        txtStatus(3).Text = "Sending..."
        If CurrentRangedHigh <> 0 Then
            URL_Part2 = "&type=Ranged&high=" & CurrentRangedHigh & "&mob=" & CurrentRangedMob & "&acc=" & CurrentRangedAcc
            Result = inet.OpenURL(URL_Part1 & URL_Part2)
            If Result = "error" Then
                txtStatus(3).Text = "Failed."
                txtStatus(3).BackColor = vbBlack
            Else
                txtStatus(3).Text = "Done."
                txtStatus(3).BackColor = &HC000&
            End If
        Else
            txtStatus(3).Text = "Skipped."
            txtStatus(3).BackColor = &HC000C0
        End If
        
        txtStatus(4).BackColor = &HFF&
        txtStatus(4).Text = "Sending..."
        If CurrentSpellHigh <> 0 Then
            URL_Part2 = "&type=Spell&high=" & CurrentSpellHigh & "&mob=" & CurrentSpellMob & "&skilltype=" & CurrentSpell
            Result = inet.OpenURL(URL_Part1 & URL_Part2)
            If Result = "error" Then
                txtStatus(4).Text = "Failed."
                txtStatus(4).BackColor = vbBlack
            Else
                txtStatus(4).Text = "Done."
                txtStatus(4).BackColor = &HC000&
            End If
        Else
            txtStatus(4).Text = "Skipped."
            txtStatus(4).BackColor = &HC000C0
        End If
    Next
ElseIf optionSend(1).Value = True Then
    lblStatus(5).Caption = "Sending: " & comboMOB.Text
    For i = 0 To txtStatus.UBound
        txtStatus(i).Text = "Waiting..."
        txtStatus(i).BackColor = &HFF0000
    Next
    txtStatus(0).BackColor = &HFF&
    For i = 0 To UBound(FullStats)
        txtStatus(0).Text = i & " of " & UBound(FullStats)
        DoEvents
        If FullStats(i).Attacker = UserName And FullStats(i).Defender = comboMOB.Text Then
            If FullStats(i).Basic.High > CurrentBasicHigh Then
                CurrentBasicHigh = FullStats(i).Basic.High
                CurrentBasicMob = FullStats(i).Defender
                CurrentBasicAcc = Round((FullStats(i).Basic.Hit / (FullStats(i).Basic.Hit + FullStats(i).Basic.Miss)) * 100, 2)
            End If
            If FullStats(i).Skill.High > CurrentSkillHigh Then
                CurrentSkillHigh = FullStats(i).Skill.High
                CurrentSkillMob = FullStats(i).Defender
                CurrentSkill = FullStats(i).Skill.HighSkillType
            End If
            If FullStats(i).Ranged.High > CurrentRangedHigh Then
                CurrentRangedHigh = FullStats(i).Ranged.High
                CurrentRangedMob = FullStats(i).Defender
                CurrentRangedAcc = Round((FullStats(i).Ranged.Hit / (FullStats(i).Ranged.Hit + FullStats(i).Ranged.Miss)) * 100, 2)
            End If
            If FullStats(i).Spell.High > CurrentSpellHigh Then
                CurrentSpellHigh = FullStats(i).Spell.High
                CurrentSpellMob = FullStats(i).Defender
                CurrentSpell = FullStats(i).Spell.HighSkillType
            End If
        End If
    Next
    txtStatus(0).Text = "Done."
    txtStatus(0).BackColor = &HC000&
    
    txtStatus(1).BackColor = &HFF&
    txtStatus(1).Text = "Sending..."
    If CurrentBasicHigh <> 0 Then
        URL_Part2 = "&type=Basic&high=" & CurrentBasicHigh & "&mob=" & CurrentBasicMob & "&acc=" & CurrentBasicAcc
        Result = inet.OpenURL(URL_Part1 & URL_Part2)
        If Result = "error" Then
            txtStatus(1).Text = "Failed."
            txtStatus(1).BackColor = vbBlack
        Else
            txtStatus(1).Text = "Done."
            txtStatus(1).BackColor = &HC000&
        End If
    Else
        txtStatus(1).Text = "Skipped."
        txtStatus(1).BackColor = &HC000C0
    End If
    DoEvents
    
    txtStatus(2).BackColor = &HFF&
    txtStatus(2).Text = "Sending..."
    If CurrentSkillHigh <> 0 Then
        URL_Part2 = "&type=Skill&high=" & CurrentSkillHigh & "&mob=" & CurrentSkillMob & "&skilltype=" & CurrentSkill
        Result = inet.OpenURL(URL_Part1 & URL_Part2)
        If Result = "error" Then
            txtStatus(2).Text = "Failed."
            txtStatus(2).BackColor = vbBlack
        Else
            txtStatus(2).Text = "Done."
            txtStatus(2).BackColor = &HC000&
        End If
    Else
        txtStatus(2).Text = "Skipped."
        txtStatus(2).BackColor = &HC000C0
    End If
    
    txtStatus(3).BackColor = &HFF&
    txtStatus(3).Text = "Sending..."
    If CurrentRangedHigh <> 0 Then
        URL_Part2 = "&type=Ranged&high=" & CurrentRangedHigh & "&mob=" & CurrentRangedMob & "&acc=" & CurrentRangedAcc
        Result = inet.OpenURL(URL_Part1 & URL_Part2)
        If Result = "error" Then
            txtStatus(3).Text = "Failed."
            txtStatus(3).BackColor = vbBlack
        Else
            txtStatus(3).Text = "Done."
            txtStatus(3).BackColor = &HC000&
        End If
    Else
        txtStatus(3).Text = "Skipped."
        txtStatus(3).BackColor = &HC000C0
    End If
    
    txtStatus(4).BackColor = &HFF&
    txtStatus(4).Text = "Sending..."
    If CurrentSpellHigh <> 0 Then
        URL_Part2 = "&type=Spell&high=" & CurrentSpellHigh & "&mob=" & CurrentSpellMob & "&skilltype=" & CurrentSpell
        Result = inet.OpenURL(URL_Part1 & URL_Part2)
        If Result = "error" Then
            txtStatus(4).Text = "Failed."
            txtStatus(4).BackColor = vbBlack
        Else
            txtStatus(4).Text = "Done."
            txtStatus(4).BackColor = &HC000&
        End If
    Else
        txtStatus(4).Text = "Skipped."
        txtStatus(4).BackColor = &HC000C0
    End If
Else
    lblStatus(5).Caption = "Sending Highs..."
    For i = 0 To txtStatus.UBound
        txtStatus(i).Text = "Waiting..."
        txtStatus(i).BackColor = &HFF0000
    Next
    txtStatus(0).BackColor = &HFF&
    For i = 0 To UBound(FullStats)
        txtStatus(0).Text = i & " of " & UBound(FullStats)
        DoEvents
        If FullStats(i).Attacker = UserName Then
            If FullStats(i).Basic.High > CurrentBasicHigh Then
                CurrentBasicHigh = FullStats(i).Basic.High
                CurrentBasicMob = FullStats(i).Defender
                CurrentBasicAcc = Round((FullStats(i).Basic.Hit / (FullStats(i).Basic.Hit + FullStats(i).Basic.Miss)) * 100, 2)
            End If
            If FullStats(i).Skill.High > CurrentSkillHigh Then
                CurrentSkillHigh = FullStats(i).Skill.High
                CurrentSkillMob = FullStats(i).Defender
                CurrentSkill = FullStats(i).Skill.HighSkillType
            End If
            If FullStats(i).Ranged.High > CurrentRangedHigh Then
                CurrentRangedHigh = FullStats(i).Ranged.High
                CurrentRangedMob = FullStats(i).Defender
                CurrentRangedAcc = Round((FullStats(i).Ranged.Hit / (FullStats(i).Ranged.Hit + FullStats(i).Ranged.Miss)) * 100, 2)
            End If
            If FullStats(i).Spell.High > CurrentSpellHigh Then
                CurrentSpellHigh = FullStats(i).Spell.High
                CurrentSpellMob = FullStats(i).Defender
                CurrentSpell = FullStats(i).Spell.HighSkillType
            End If
        End If
    Next
    txtStatus(0).Text = "Done."
    txtStatus(0).BackColor = &HC000&
    
    txtStatus(1).BackColor = &HFF&
    txtStatus(1).Text = "Sending..."
    If CurrentBasicHigh <> 0 Then
        URL_Part2 = "&type=Basic&high=" & CurrentBasicHigh & "&mob=" & CurrentBasicMob & "&acc=" & CurrentBasicAcc
        Result = inet.OpenURL(URL_Part1 & URL_Part2)
        If Result = "error" Then
            txtStatus(1).Text = "Failed."
            txtStatus(1).BackColor = vbBlack
        Else
            txtStatus(1).Text = "Done."
            txtStatus(1).BackColor = &HC000&
        End If
    Else
        txtStatus(1).Text = "Skipped."
        txtStatus(1).BackColor = &HC000C0
    End If
    DoEvents
    
    txtStatus(2).BackColor = &HFF&
    txtStatus(2).Text = "Sending..."
    If CurrentSkillHigh <> 0 Then
        URL_Part2 = "&type=Skill&high=" & CurrentSkillHigh & "&mob=" & CurrentSkillMob & "&skilltype=" & CurrentSkill
        Result = inet.OpenURL(URL_Part1 & URL_Part2)
        If Result = "error" Then
            txtStatus(2).Text = "Failed."
            txtStatus(2).BackColor = vbBlack
        Else
            txtStatus(2).Text = "Done."
            txtStatus(2).BackColor = &HC000&
        End If
    Else
        txtStatus(2).Text = "Skipped."
        txtStatus(2).BackColor = &HC000C0
    End If
    
    txtStatus(3).BackColor = &HFF&
    txtStatus(3).Text = "Sending..."
    If CurrentRangedHigh <> 0 Then
        URL_Part2 = "&type=Ranged&high=" & CurrentRangedHigh & "&mob=" & CurrentRangedMob & "&acc=" & CurrentRangedAcc
        Result = inet.OpenURL(URL_Part1 & URL_Part2)
        If Result = "error" Then
            txtStatus(3).Text = "Failed."
            txtStatus(3).BackColor = vbBlack
        Else
            txtStatus(3).Text = "Done."
            txtStatus(3).BackColor = &HC000&
        End If
    Else
        txtStatus(3).Text = "Skipped."
        txtStatus(3).BackColor = &HC000C0
    End If
    
    txtStatus(4).BackColor = &HFF&
    txtStatus(4).Text = "Sending..."
    If CurrentSpellHigh <> 0 Then
        URL_Part2 = "&type=Spell&high=" & CurrentSpellHigh & "&mob=" & CurrentSpellMob & "&skilltype=" & CurrentSpell
        Result = inet.OpenURL(URL_Part1 & URL_Part2)
        If Result = "error" Then
            txtStatus(4).Text = "Failed."
            txtStatus(4).BackColor = vbBlack
        Else
            txtStatus(4).Text = "Done."
            txtStatus(4).BackColor = &HC000&
        End If
    Else
        txtStatus(4).Text = "Skipped."
        txtStatus(4).BackColor = &HC000C0
    End If
End If

MsgBox "Finished!", vbInformation, "Send"
Screen.MousePointer = vbDefault
frmTransmit.SetFocus
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdSend_Click()
If optionSend(3).Value = False Then
    SendMOB
Else
    SendCraft
End If
End Sub


Private Sub Form_Load()
Dim i, MobList As String
Me.Left = frmRead.Left + 100
Me.Top = frmRead.Top + 100
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
For i = 0 To frmRead.listResults.ListCount
    If InStr(1, MobList, Mid(frmRead.listResults.List(i), InStr(1, frmRead.listResults.List(i), "-") + 1)) = 0 Then
        comboMOB.AddItem Mid(frmRead.listResults.List(i), InStr(1, frmRead.listResults.List(i), "-") + 1)
        MobList = MobList & Mid(frmRead.listResults.List(i), InStr(1, frmRead.listResults.List(i), "-") + 1)
    End If
Next
If UBound(CraftingCSV) = 0 Then
    optionSend(3).Enabled = False
End If
If comboPlayer.ListCount <> 0 Then
    comboPlayer.ListIndex = 0
End If
End Sub
Private Sub Label1_Click()
frmSetup.Show
Unload Me
End Sub


Private Sub optionSend_Click(Index As Integer)
lblStatus(0).Caption = "Scanning Statistics:"
If optionSend(1).Value = True Then
    Label2.Enabled = True
    comboMOB.Enabled = True
ElseIf optionSend(3).Value = True Then
    comboMOB.Enabled = False
    lblStatus(0).Caption = "Crafting Data:"
Else
    Label2.Enabled = False
    comboMOB.Enabled = False
End If
End Sub


