VERSION 5.00
Begin VB.Form frmReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporting Options"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4590
   Icon            =   "frmReport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox checkOption 
      Caption         =   "Ability"
      Height          =   240
      Index           =   37
      Left            =   60
      TabIndex        =   39
      Top             =   1026
      Width           =   735
   End
   Begin VB.CheckBox checkOption 
      Caption         =   "Skill"
      Height          =   240
      Index           =   1
      Left            =   60
      TabIndex        =   38
      Top             =   792
      Width           =   735
   End
   Begin VB.CheckBox checkOption 
      Caption         =   "Parries"
      Height          =   240
      Index           =   12
      Left            =   3045
      TabIndex        =   37
      Top             =   324
      Width           =   1005
   End
   Begin VB.CheckBox checkOption 
      Caption         =   "Evades"
      Height          =   240
      Index           =   11
      Left            =   3060
      TabIndex        =   36
      Top             =   90
      Width           =   1005
   End
   Begin VB.CheckBox checkOption 
      Caption         =   "R. Hit %"
      Height          =   240
      Index           =   26
      Left            =   60
      TabIndex        =   35
      ToolTipText     =   "Ranged Hit %"
      Top             =   2430
      Width           =   1230
   End
   Begin VB.CheckBox checkOption 
      Caption         =   "Sp. MP Used"
      Height          =   240
      Index           =   36
      Left            =   1380
      TabIndex        =   34
      ToolTipText     =   "MP Used Attacking"
      Top             =   558
      Width           =   1515
   End
   Begin VB.CheckBox checkOption 
      Caption         =   "HP. MP Used"
      Height          =   240
      Index           =   35
      Left            =   3045
      TabIndex        =   33
      ToolTipText     =   "MP Used Healing"
      Top             =   2430
      Width           =   1515
   End
   Begin VB.CheckBox checkOption 
      Caption         =   "Ab. Average"
      Height          =   240
      Index           =   34
      Left            =   1380
      TabIndex        =   32
      ToolTipText     =   "Ability Average"
      Top             =   1740
      Width           =   1290
   End
   Begin VB.CheckBox checkOption 
      Caption         =   "Ab. High/Low"
      Height          =   240
      Index           =   33
      Left            =   1380
      TabIndex        =   31
      ToolTipText     =   "Ability High/Low"
      Top             =   1500
      Width           =   1410
   End
   Begin VB.CheckBox checkOption 
      Caption         =   "Sk. High/Low"
      Height          =   240
      Index           =   32
      Left            =   1380
      TabIndex        =   30
      ToolTipText     =   "Skill High/Low"
      Top             =   792
      Width           =   1410
   End
   Begin VB.CheckBox checkOption 
      Caption         =   "Sk. Average"
      Height          =   240
      Index           =   31
      Left            =   1380
      TabIndex        =   29
      ToolTipText     =   "Skill Average"
      Top             =   1026
      Width           =   1290
   End
   Begin VB.CheckBox checkOption 
      Caption         =   "Sp. High/Low"
      Height          =   240
      Index           =   30
      Left            =   1380
      TabIndex        =   28
      ToolTipText     =   "Spell High/Low"
      Top             =   90
      Width           =   1410
   End
   Begin VB.CheckBox checkOption 
      Caption         =   "Sp. Average"
      Height          =   240
      Index           =   29
      Left            =   1380
      TabIndex        =   27
      ToolTipText     =   "Spell Average"
      Top             =   324
      Width           =   1470
   End
   Begin VB.CheckBox checkOption 
      Caption         =   "R. Average"
      Height          =   240
      Index           =   28
      Left            =   60
      TabIndex        =   26
      ToolTipText     =   "Ranged Average"
      Top             =   3120
      Width           =   1170
   End
   Begin VB.CheckBox checkOption 
      Caption         =   "R. High/Low"
      Height          =   240
      Index           =   27
      Left            =   60
      TabIndex        =   25
      ToolTipText     =   "Ranged High/Low"
      Top             =   2898
      Width           =   1230
   End
   Begin VB.CheckBox checkOption 
      Caption         =   "R. Hit/Miss"
      Height          =   240
      Index           =   25
      Left            =   60
      TabIndex        =   24
      ToolTipText     =   "Ranged Hit/Miss"
      Top             =   2664
      Width           =   1245
   End
   Begin VB.CheckBox checkOption 
      Caption         =   "Ranged"
      Height          =   240
      Index           =   24
      Left            =   60
      TabIndex        =   23
      Top             =   324
      Width           =   915
   End
   Begin VB.CheckBox checkOption 
      Caption         =   "% of TTL DMG"
      Height          =   240
      Index           =   18
      Left            =   3045
      TabIndex        =   22
      Top             =   2664
      Width           =   1410
   End
   Begin VB.CheckBox checkOption 
      Caption         =   "Counters"
      Height          =   240
      Index           =   23
      Left            =   3045
      TabIndex        =   21
      Top             =   1494
      Width           =   1185
   End
   Begin VB.CheckBox checkOption 
      Caption         =   "Sk. Uses"
      Height          =   240
      Index           =   22
      Left            =   1380
      TabIndex        =   20
      ToolTipText     =   "WeaponSkill Count"
      Top             =   1260
      Width           =   1095
   End
   Begin VB.CheckBox checkOption 
      Caption         =   "Addt'l Effect"
      Height          =   240
      Index           =   21
      Left            =   60
      TabIndex        =   19
      Top             =   1260
      Width           =   1185
   End
   Begin VB.CheckBox checkOption 
      Caption         =   "Anticipates"
      Height          =   240
      Index           =   20
      Left            =   3045
      TabIndex        =   18
      Top             =   1260
      Width           =   1185
   End
   Begin VB.CheckBox checkOption 
      Caption         =   "HP Healed"
      Height          =   240
      Index           =   19
      Left            =   3045
      TabIndex        =   17
      Top             =   2196
      Width           =   1410
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   330
      Left            =   1380
      TabIndex        =   16
      Top             =   3060
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   330
      Left            =   3300
      TabIndex        =   15
      Top             =   3060
      Width           =   1140
   End
   Begin VB.CheckBox checkOption 
      Caption         =   "HP Recovered"
      Height          =   240
      Index           =   17
      Left            =   3045
      TabIndex        =   14
      Top             =   1962
      Width           =   1410
   End
   Begin VB.CheckBox checkOption 
      Caption         =   "Critical Hit %"
      Height          =   240
      Index           =   5
      Left            =   1380
      TabIndex        =   13
      Top             =   1965
      Width           =   1215
   End
   Begin VB.CheckBox checkOption 
      Caption         =   "DMG Taken"
      Height          =   240
      Index           =   16
      Left            =   3045
      TabIndex        =   12
      Top             =   1728
      Width           =   1275
   End
   Begin VB.CheckBox checkOption 
      Caption         =   "Avoids"
      Height          =   240
      Index           =   15
      Left            =   3045
      TabIndex        =   11
      Top             =   1026
      Width           =   915
   End
   Begin VB.CheckBox checkOption 
      Caption         =   "Absorbs"
      Height          =   240
      Index           =   14
      Left            =   3045
      TabIndex        =   10
      Top             =   792
      Width           =   915
   End
   Begin VB.CheckBox checkOption 
      Caption         =   "Blocks"
      Height          =   240
      Index           =   13
      Left            =   3045
      TabIndex        =   9
      Top             =   558
      Width           =   1050
   End
   Begin VB.CheckBox checkOption 
      Caption         =   "Take/Avoid"
      Height          =   240
      Index           =   10
      Left            =   1380
      TabIndex        =   8
      Top             =   2664
      Width           =   1230
   End
   Begin VB.CheckBox checkOption 
      Caption         =   "Avoid %"
      Height          =   240
      Index           =   9
      Left            =   1380
      TabIndex        =   7
      Top             =   2430
      Width           =   915
   End
   Begin VB.CheckBox checkOption 
      Caption         =   "M. Hit/Miss"
      Height          =   240
      Index           =   8
      Left            =   60
      TabIndex        =   6
      ToolTipText     =   "Melee Hit/Miss"
      Top             =   1728
      Width           =   1365
   End
   Begin VB.CheckBox checkOption 
      Caption         =   "M. Hit %"
      Height          =   240
      Index           =   7
      Left            =   60
      TabIndex        =   5
      ToolTipText     =   "Melee Hit %"
      Top             =   1494
      Width           =   1230
   End
   Begin VB.CheckBox checkOption 
      Caption         =   "Critical Hits"
      Height          =   240
      Index           =   6
      Left            =   1380
      TabIndex        =   4
      Top             =   2205
      Width           =   1155
   End
   Begin VB.CheckBox checkOption 
      Caption         =   "M. Average"
      Height          =   240
      Index           =   4
      Left            =   60
      TabIndex        =   3
      ToolTipText     =   "Melee Average"
      Top             =   2196
      Width           =   1170
   End
   Begin VB.CheckBox checkOption 
      Caption         =   "M. High/Low"
      Height          =   240
      Index           =   3
      Left            =   60
      TabIndex        =   2
      ToolTipText     =   "Melee High/Low"
      Top             =   1962
      Width           =   1230
   End
   Begin VB.CheckBox checkOption 
      Caption         =   "Spell"
      Height          =   240
      Index           =   2
      Left            =   60
      TabIndex        =   1
      Top             =   558
      Width           =   735
   End
   Begin VB.CheckBox checkOption 
      Caption         =   "Melee"
      Height          =   240
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   90
      Width           =   735
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Sub cmdCancel_Click()
Unload Me
Set frmReport = Nothing
End Sub


Private Sub cmdOK_Click()
Dim i
For i = 0 To checkOption.UBound
    SaveSetting App.Title, "Settings", "Report" & i, checkOption(i).Value
    ReportOptions(i) = checkOption(i).Value
Next
Unload Me
Set frmReport = Nothing
End Sub


Private Sub Form_Load()
frmRead.Enabled = False
Me.Left = frmRead.Left + 100
Me.Top = frmRead.Top + 100
Dim i
For i = 0 To checkOption.UBound
    checkOption(i).Value = ReportOptions(i)
Next
End Sub


Private Sub Form_Unload(Cancel As Integer)
frmRead.Enabled = True
End Sub





