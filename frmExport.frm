VERSION 5.00
Begin VB.Form frmExport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Export Report"
   ClientHeight    =   7185
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5070
   Icon            =   "frmExport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox checkOption 
      Caption         =   "Rng Hit %"
      Height          =   240
      Index           =   26
      Left            =   90
      TabIndex        =   30
      Top             =   6255
      Width           =   1140
   End
   Begin VB.CheckBox checkOption 
      Caption         =   "Rng Hit/Miss"
      Height          =   240
      Index           =   25
      Left            =   90
      TabIndex        =   29
      Top             =   6030
      Width           =   1320
   End
   Begin VB.CheckBox checkOption 
      Caption         =   "Ranged"
      Height          =   240
      Index           =   24
      Left            =   90
      TabIndex        =   28
      Top             =   5805
      Width           =   1185
   End
   Begin VB.CheckBox checkOption 
      Caption         =   "Counters"
      Height          =   240
      Index           =   23
      Left            =   90
      TabIndex        =   27
      Top             =   5580
      Width           =   1185
   End
   Begin VB.CheckBox checkOption 
      Caption         =   "WS Count"
      Height          =   240
      Index           =   22
      Left            =   90
      TabIndex        =   26
      Top             =   5355
      Width           =   1095
   End
   Begin VB.CheckBox checkOption 
      Caption         =   "Addt'l Effect"
      Height          =   240
      Index           =   21
      Left            =   90
      TabIndex        =   25
      Top             =   5130
      Width           =   1185
   End
   Begin VB.CheckBox checkOption 
      Caption         =   "Anticipates"
      Height          =   240
      Index           =   20
      Left            =   90
      TabIndex        =   24
      Top             =   4905
      Width           =   1185
   End
   Begin VB.CheckBox checkOption 
      Caption         =   "HP Healed"
      Height          =   240
      Index           =   19
      Left            =   90
      TabIndex        =   23
      Top             =   4680
      Width           =   1410
   End
   Begin VB.CheckBox checkOption 
      Caption         =   "Summary Only"
      Height          =   240
      Index           =   18
      Left            =   90
      TabIndex        =   22
      Top             =   4455
      Width           =   1410
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   330
      Left            =   3330
      TabIndex        =   21
      Top             =   5985
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   330
      Left            =   3375
      TabIndex        =   20
      Top             =   6300
      Width           =   1140
   End
   Begin VB.CheckBox checkOption 
      Caption         =   "HP Recovered"
      Height          =   240
      Index           =   17
      Left            =   90
      TabIndex        =   19
      Top             =   4230
      Width           =   1410
   End
   Begin VB.CheckBox checkOption 
      Caption         =   "Evades"
      Height          =   240
      Index           =   11
      Left            =   90
      TabIndex        =   18
      Top             =   2880
      Width           =   1005
   End
   Begin VB.CheckBox checkOption 
      Caption         =   "Crit %"
      Height          =   240
      Index           =   5
      Left            =   90
      TabIndex        =   17
      Top             =   1530
      Width           =   735
   End
   Begin VB.CheckBox checkOption 
      Caption         =   "DMG Taken"
      Height          =   240
      Index           =   16
      Left            =   90
      TabIndex        =   16
      Top             =   4005
      Width           =   1275
   End
   Begin VB.CheckBox checkOption 
      Caption         =   "Avoids"
      Height          =   240
      Index           =   15
      Left            =   90
      TabIndex        =   15
      Top             =   3780
      Width           =   915
   End
   Begin VB.CheckBox checkOption 
      Caption         =   "Absorbs"
      Height          =   240
      Index           =   14
      Left            =   90
      TabIndex        =   14
      Top             =   3555
      Width           =   915
   End
   Begin VB.CheckBox checkOption 
      Caption         =   "Blocks"
      Height          =   240
      Index           =   13
      Left            =   90
      TabIndex        =   13
      Top             =   3330
      Width           =   1050
   End
   Begin VB.CheckBox checkOption 
      Caption         =   "Parries"
      Height          =   240
      Index           =   12
      Left            =   90
      TabIndex        =   12
      Top             =   3105
      Width           =   1005
   End
   Begin VB.CheckBox checkOption 
      Caption         =   "Take/Avoid"
      Height          =   240
      Index           =   10
      Left            =   90
      TabIndex        =   11
      Top             =   2655
      Width           =   1230
   End
   Begin VB.CheckBox checkOption 
      Caption         =   "Avoid %"
      Height          =   240
      Index           =   9
      Left            =   90
      TabIndex        =   10
      Top             =   2430
      Width           =   915
   End
   Begin VB.CheckBox checkOption 
      Caption         =   "Melee Hit/Miss"
      Height          =   240
      Index           =   8
      Left            =   90
      TabIndex        =   9
      Top             =   2205
      Width           =   1590
   End
   Begin VB.CheckBox checkOption 
      Caption         =   "Melee Hit %"
      Height          =   240
      Index           =   7
      Left            =   90
      TabIndex        =   8
      Top             =   1980
      Width           =   1410
   End
   Begin VB.CheckBox checkOption 
      Caption         =   "Crit #"
      Height          =   240
      Index           =   6
      Left            =   90
      TabIndex        =   7
      Top             =   1755
      Width           =   735
   End
   Begin VB.CheckBox checkOption 
      Caption         =   "Average"
      Height          =   240
      Index           =   4
      Left            =   90
      TabIndex        =   6
      Top             =   1305
      Width           =   1050
   End
   Begin VB.CheckBox checkOption 
      Caption         =   "High/Low"
      Height          =   240
      Index           =   3
      Left            =   90
      TabIndex        =   5
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CheckBox checkOption 
      Caption         =   "Spell"
      Height          =   240
      Index           =   2
      Left            =   90
      TabIndex        =   4
      Top             =   855
      Width           =   735
   End
   Begin VB.CheckBox checkOption 
      Caption         =   "Skill"
      Height          =   240
      Index           =   1
      Left            =   90
      TabIndex        =   3
      Top             =   630
      Width           =   735
   End
   Begin VB.TextBox txtExport 
      Height          =   285
      Left            =   945
      TabIndex        =   1
      Text            =   "export.html"
      Top             =   45
      Width           =   3210
   End
   Begin VB.CheckBox checkOption 
      Caption         =   "Melee"
      Height          =   240
      Index           =   0
      Left            =   90
      TabIndex        =   0
      Top             =   405
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Filename:"
      Height          =   240
      Left            =   90
      TabIndex        =   2
      Top             =   90
      Width           =   1320
   End
End
Attribute VB_Name = "frmExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Sub cmdCancel_Click()
Unload Me
Set frmExport = Nothing
End Sub


Private Sub cmdOK_Click()
Dim i
For i = 0 To checkOption.ubound
    SaveSetting App.Title, "Settings", "Export" & i, checkOption(i).Value
    ExportOptions(i, 0) = checkOption(i).Value
Next
ExportFile = txtExport.Text

frmRead.ExportHTML checkOption(18)
MsgBox "Done!", vbOKOnly + vbInformation, "Export"
Unload Me
End Sub


Private Sub Form_Load()
txtExport.Text = Format(Date, "MM-DD-YYYY") & ".html"
frmRead.Enabled = False
Me.Left = frmRead.Left + 100
Me.Top = frmRead.Top + 100
Dim i
For i = 0 To checkOption.ubound
    checkOption(i).Value = GetSetting(App.Title, "Settings", "Export" & i, Default:=1)
Next
End Sub


Private Sub Form_Unload(Cancel As Integer)
frmRead.Enabled = True
End Sub


Private Sub txtExport_Validate(Cancel As Boolean)
If InStr(1, txtExport, "\") Then
    txtExport.Text = "export.html"
    MsgBox "Invalid filename.", vbInformation, "Filename"
ElseIf InStr(1, txtExport, ":") Then
    txtExport.Text = "export.html"
    MsgBox "Invalid filename.", vbInformation, "Filename"
ElseIf txtExport = "" Then
    txtExport.Text = "export.html"
End If
End Sub


