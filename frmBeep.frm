VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmBeep 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Timer Sound Setup"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5265
   Icon            =   "frmBeep.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   5265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtBeep 
      Height          =   285
      Index           =   10
      Left            =   420
      TabIndex        =   45
      Top             =   3480
      Width           =   4155
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   ">"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   10
      Left            =   4920
      TabIndex        =   44
      ToolTipText     =   "Play Sound"
      Top             =   3480
      Width           =   285
   End
   Begin VB.CommandButton cmdBeep 
      Caption         =   "..."
      Height          =   285
      Index           =   10
      Left            =   4605
      TabIndex        =   43
      ToolTipText     =   "Select WAV File"
      Top             =   3480
      Width           =   285
   End
   Begin VB.CommandButton cmdBeep 
      Caption         =   "..."
      Height          =   285
      Index           =   9
      Left            =   4605
      TabIndex        =   42
      ToolTipText     =   "Select WAV File"
      Top             =   3150
      Width           =   285
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   ">"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   4920
      Picture         =   "frmBeep.frx":1E72
      TabIndex        =   41
      ToolTipText     =   "Play Sound"
      Top             =   315
      Width           =   285
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   ">"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   4920
      TabIndex        =   40
      ToolTipText     =   "Play Sound"
      Top             =   630
      Width           =   285
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   ">"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   4920
      TabIndex        =   39
      ToolTipText     =   "Play Sound"
      Top             =   945
      Width           =   285
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   ">"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   4920
      TabIndex        =   38
      ToolTipText     =   "Play Sound"
      Top             =   1260
      Width           =   285
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   ">"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   4920
      TabIndex        =   37
      ToolTipText     =   "Play Sound"
      Top             =   1575
      Width           =   285
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   ">"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   4920
      TabIndex        =   36
      ToolTipText     =   "Play Sound"
      Top             =   1890
      Width           =   285
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   ">"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   6
      Left            =   4920
      TabIndex        =   35
      ToolTipText     =   "Play Sound"
      Top             =   2205
      Width           =   285
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   ">"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   7
      Left            =   4920
      TabIndex        =   34
      ToolTipText     =   "Play Sound"
      Top             =   2520
      Width           =   285
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   ">"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   8
      Left            =   4920
      TabIndex        =   33
      ToolTipText     =   "Play Sound"
      Top             =   2835
      Width           =   285
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   ">"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   9
      Left            =   4920
      TabIndex        =   32
      ToolTipText     =   "Play Sound"
      Top             =   3150
      Width           =   285
   End
   Begin MSComDlg.CommonDialog CDBeep 
      Left            =   2400
      Top             =   4080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   330
      Left            =   4065
      TabIndex        =   31
      Top             =   3810
      Width           =   1140
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   330
      Left            =   105
      TabIndex        =   30
      Top             =   3810
      Width           =   1140
   End
   Begin VB.CheckBox checkBeeps 
      Caption         =   "Just use system beep and let me specify the # of beeps instead."
      Height          =   240
      Left            =   420
      TabIndex        =   29
      Top             =   45
      Width           =   4830
   End
   Begin VB.TextBox txtBeep 
      Height          =   285
      Index           =   9
      Left            =   420
      TabIndex        =   28
      Top             =   3150
      Width           =   4155
   End
   Begin VB.CommandButton cmdBeep 
      Caption         =   "..."
      Height          =   285
      Index           =   8
      Left            =   4605
      TabIndex        =   27
      ToolTipText     =   "Select WAV File"
      Top             =   2835
      Width           =   285
   End
   Begin VB.TextBox txtBeep 
      Height          =   285
      Index           =   8
      Left            =   420
      TabIndex        =   26
      Top             =   2835
      Width           =   4155
   End
   Begin VB.CommandButton cmdBeep 
      Caption         =   "..."
      Height          =   285
      Index           =   7
      Left            =   4605
      TabIndex        =   25
      ToolTipText     =   "Select WAV File"
      Top             =   2520
      Width           =   285
   End
   Begin VB.TextBox txtBeep 
      Height          =   285
      Index           =   7
      Left            =   420
      TabIndex        =   24
      Top             =   2520
      Width           =   4155
   End
   Begin VB.CommandButton cmdBeep 
      Caption         =   "..."
      Height          =   285
      Index           =   6
      Left            =   4605
      TabIndex        =   23
      ToolTipText     =   "Select WAV File"
      Top             =   2205
      Width           =   285
   End
   Begin VB.TextBox txtBeep 
      Height          =   285
      Index           =   6
      Left            =   420
      TabIndex        =   22
      Top             =   2205
      Width           =   4155
   End
   Begin VB.CommandButton cmdBeep 
      Caption         =   "..."
      Height          =   285
      Index           =   5
      Left            =   4605
      TabIndex        =   21
      ToolTipText     =   "Select WAV File"
      Top             =   1890
      Width           =   285
   End
   Begin VB.TextBox txtBeep 
      Height          =   285
      Index           =   5
      Left            =   420
      TabIndex        =   20
      Top             =   1890
      Width           =   4155
   End
   Begin VB.CommandButton cmdBeep 
      Caption         =   "..."
      Height          =   285
      Index           =   4
      Left            =   4605
      TabIndex        =   19
      ToolTipText     =   "Select WAV File"
      Top             =   1575
      Width           =   285
   End
   Begin VB.TextBox txtBeep 
      Height          =   285
      Index           =   4
      Left            =   420
      TabIndex        =   18
      Top             =   1575
      Width           =   4155
   End
   Begin VB.CommandButton cmdBeep 
      Caption         =   "..."
      Height          =   285
      Index           =   3
      Left            =   4605
      TabIndex        =   17
      ToolTipText     =   "Select WAV File"
      Top             =   1260
      Width           =   285
   End
   Begin VB.TextBox txtBeep 
      Height          =   285
      Index           =   3
      Left            =   420
      TabIndex        =   16
      Top             =   1260
      Width           =   4155
   End
   Begin VB.CommandButton cmdBeep 
      Caption         =   "..."
      Height          =   285
      Index           =   2
      Left            =   4605
      TabIndex        =   15
      ToolTipText     =   "Select WAV File"
      Top             =   945
      Width           =   285
   End
   Begin VB.TextBox txtBeep 
      Height          =   285
      Index           =   2
      Left            =   420
      TabIndex        =   14
      Top             =   945
      Width           =   4155
   End
   Begin VB.CommandButton cmdBeep 
      Caption         =   "..."
      Height          =   285
      Index           =   1
      Left            =   4605
      TabIndex        =   13
      ToolTipText     =   "Select WAV File"
      Top             =   630
      Width           =   285
   End
   Begin VB.TextBox txtBeep 
      Height          =   285
      Index           =   1
      Left            =   420
      TabIndex        =   12
      Top             =   630
      Width           =   4155
   End
   Begin VB.CommandButton cmdBeep 
      Caption         =   "..."
      Height          =   285
      Index           =   0
      Left            =   4605
      TabIndex        =   11
      ToolTipText     =   "Select WAV File"
      Top             =   315
      Width           =   285
   End
   Begin VB.TextBox txtBeep 
      Height          =   285
      Index           =   0
      Left            =   420
      TabIndex        =   10
      Top             =   315
      Width           =   4155
   End
   Begin VB.Label lblBeep 
      Caption         =   "Rdy:"
      Height          =   240
      Index           =   10
      Left            =   90
      TabIndex        =   46
      Top             =   3495
      Width           =   360
   End
   Begin VB.Label lblBeep 
      Caption         =   "10:"
      Height          =   240
      Index           =   9
      Left            =   90
      TabIndex        =   9
      Top             =   3165
      Width           =   240
   End
   Begin VB.Label lblBeep 
      Caption         =   "9:"
      Height          =   240
      Index           =   8
      Left            =   90
      TabIndex        =   8
      Top             =   2850
      Width           =   240
   End
   Begin VB.Label lblBeep 
      Caption         =   "8:"
      Height          =   240
      Index           =   7
      Left            =   90
      TabIndex        =   7
      Top             =   2535
      Width           =   240
   End
   Begin VB.Label lblBeep 
      Caption         =   "7:"
      Height          =   240
      Index           =   6
      Left            =   90
      TabIndex        =   6
      Top             =   2220
      Width           =   240
   End
   Begin VB.Label lblBeep 
      Caption         =   "6:"
      Height          =   240
      Index           =   5
      Left            =   90
      TabIndex        =   5
      Top             =   1905
      Width           =   240
   End
   Begin VB.Label lblBeep 
      Caption         =   "5:"
      Height          =   240
      Index           =   4
      Left            =   90
      TabIndex        =   4
      Top             =   1590
      Width           =   240
   End
   Begin VB.Label lblBeep 
      Caption         =   "4:"
      Height          =   240
      Index           =   3
      Left            =   90
      TabIndex        =   3
      Top             =   1275
      Width           =   240
   End
   Begin VB.Label lblBeep 
      Caption         =   "3:"
      Height          =   240
      Index           =   2
      Left            =   90
      TabIndex        =   2
      Top             =   960
      Width           =   240
   End
   Begin VB.Label lblBeep 
      Caption         =   "2:"
      Height          =   240
      Index           =   1
      Left            =   90
      TabIndex        =   1
      Top             =   645
      Width           =   240
   End
   Begin VB.Label lblBeep 
      Caption         =   "1:"
      Height          =   240
      Index           =   0
      Left            =   90
      TabIndex        =   0
      Top             =   330
      Width           =   240
   End
End
Attribute VB_Name = "frmBeep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub checkBeeps_Click()
BeepNotWave = checkBeeps.Value
End Sub

Private Sub cmdBeep_Click(Index As Integer)
On Error GoTo Err_Handler
CDBeep.Flags = cdlOFNHideReadOnly
CDBeep.DialogTitle = "Select Wave File"
CDBeep.Filter = "WAV (*.wav)|*.wav"
CDBeep.CancelError = True
CDBeep.InitDir = App.Path
CDBeep.ShowOpen
txtBeep(Index).Text = CDBeep.FileName
If checkBeeps.Value = 1 Then
    If MsgBox("System Beeps Only must be unchecked for these to work." & vbNewLine & vbNewLine & "Uncheck System Beeps?", vbYesNo + vbQuestion, "Sounds") = vbYes Then
        checkBeeps.Value = 0
    End If
End If
Exit Sub

Err_Handler:
If Err.Number = 32755 Then
    Exit Sub
End If
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
Dim i As Integer
For i = 0 To UBound(BeepSounds)
    BeepSounds(i) = txtBeep(i).Text
    SaveSetting App.Title, "Settings", "Sound-" & i, txtBeep(i).Text
Next
SaveSetting App.Title, "Settings", "Beeps", checkBeeps.Value
frmRead.BeepNotWave = checkBeeps.Value
Unload Me
End Sub

Private Sub cmdPlay_Click(Index As Integer)
PlaySound txtBeep(Index).Text, 0&, SND_FILENAME Or SND_ASYNC
End Sub

Private Sub Form_Load()
frmRead.Enabled = False
If frmRead.BeepNotWave = True Then
    checkBeeps.Value = 1
Else
    checkBeeps.Value = 0
End If
Me.Left = frmRead.Left + 100
Me.Top = frmRead.Top + 100
Dim i As Integer
For i = 0 To UBound(BeepSounds)
    txtBeep(i).Text = BeepSounds(i)
Next
End Sub


Private Sub Form_Unload(Cancel As Integer)
frmRead.Enabled = True
End Sub


