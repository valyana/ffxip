VERSION 5.00
Begin VB.Form frmKeylogger 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Key Activated Timers"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6135
   Icon            =   "frmKeylogger.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
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
      Left            =   5760
      TabIndex        =   66
      ToolTipText     =   "Play Sound"
      Top             =   3555
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
      Left            =   5760
      TabIndex        =   65
      ToolTipText     =   "Play Sound"
      Top             =   3195
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
      Left            =   5760
      TabIndex        =   64
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
      Index           =   6
      Left            =   5760
      TabIndex        =   63
      ToolTipText     =   "Play Sound"
      Top             =   2475
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
      Left            =   5760
      TabIndex        =   62
      ToolTipText     =   "Play Sound"
      Top             =   2115
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
      Left            =   5760
      TabIndex        =   61
      ToolTipText     =   "Play Sound"
      Top             =   1755
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
      Left            =   5760
      TabIndex        =   60
      ToolTipText     =   "Play Sound"
      Top             =   1395
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
      Left            =   5760
      TabIndex        =   59
      ToolTipText     =   "Play Sound"
      Top             =   1035
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
      Left            =   5760
      TabIndex        =   58
      ToolTipText     =   "Play Sound"
      Top             =   675
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
      Left            =   5760
      Picture         =   "frmKeylogger.frx":1E72
      TabIndex        =   57
      ToolTipText     =   "Play Sound"
      Top             =   335
      Width           =   285
   End
   Begin VB.TextBox txtMessage 
      Height          =   315
      Index           =   9
      Left            =   3105
      TabIndex        =   49
      Top             =   3555
      Width           =   2580
   End
   Begin VB.TextBox txtMessage 
      Height          =   315
      Index           =   8
      Left            =   3105
      TabIndex        =   44
      Top             =   3195
      Width           =   2580
   End
   Begin VB.TextBox txtMessage 
      Height          =   315
      Index           =   7
      Left            =   3105
      TabIndex        =   39
      Top             =   2835
      Width           =   2580
   End
   Begin VB.TextBox txtMessage 
      Height          =   315
      Index           =   6
      Left            =   3105
      TabIndex        =   34
      Top             =   2475
      Width           =   2580
   End
   Begin VB.TextBox txtMessage 
      Height          =   315
      Index           =   5
      Left            =   3105
      TabIndex        =   29
      Top             =   2115
      Width           =   2580
   End
   Begin VB.TextBox txtMessage 
      Height          =   315
      Index           =   4
      Left            =   3105
      TabIndex        =   24
      Top             =   1755
      Width           =   2580
   End
   Begin VB.TextBox txtMessage 
      Height          =   315
      Index           =   3
      Left            =   3105
      TabIndex        =   19
      Top             =   1395
      Width           =   2580
   End
   Begin VB.TextBox txtMessage 
      Height          =   315
      Index           =   2
      Left            =   3105
      TabIndex        =   14
      Top             =   1035
      Width           =   2580
   End
   Begin VB.TextBox txtMessage 
      Height          =   315
      Index           =   1
      Left            =   3105
      TabIndex        =   9
      Top             =   675
      Width           =   2580
   End
   Begin VB.TextBox txtMessage 
      Height          =   285
      Index           =   0
      Left            =   3105
      TabIndex        =   4
      Top             =   335
      Width           =   2580
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   330
      Left            =   5040
      TabIndex        =   55
      Top             =   3960
      Width           =   1005
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Height          =   330
      Left            =   90
      TabIndex        =   54
      Top             =   3960
      Width           =   1005
   End
   Begin VB.ComboBox comboSound 
      Height          =   315
      Index           =   9
      ItemData        =   "frmKeylogger.frx":2806
      Left            =   2475
      List            =   "frmKeylogger.frx":2828
      Style           =   2  'Dropdown List
      TabIndex        =   48
      Top             =   3555
      Width           =   555
   End
   Begin VB.ComboBox comboSound 
      Height          =   315
      Index           =   8
      ItemData        =   "frmKeylogger.frx":284B
      Left            =   2475
      List            =   "frmKeylogger.frx":286D
      Style           =   2  'Dropdown List
      TabIndex        =   43
      Top             =   3195
      Width           =   555
   End
   Begin VB.ComboBox comboSound 
      Height          =   315
      Index           =   7
      ItemData        =   "frmKeylogger.frx":2890
      Left            =   2475
      List            =   "frmKeylogger.frx":28B2
      Style           =   2  'Dropdown List
      TabIndex        =   38
      Top             =   2835
      Width           =   555
   End
   Begin VB.ComboBox comboSound 
      Height          =   315
      Index           =   6
      ItemData        =   "frmKeylogger.frx":28D5
      Left            =   2475
      List            =   "frmKeylogger.frx":28F7
      Style           =   2  'Dropdown List
      TabIndex        =   33
      Top             =   2475
      Width           =   555
   End
   Begin VB.ComboBox comboSound 
      Height          =   315
      Index           =   5
      ItemData        =   "frmKeylogger.frx":291A
      Left            =   2475
      List            =   "frmKeylogger.frx":293C
      Style           =   2  'Dropdown List
      TabIndex        =   28
      Top             =   2115
      Width           =   555
   End
   Begin VB.ComboBox comboSound 
      Height          =   315
      Index           =   4
      ItemData        =   "frmKeylogger.frx":295F
      Left            =   2475
      List            =   "frmKeylogger.frx":2981
      Style           =   2  'Dropdown List
      TabIndex        =   23
      Top             =   1755
      Width           =   555
   End
   Begin VB.ComboBox comboSound 
      Height          =   315
      Index           =   3
      ItemData        =   "frmKeylogger.frx":29A4
      Left            =   2475
      List            =   "frmKeylogger.frx":29C6
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   1395
      Width           =   555
   End
   Begin VB.ComboBox comboSound 
      Height          =   315
      Index           =   2
      ItemData        =   "frmKeylogger.frx":29E9
      Left            =   2475
      List            =   "frmKeylogger.frx":2A0B
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   1035
      Width           =   555
   End
   Begin VB.ComboBox comboSound 
      Height          =   315
      Index           =   1
      ItemData        =   "frmKeylogger.frx":2A2E
      Left            =   2475
      List            =   "frmKeylogger.frx":2A50
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   675
      Width           =   555
   End
   Begin VB.ComboBox comboSound 
      Height          =   315
      Index           =   0
      ItemData        =   "frmKeylogger.frx":2A73
      Left            =   2475
      List            =   "frmKeylogger.frx":2A95
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   335
      Width           =   555
   End
   Begin VB.TextBox txtLen 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   9
      Left            =   1530
      TabIndex        =   47
      Text            =   "00:00:00"
      Top             =   3555
      Width           =   825
   End
   Begin VB.TextBox txtLen 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   8
      Left            =   1530
      TabIndex        =   42
      Text            =   "00:00:00"
      Top             =   3195
      Width           =   825
   End
   Begin VB.TextBox txtLen 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   7
      Left            =   1530
      TabIndex        =   37
      Text            =   "00:00:00"
      Top             =   2835
      Width           =   825
   End
   Begin VB.TextBox txtLen 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   6
      Left            =   1530
      TabIndex        =   32
      Text            =   "00:00:00"
      Top             =   2475
      Width           =   825
   End
   Begin VB.TextBox txtLen 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   5
      Left            =   1530
      TabIndex        =   27
      Text            =   "00:00:00"
      Top             =   2115
      Width           =   825
   End
   Begin VB.TextBox txtLen 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   4
      Left            =   1530
      TabIndex        =   22
      Text            =   "00:00:00"
      Top             =   1755
      Width           =   825
   End
   Begin VB.TextBox txtLen 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   3
      Left            =   1530
      TabIndex        =   17
      Text            =   "00:00:00"
      Top             =   1395
      Width           =   825
   End
   Begin VB.TextBox txtLen 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   2
      Left            =   1530
      TabIndex        =   12
      Text            =   "00:00:00"
      Top             =   1035
      Width           =   825
   End
   Begin VB.TextBox txtLen 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   1
      Left            =   1530
      TabIndex        =   7
      Text            =   "00:00:00"
      Top             =   675
      Width           =   825
   End
   Begin VB.TextBox txtLen 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   1530
      TabIndex        =   2
      Text            =   "00:00:00"
      Top             =   335
      Width           =   825
   End
   Begin VB.TextBox txtKey 
      Alignment       =   2  'Center
      Height          =   315
      Index           =   9
      Left            =   1035
      MaxLength       =   1
      TabIndex        =   46
      Top             =   3555
      Width           =   330
   End
   Begin VB.TextBox txtKey 
      Alignment       =   2  'Center
      Height          =   315
      Index           =   8
      Left            =   1035
      MaxLength       =   1
      TabIndex        =   41
      Top             =   3195
      Width           =   330
   End
   Begin VB.TextBox txtKey 
      Alignment       =   2  'Center
      Height          =   315
      Index           =   7
      Left            =   1035
      MaxLength       =   1
      TabIndex        =   36
      Top             =   2835
      Width           =   330
   End
   Begin VB.TextBox txtKey 
      Alignment       =   2  'Center
      Height          =   315
      Index           =   6
      Left            =   1035
      MaxLength       =   1
      TabIndex        =   31
      Top             =   2475
      Width           =   330
   End
   Begin VB.TextBox txtKey 
      Alignment       =   2  'Center
      Height          =   315
      Index           =   5
      Left            =   1035
      MaxLength       =   1
      TabIndex        =   26
      Top             =   2115
      Width           =   330
   End
   Begin VB.TextBox txtKey 
      Alignment       =   2  'Center
      Height          =   315
      Index           =   4
      Left            =   1035
      MaxLength       =   1
      TabIndex        =   21
      Top             =   1755
      Width           =   330
   End
   Begin VB.TextBox txtKey 
      Alignment       =   2  'Center
      Height          =   315
      Index           =   3
      Left            =   1035
      MaxLength       =   1
      TabIndex        =   16
      Top             =   1395
      Width           =   330
   End
   Begin VB.TextBox txtKey 
      Alignment       =   2  'Center
      Height          =   315
      Index           =   2
      Left            =   1035
      MaxLength       =   1
      TabIndex        =   11
      Top             =   1035
      Width           =   330
   End
   Begin VB.TextBox txtKey 
      Alignment       =   2  'Center
      Height          =   315
      Index           =   1
      Left            =   1035
      MaxLength       =   1
      TabIndex        =   6
      Top             =   675
      Width           =   330
   End
   Begin VB.TextBox txtKey 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   0
      Left            =   1035
      MaxLength       =   1
      TabIndex        =   1
      Top             =   335
      Width           =   330
   End
   Begin VB.ComboBox comboType 
      Height          =   315
      Index           =   9
      ItemData        =   "frmKeylogger.frx":2AB8
      Left            =   90
      List            =   "frmKeylogger.frx":2AC5
      Style           =   2  'Dropdown List
      TabIndex        =   45
      Top             =   3555
      Width           =   825
   End
   Begin VB.ComboBox comboType 
      Height          =   315
      Index           =   8
      ItemData        =   "frmKeylogger.frx":2ADB
      Left            =   90
      List            =   "frmKeylogger.frx":2AE8
      Style           =   2  'Dropdown List
      TabIndex        =   40
      Top             =   3195
      Width           =   825
   End
   Begin VB.ComboBox comboType 
      Height          =   315
      Index           =   7
      ItemData        =   "frmKeylogger.frx":2AFE
      Left            =   90
      List            =   "frmKeylogger.frx":2B0B
      Style           =   2  'Dropdown List
      TabIndex        =   35
      Top             =   2835
      Width           =   825
   End
   Begin VB.ComboBox comboType 
      Height          =   315
      Index           =   6
      ItemData        =   "frmKeylogger.frx":2B21
      Left            =   90
      List            =   "frmKeylogger.frx":2B2E
      Style           =   2  'Dropdown List
      TabIndex        =   30
      Top             =   2475
      Width           =   825
   End
   Begin VB.ComboBox comboType 
      Height          =   315
      Index           =   5
      ItemData        =   "frmKeylogger.frx":2B44
      Left            =   90
      List            =   "frmKeylogger.frx":2B51
      Style           =   2  'Dropdown List
      TabIndex        =   25
      Top             =   2115
      Width           =   825
   End
   Begin VB.ComboBox comboType 
      Height          =   315
      Index           =   4
      ItemData        =   "frmKeylogger.frx":2B67
      Left            =   90
      List            =   "frmKeylogger.frx":2B74
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   1755
      Width           =   825
   End
   Begin VB.ComboBox comboType 
      Height          =   315
      Index           =   3
      ItemData        =   "frmKeylogger.frx":2B8A
      Left            =   90
      List            =   "frmKeylogger.frx":2B97
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   1395
      Width           =   825
   End
   Begin VB.ComboBox comboType 
      Height          =   315
      Index           =   2
      ItemData        =   "frmKeylogger.frx":2BAD
      Left            =   90
      List            =   "frmKeylogger.frx":2BBA
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   1035
      Width           =   825
   End
   Begin VB.ComboBox comboType 
      Height          =   315
      Index           =   1
      ItemData        =   "frmKeylogger.frx":2BD0
      Left            =   90
      List            =   "frmKeylogger.frx":2BDD
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   675
      Width           =   825
   End
   Begin VB.ComboBox comboType 
      Height          =   315
      Index           =   0
      ItemData        =   "frmKeylogger.frx":2BF3
      Left            =   90
      List            =   "frmKeylogger.frx":2C00
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   335
      Width           =   825
   End
   Begin VB.Label lblHeader 
      Caption         =   "Description"
      Height          =   240
      Index           =   4
      Left            =   3105
      TabIndex        =   56
      Top             =   90
      Width           =   1545
   End
   Begin VB.Label lblHeader 
      Caption         =   "Sound"
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
      Height          =   240
      Index           =   3
      Left            =   2475
      MousePointer    =   10  'Up Arrow
      TabIndex        =   53
      Top             =   90
      Width           =   555
   End
   Begin VB.Label lblHeader 
      Caption         =   "Duration"
      Height          =   240
      Index           =   2
      Left            =   1530
      TabIndex        =   52
      Top             =   90
      Width           =   780
   End
   Begin VB.Label lblHeader 
      Caption         =   "Key"
      Height          =   240
      Index           =   1
      Left            =   1035
      TabIndex        =   51
      Top             =   90
      Width           =   330
   End
   Begin VB.Label lblHeader 
      Caption         =   "Type"
      Height          =   240
      Index           =   0
      Left            =   90
      TabIndex        =   50
      Top             =   90
      Width           =   780
   End
End
Attribute VB_Name = "frmKeylogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
Dim i As Integer
For i = 0 To UBound(KeyLogs)
    If txtKey(i).Text <> "" Then
        KeyLogs(i, 0) = comboType(i) & "-" & txtKey(i).Text
        KeyLogs(i, 1) = txtLen(i).Text
        KeyLogs(i, 2) = comboSound(i).Text
        KeyLogs(i, 3) = txtMessage(i).Text
        SaveSetting App.Title, "Keylogs", "Type/Key-" & i, KeyLogs(i, 0)
        SaveSetting App.Title, "Keylogs", "Duration-" & i, KeyLogs(i, 1)
        SaveSetting App.Title, "Keylogs", "Sound-" & i, KeyLogs(i, 2)
        SaveSetting App.Title, "Keylogs", "Message-" & i, KeyLogs(i, 3)
    End If
Next
Unload Me
End Sub


Private Sub cmdPlay_Click(Index As Integer)
PlaySound BeepSounds(comboSound(Index).ListIndex), 0&, SND_FILENAME Or SND_ASYNC
End Sub

Private Sub comboType_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim ShiftDown, AltDown, CtrlDown
ShiftDown = (Shift And vbShiftMask) > 0
AltDown = (Shift And vbAltMask) > 0
CtrlDown = (Shift And vbCtrlMask) > 0
If CtrlDown Then
    comboType(Index).ListIndex = 0
    KeyCode = 0
ElseIf AltDown Then
    comboType(Index).ListIndex = 1
    KeyCode = 0
ElseIf ShiftDown Then
    comboType(Index).ListIndex = 2
    KeyCode = 0
End If
End Sub


Private Sub Form_Load()
frmRead.Enabled = False
Me.Left = frmRead.Left + 100
Me.Top = frmRead.Top + 100
Dim i As Integer, MyPos As Integer
For i = 0 To UBound(KeyLogs)
    comboSound(i).ListIndex = 0
    comboType(i).ListIndex = 0
Next
For i = 0 To UBound(KeyLogs)
    If KeyLogs(i, 0) <> "" Then
        MyPos = InStr(1, KeyLogs(i, 0), "-")
        If Left$(KeyLogs(i, 0), MyPos - 1) = "CTRL" Or Left$(KeyLogs(i, 0), MyPos - 1) = "ALT" Or Left$(KeyLogs(i, 0), MyPos - 1) = "SHIFT" Then
            comboType(i).Text = Left$(KeyLogs(i, 0), MyPos - 1)
        End If
        If Right$(KeyLogs(i, 0), 1) <> "-" Then
            txtKey(i).Text = Right$(KeyLogs(i, 0), 1)
        End If
        txtLen(i).Text = KeyLogs(i, 1)
        If KeyLogs(i, 2) <> "" Then
            comboSound(i).Text = KeyLogs(i, 2)
        End If
        txtMessage(i).Text = KeyLogs(i, 3)
    End If
Next
End Sub


Private Sub Form_Unload(Cancel As Integer)
frmRead.Enabled = True
End Sub


Private Sub lblHeader_Click(Index As Integer)
frmBeep.Show
End Sub

Private Sub txtKey_Validate(Index As Integer, Cancel As Boolean)
txtKey(Index).Text = UCase(txtKey(Index).Text)
End Sub


Private Sub txtLen_GotFocus(Index As Integer)
txtLen(Index).SelStart = 0
txtLen(Index).SelLength = Len(txtLen(Index).Text)
End Sub

Private Sub txtLen_KeyPress(Index As Integer, KeyAscii As Integer)
If IsNumeric(CStr(Chr(KeyAscii))) = False And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
    Exit Sub
End If
If KeyAscii <> vbKeyBack Then
    If Len(txtLen(Index)) = 2 Then
        txtLen(Index).Text = txtLen(Index).Text & ":"
        txtLen(Index).SelStart = Len(txtLen(Index).Text)
    ElseIf Len(txtLen(Index)) = 5 Then
        txtLen(Index).Text = txtLen(Index).Text & ":"
        txtLen(Index).SelStart = Len(txtLen(Index).Text)
    End If
End If
End Sub


Private Sub txtLen_Validate(Index As Integer, Cancel As Boolean)
On Error Resume Next
If IsDate(Date & " " & txtLen(Index)) = False Then
    MsgBox "Invalid format, please use: " & vbNewLine & vbNewLine & "##:##:##", vbInformation, "Error"
End If
End Sub


