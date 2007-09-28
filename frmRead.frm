VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmRead 
   Caption         =   "FFXI Parser Online 6.1.0"
   ClientHeight    =   6390
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   7275
   Icon            =   "frmRead.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6390
   ScaleWidth      =   7275
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frameCraft 
      BorderStyle     =   0  'None
      Height          =   4650
      Left            =   60
      TabIndex        =   41
      Top             =   1320
      Visible         =   0   'False
      Width           =   7215
      Begin VB.ComboBox comboTime 
         Height          =   315
         ItemData        =   "frmRead.frx":1E72
         Left            =   4680
         List            =   "frmRead.frx":1E7F
         Style           =   2  'Dropdown List
         TabIndex        =   49
         Top             =   30
         Width           =   1215
      End
      Begin VB.CheckBox checkCraft 
         Caption         =   "Direction"
         Height          =   255
         Index           =   4
         Left            =   2400
         TabIndex        =   48
         Top             =   60
         Width           =   975
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Refresh"
         Height          =   315
         Left            =   5940
         TabIndex        =   47
         Top             =   30
         Width           =   1155
      End
      Begin VB.CheckBox checkCraft 
         Caption         =   "Time Interval"
         Height          =   255
         Index           =   3
         Left            =   3420
         TabIndex        =   46
         Top             =   60
         Width           =   1275
      End
      Begin VB.CheckBox checkCraft 
         Caption         =   "Moon %"
         Height          =   255
         Index           =   2
         Left            =   1440
         TabIndex        =   45
         Top             =   60
         Width           =   975
      End
      Begin VB.CheckBox checkCraft 
         Caption         =   "Moon"
         Height          =   255
         Index           =   1
         Left            =   660
         TabIndex        =   44
         Top             =   60
         Width           =   735
      End
      Begin VB.CheckBox checkCraft 
         Caption         =   "Day"
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   43
         Top             =   60
         Width           =   675
      End
      Begin RichTextLib.RichTextBox RTB_Crafting 
         Height          =   4230
         Left            =   0
         TabIndex        =   42
         TabStop         =   0   'False
         ToolTipText     =   "Double Click to Save"
         Top             =   420
         Width           =   7170
         _ExtentX        =   12647
         _ExtentY        =   7461
         _Version        =   393217
         BorderStyle     =   0
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         RightMargin     =   5.00000e5
         TextRTF         =   $"frmRead.frx":1E9E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame frameSummary 
      BorderStyle     =   0  'None
      Height          =   4650
      Left            =   60
      TabIndex        =   37
      Top             =   1320
      Visible         =   0   'False
      Width           =   7215
      Begin VB.OptionButton optionSummary 
         Caption         =   "Averages"
         Height          =   240
         Index           =   0
         Left            =   0
         TabIndex        =   39
         Top             =   45
         Value           =   -1  'True
         Width           =   1050
      End
      Begin VB.OptionButton optionSummary 
         Caption         =   "Report Summary"
         Height          =   240
         Index           =   1
         Left            =   1080
         TabIndex        =   38
         Top             =   45
         Width           =   1545
      End
      Begin RichTextLib.RichTextBox RTB_Averages 
         Height          =   4350
         Left            =   0
         TabIndex        =   40
         TabStop         =   0   'False
         ToolTipText     =   "Double Click to Save"
         Top             =   300
         Width           =   7170
         _ExtentX        =   12647
         _ExtentY        =   7673
         _Version        =   393217
         BorderStyle     =   0
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         RightMargin     =   5.00000e5
         TextRTF         =   $"frmRead.frx":1F15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame frameChat 
      BorderStyle     =   0  'None
      Height          =   4650
      Left            =   45
      TabIndex        =   17
      Top             =   1305
      Visible         =   0   'False
      Width           =   7215
      Begin VB.OptionButton optionChat 
         Caption         =   "Emotes"
         Height          =   240
         Index           =   6
         Left            =   3900
         TabIndex        =   36
         Top             =   45
         Width           =   825
      End
      Begin RichTextLib.RichTextBox RTB_Chat 
         Height          =   4290
         Left            =   0
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "Double Click to Save"
         Top             =   315
         Width           =   7170
         _ExtentX        =   12647
         _ExtentY        =   7567
         _Version        =   393217
         BorderStyle     =   0
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         RightMargin     =   5.00000e5
         TextRTF         =   $"frmRead.frx":1F8C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.OptionButton optionChat 
         Caption         =   "All"
         Height          =   240
         Index           =   5
         Left            =   4860
         TabIndex        =   24
         Top             =   45
         Width           =   1005
      End
      Begin VB.OptionButton optionChat 
         Caption         =   "LinkShell"
         Height          =   240
         Index           =   4
         Left            =   2835
         TabIndex        =   23
         Top             =   45
         Width           =   1005
      End
      Begin VB.OptionButton optionChat 
         Caption         =   "Party"
         Height          =   240
         Index           =   3
         Left            =   2115
         TabIndex        =   22
         Top             =   45
         Width           =   690
      End
      Begin VB.OptionButton optionChat 
         Caption         =   "Tell"
         Height          =   240
         Index           =   2
         Left            =   1440
         TabIndex        =   21
         Top             =   45
         Width           =   690
      End
      Begin VB.OptionButton optionChat 
         Caption         =   "Shout"
         Height          =   240
         Index           =   1
         Left            =   630
         TabIndex        =   20
         Top             =   45
         Width           =   825
      End
      Begin VB.OptionButton optionChat 
         Caption         =   "Say"
         Height          =   240
         Index           =   0
         Left            =   0
         TabIndex        =   19
         Top             =   45
         Width           =   690
      End
   End
   Begin VB.Timer timerAltHome 
      Interval        =   1
      Left            =   6360
      Top             =   6060
   End
   Begin VB.Frame frameSupport 
      Caption         =   "Support FFXIP!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1215
      Left            =   60
      TabIndex        =   33
      Top             =   60
      Width           =   7155
      Begin VB.Timer timerAd 
         Interval        =   5000
         Left            =   6960
         Top             =   480
      End
      Begin VB.Label lblHide 
         Alignment       =   1  'Right Justify
         Caption         =   "Hide"
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
         Left            =   6720
         MouseIcon       =   "frmRead.frx":200C
         MousePointer    =   99  'Custom
         TabIndex        =   35
         Top             =   0
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Image imgPaypal 
         Height          =   465
         Left            =   5760
         MouseIcon       =   "frmRead.frx":215E
         MousePointer    =   99  'Custom
         Picture         =   "frmRead.frx":22B0
         Top             =   420
         Width           =   930
      End
      Begin VB.Image imgB 
         Height          =   855
         Left            =   420
         Picture         =   "frmRead.frx":2615
         Top             =   1080
         Visible         =   0   'False
         Width           =   5055
      End
      Begin VB.Image imgA 
         Height          =   855
         Left            =   360
         Picture         =   "frmRead.frx":8961
         Top             =   1080
         Visible         =   0   'False
         Width           =   5055
      End
      Begin VB.Image imgAd 
         Height          =   855
         Left            =   120
         MouseIcon       =   "frmRead.frx":F126
         MousePointer    =   99  'Custom
         Picture         =   "frmRead.frx":F278
         Top             =   240
         Width           =   5055
      End
   End
   Begin VB.Timer timerKeyLogger 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   6795
      Top             =   1305
   End
   Begin VB.Timer timerBeeps 
      Interval        =   2500
      Left            =   6795
      Top             =   6075
   End
   Begin VB.ComboBox comboUser 
      Height          =   315
      ItemData        =   "frmRead.frx":15A3D
      Left            =   1800
      List            =   "frmRead.frx":15A3F
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   6030
      Width           =   1095
   End
   Begin VB.OptionButton optionResults 
      Height          =   195
      Index           =   0
      Left            =   90
      TabIndex        =   9
      ToolTipText     =   "Double click / Right click"
      Top             =   6075
      Value           =   -1  'True
      Width           =   240
   End
   Begin VB.ComboBox comboDisplay 
      Height          =   315
      ItemData        =   "frmRead.frx":15A41
      Left            =   360
      List            =   "frmRead.frx":15A60
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   6030
      Width           =   1095
   End
   Begin VB.OptionButton optionResults 
      Height          =   195
      Index           =   1
      Left            =   1530
      TabIndex        =   7
      Top             =   6090
      Width           =   240
   End
   Begin MSComDlg.CommonDialog CD_Save 
      Left            =   -45
      Top             =   1530
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   4455
      Top             =   1350
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer timerRead 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   4995
      Top             =   1395
   End
   Begin VB.ListBox fileListBox 
      Height          =   2595
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   6540
      Visible         =   0   'False
      Width           =   7170
   End
   Begin InetCtlsObjects.Inet inet 
      Left            =   6615
      Top             =   1665
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   4
      URL             =   "http://"
   End
   Begin RichTextLib.RichTextBox RTB_User 
      Height          =   4650
      Left            =   45
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "Double Click to Save"
      Top             =   1305
      Visible         =   0   'False
      Width           =   7170
      _ExtentX        =   12647
      _ExtentY        =   8202
      _Version        =   393217
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      RightMargin     =   5.00000e5
      TextRTF         =   $"frmRead.frx":15AAF
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox RTB 
      Height          =   2400
      Left            =   405
      TabIndex        =   5
      Top             =   2970
      Visible         =   0   'False
      Width           =   6315
      _ExtentX        =   11139
      _ExtentY        =   4233
      _Version        =   393217
      TextRTF         =   $"frmRead.frx":15B26
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox RTB_Report 
      Height          =   4560
      Left            =   45
      TabIndex        =   26
      TabStop         =   0   'False
      ToolTipText     =   "Double Click to Save"
      Top             =   1305
      Width           =   7170
      _ExtentX        =   12647
      _ExtentY        =   8043
      _Version        =   393217
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      RightMargin     =   5.00000e5
      TextRTF         =   $"frmRead.frx":15B9D
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.DirListBox dirList 
      Height          =   1890
      Left            =   45
      TabIndex        =   0
      Top             =   1305
      Visible         =   0   'False
      Width           =   2805
   End
   Begin VB.FileListBox fileList 
      Height          =   1455
      Left            =   90
      Pattern         =   "*.log"
      TabIndex        =   2
      Top             =   3330
      Visible         =   0   'False
      Width           =   1275
   End
   Begin RichTextLib.RichTextBox RTB_Fish 
      Height          =   4650
      Left            =   45
      TabIndex        =   25
      TabStop         =   0   'False
      ToolTipText     =   "Double Click to Save"
      Top             =   1305
      Visible         =   0   'False
      Width           =   7170
      _ExtentX        =   12647
      _ExtentY        =   8202
      _Version        =   393217
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      RightMargin     =   5.00000e5
      TextRTF         =   $"frmRead.frx":15C14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame frameEdit 
      Height          =   4650
      Left            =   45
      TabIndex        =   12
      Top             =   1305
      Width           =   7215
      Begin VB.CommandButton cmdUnselectPlayers 
         Caption         =   "Unselect All"
         Height          =   285
         Left            =   4770
         TabIndex        =   31
         Top             =   3375
         Width           =   1140
      End
      Begin VB.CommandButton cmdSelectPlayers 
         Caption         =   "Select All"
         Height          =   285
         Left            =   3600
         TabIndex        =   30
         Top             =   3375
         Width           =   1140
      End
      Begin VB.ListBox listPlayers 
         Height          =   2985
         Left            =   3600
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   27
         Top             =   360
         Width           =   2355
      End
      Begin VB.ComboBox comboMOB 
         Height          =   315
         ItemData        =   "frmRead.frx":15C8B
         Left            =   45
         List            =   "frmRead.frx":15C8D
         Sorted          =   -1  'True
         TabIndex        =   16
         ToolTipText     =   "Select/Type and hit Enter"
         Top             =   3870
         Width           =   2355
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "Select All"
         Height          =   285
         Left            =   45
         TabIndex        =   15
         Top             =   3555
         Width           =   1095
      End
      Begin VB.CommandButton cmdUnselect 
         Caption         =   "Unselect All"
         Height          =   285
         Left            =   1260
         TabIndex        =   14
         Top             =   3555
         Width           =   1095
      End
      Begin VB.ListBox listResults 
         Height          =   3180
         Left            =   45
         MultiSelect     =   2  'Extended
         TabIndex        =   13
         Top             =   360
         Width           =   2355
      End
      Begin VB.Label lbl 
         Caption         =   "Select/Type and hit Enter"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   2
         Left            =   2430
         TabIndex        =   32
         Top             =   3915
         Width           =   2130
      End
      Begin VB.Label lbl 
         Caption         =   "Players"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   3600
         TabIndex        =   29
         Top             =   135
         Width           =   1635
      End
      Begin VB.Label lbl 
         Caption         =   "Battles"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   45
         TabIndex        =   28
         Top             =   135
         Width           =   1635
      End
   End
   Begin RichTextLib.RichTextBox RTB_Details 
      Height          =   4650
      Left            =   45
      TabIndex        =   10
      TabStop         =   0   'False
      ToolTipText     =   "Double Click to Save"
      Top             =   1305
      Visible         =   0   'False
      Width           =   7170
      _ExtentX        =   12647
      _ExtentY        =   8202
      _Version        =   393217
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      RightMargin     =   5.00000e5
      TextRTF         =   $"frmRead.frx":15C8F
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox RTB_Log 
      Height          =   4650
      Left            =   45
      TabIndex        =   34
      TabStop         =   0   'False
      ToolTipText     =   "Double Click to Save"
      Top             =   1305
      Visible         =   0   'False
      Width           =   7170
      _ExtentX        =   12647
      _ExtentY        =   8202
      _Version        =   393217
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      RightMargin     =   5.00000e5
      TextRTF         =   $"frmRead.frx":15D06
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblStatus 
      Alignment       =   1  'Right Justify
      Caption         =   "Waiting."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3465
      TabIndex        =   4
      Top             =   6075
      Width           =   3705
   End
   Begin VB.Label Label1 
      Caption         =   "Status:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2970
      TabIndex        =   3
      Top             =   6075
      Width           =   690
   End
   Begin VB.Shape Shape1 
      Height          =   285
      Left            =   2925
      Top             =   6030
      Width           =   4290
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuParse 
         Caption         =   "&Start Parsing"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Parse Gathered Log"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSepa 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGather 
         Caption         =   "Gather Logs to &File"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuGatherDate 
         Caption         =   "Gather Logs to &Date"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuBothFile 
         Caption         =   "Parse/Gather Logs to F&ile"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuBoth 
         Caption         =   "Parse/Gather &Logs to Date"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuStopGathering 
         Caption         =   "Stop &Gathering"
         Enabled         =   0   'False
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuSpacer3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save Current Report"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveLog 
         Caption         =   "Save Data As..."
         Enabled         =   0   'False
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuSpacer4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMainExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuMainOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuLocation 
         Caption         =   "Set FFXI &Log Location"
      End
      Begin VB.Menu mnuFieldsSpacer 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReport 
         Caption         =   "&Select Reporting Fields"
      End
      Begin VB.Menu mnuSpacer41 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBeep 
         Caption         =   "Select Timer S&ounds"
      End
      Begin VB.Menu mnuEnableSounds 
         Caption         =   "Enable Timer Sounds"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuSpacer5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAltHome 
         Caption         =   "Enable Alt-Home Feature"
      End
      Begin VB.Menu mnuOnly 
         Caption         =   "&Read New Logs Only"
      End
      Begin VB.Menu mnuUpdate 
         Caption         =   "&Auto Update Check"
      End
      Begin VB.Menu mnuTray 
         Caption         =   "Show as &Tray Icon"
      End
      Begin VB.Menu mnuSpacer14 
         Caption         =   "-"
      End
      Begin VB.Menu mnuKey 
         Caption         =   "Setup Key Activated Timers"
      End
      Begin VB.Menu mnuKeyEnable 
         Caption         =   "Enable Key Activated Timers"
      End
      Begin VB.Menu mnuSpacer8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuParserCommands 
         Caption         =   "Enable Parser Commands"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuRecalculate 
         Caption         =   "&Recalculate"
      End
      Begin VB.Menu mnuSpacer9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExport 
         Caption         =   "&Export HTML"
      End
      Begin VB.Menu mnuExportXML 
         Caption         =   "Export &XML"
      End
      Begin VB.Menu mnuCSV 
         Caption         =   "Export CSV (Crafting)"
      End
      Begin VB.Menu mnuSpacerExp 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClear 
         Caption         =   "&Clear Data"
      End
   End
   Begin VB.Menu mnuOnline 
      Caption         =   "O&nline"
      Begin VB.Menu mnuOnlineSetup 
         Caption         =   "&Setup"
      End
      Begin VB.Menu mnuOnlineTransmit 
         Caption         =   "&Transmit"
      End
      Begin VB.Menu mnuOnlineView 
         Caption         =   "&View"
         Begin VB.Menu mnuDamage 
            Caption         =   "Damage Data"
         End
         Begin VB.Menu mnuCrafting 
            Caption         =   "Crafting Data"
         End
      End
   End
   Begin VB.Menu mnuViewMain 
      Caption         =   "&View"
      Begin VB.Menu mnuView 
         Caption         =   "&Report"
         Checked         =   -1  'True
         Index           =   0
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuView 
         Caption         =   "&Summary"
         Index           =   1
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuView 
         Caption         =   "&Details"
         Index           =   2
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuView 
         Caption         =   "&Chat"
         Index           =   3
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuView 
         Caption         =   "&Loot!"
         Index           =   4
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuView 
         Caption         =   "&Fishing"
         Index           =   5
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuView 
         Caption         =   "&Edit"
         Index           =   6
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuView 
         Caption         =   "&FFXIP Log"
         Index           =   7
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuView 
         Caption         =   "Craf&ting"
         Index           =   8
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnuSpacer7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewPlayer 
         Caption         =   "&1. Player1"
         Index           =   0
      End
      Begin VB.Menu mnuViewPlayer 
         Caption         =   "&2. Player2"
         Index           =   1
      End
      Begin VB.Menu mnuViewPlayer 
         Caption         =   "&3. Player3"
         Index           =   2
      End
      Begin VB.Menu mnuViewPlayer 
         Caption         =   "&4. Player4"
         Index           =   3
      End
      Begin VB.Menu mnuViewPlayer 
         Caption         =   "&5. Player5"
         Index           =   4
      End
      Begin VB.Menu mnuViewPlayer 
         Caption         =   "&6. Player6"
         Index           =   5
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuCommands 
         Caption         =   "&Parser Commands"
      End
      Begin VB.Menu mnuSpacer6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
   Begin VB.Menu mnuIcon 
      Caption         =   "mnuIcon"
      Visible         =   0   'False
      Begin VB.Menu mnuRestore 
         Caption         =   "Restore"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmRead"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim IGE As Boolean
Public BeepNotWave As Boolean, LastGather As Boolean

Dim CurrentDay As String, CurrentTime As String, CurrentMoon As String, CurrentPerc As String, PrevResult As String, CraftDirection As String


Dim HasErrors As Boolean, AltHome As String
Dim FishFound() As String, LootFound() As String, PlayerLoot() As String, ChatText() As String, TimerStart(10, 3) As String

Dim CurrentLine As String, CurrentFight As String, PrevLineA As String, PrevLineB As String
Dim ActiveLineType As Integer, PrevActiveLineType As Integer, LineType As String
Dim Attacker As String, Defender As String
Dim AttackerUses As String, AttackerSpecial As String, PreviousAttacker As String

Dim LastItem As String 'Last Log File
Dim ExpType As Integer

Dim ChainExp(20, 1) As Long
Dim DPS(17, 2) As String
Dim BattleID As Integer
Dim FightComment As String
Dim MonsterCheck As String
Dim TotalExp As Long
Dim ErrorCount As Integer
Dim timerBeepAmt As Integer, timerLength As String
Dim ReadDPS_Start As Boolean, ReadEXP_Start As Boolean, StopEXP As Boolean, BeginDPS As Boolean
Dim ReadDPS_Stop As Boolean, ReadEXP_Stop As Boolean, StopDPS As Boolean, ReadTimer As Boolean
Dim Read_Start As Boolean, Read_Stop As Boolean, ReadFISH_Start As Boolean

Dim StopLogging As Boolean, ParserCommands As Boolean

Dim FishHeader As String, FishComment As String

Dim StartTime As Date, StopTime As Date
Dim StartTimeDPS As Date, StopTimeDPS As Date
Dim FightStartTime As Date, FightStopTime As Date


Private WithEvents tIcon As TrayIcon
Attribute tIcon.VB_VarHelpID = -1

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, _
                                                                               ByVal lpOperation As String, _
                                                                               ByVal lpFile As String, _
                                                                               ByVal lpParameters As String, _
                                                                               ByVal lpDirectory As String, _
                                                                               ByVal nShowCmd As Long) _
                                                                               As Long


Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
    Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type
Private Const GENERIC_WRITE = &H40000000
Private Const OPEN_EXISTING = 3
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function SetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Private Declare Function SystemTimeToFileTime Lib "kernel32" (lpSystemTime As SYSTEMTIME, lpFileTime As FILETIME) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function LocalFileTimeToFileTime Lib "kernel32" (lpLocalFileTime As FILETIME, lpFileTime As FILETIME) As Long

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long
Private Declare Function ToAscii Lib "user32" (ByVal uVirtKey As Long, ByVal uScanCode As Long, lpbKeyState As Byte, lpwTransKey As Long, ByVal fuState As Long) As Long

Const SW_SHOWNORMAL = 1

Private Declare Sub Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long)











Private Sub AddBattleStats(i As Integer, NPCvsPC As Boolean)
10    On Error GoTo Err_Handler
      Dim DMG As Long, MyPos As Integer, FoundInstance As Boolean, dp As Integer, o
20    DMG = FindNumber
30    If NPCvsPC = False Then
40        FoundInstance = False
50        With BattleStats(i)
60            If ActiveLineType = 0 Then 'Basic Melee DMG
70                .TotalMeleeHit = .TotalMeleeHit + 1
80                .Basic.Hit = .Basic.Hit + 1
90                .Basic.Damage = .Basic.Damage + DMG
100               .Basic.List = .Basic.List & DMG & ", "
110               If .Basic.Low = 0 Then .Basic.Low = 10000
120               If .Basic.High < DMG Then
130                   .Basic.High = DMG
140               End If
150               If .Basic.Low > DMG And DMG <> 0 Then
160                   .Basic.Low = DMG
170               End If
180           ElseIf ActiveLineType = 4 Or ActiveLineType = 5 Then 'Basic Melee Miss
190               .Basic.Miss = .Basic.Miss + 1
200               .TotalMeleeMiss = .TotalMeleeMiss + 1
210           ElseIf ActiveLineType = 2 Then 'Ranged DMG
220               .TotalRangedHit = .TotalRangedHit + 1
230               .Ranged.Damage = .Ranged.Damage + DMG
240               .Ranged.Hit = .Ranged.Hit + 1
250               .Ranged.List = .Ranged.List & DMG & ", "
260               If .Ranged.Low = 0 Then .Ranged.Low = 10000
270               If .Ranged.High < DMG Then .Ranged.High = DMG
280               If .Ranged.Low > DMG And DMG <> 0 Then
290                   .Ranged.Low = DMG
300               End If
310           ElseIf ActiveLineType = 10 Then 'Ranged Miss
320               .Ranged.Miss = .Ranged.Miss + 1
330               .TotalRangedMiss = .TotalRangedMiss + 1
340           ElseIf ActiveLineType = 3 Then 'Counter DMG
350               .TotalMeleeHit = .TotalMeleeHit + 1
360               .Counter.Hit = .Counter.Hit + 1
370               .Counter.Damage = .Counter.Damage + DMG
380               .Counter.List = .Counter.List & DMG & ", "
390               If .Counter.Low = 0 Then .Counter.Low = 10000
400               If .Counter.High < DMG Then .Counter.High = DMG
410               If .Counter.Low > DMG And DMG <> 0 Then
420                   .Counter.Low = DMG
430               End If
440           ElseIf ActiveLineType = 19 Then 'Heals
450               .Heal.Healed = .Heal.Healed + DMG
460               .Heal.HealedList = .Heal.HealedList & DMG & ","
470               MyPos = InStr(3, CurrentLine, " recove")
480               Attacker = Defender
490               Defender = Mid$(CurrentLine, 3, MyPos - 3)
500               For o = 0 To UBound(SpellList)
510                   If AttackerSpecial = SpellList(o).Name Then
520                       .Heal.MPCost = .Heal.MPCost + SpellList(o).MPCost
530                       Exit For
540                   End If
550               Next
560           ElseIf ActiveLineType = 18 Then 'Weaponskill Miss
570               .TotalMeleeMiss = .TotalMeleeMiss + 1
580               .Basic.Miss = .Basic.Miss + 1
590               If InStr(1, SkillList, AttackerSpecial) Or InStr(1, .Attacker, "SC:") Or SkillList = "" Then
600                   .Skill.Miss = .Skill.Miss + 1
610                   .Skill.Uses = .Skill.Uses + 1
620               Else
630                   .Ability.Miss = .Ability.Miss + 1
640                   .Ability.Uses = .Ability.Uses + 1
650               End If
660           ElseIf ActiveLineType = 1 Then 'Additional Effect DMG
670               .Effect.Damage = .Effect.Damage + DMG
680               .Effect.List = .Effect.List & DMG & ", "
690               If .Effect.Low = 0 Then .Effect.Low = 10000
700               If .Effect.High < DMG Then .Effect.High = DMG
710               If .Effect.Low > DMG And DMG <> 0 Then
720                   .Effect.Low = DMG
730               End If
740           ElseIf PrevActiveLineType = 12 Or PrevActiveLineType = 14 Then 'Weaponskill DMG
750               .TotalMeleeHit = .TotalMeleeHit + 1
760               If InStr(1, SkillList, AttackerSpecial) Or InStr(1, .Attacker, "SC:") Or SkillList = "" Then
770                   .Skill.Hit = .Skill.Hit + 1
780                   .Skill.Damage = .Skill.Damage + DMG
790                   .Skill.Uses = .Skill.Uses + 1
800                   If InStr(1, .Attacker, "SC:") Then
810                       .Skill.List = .Skill.List & DMG & ", "
820                   Else
830                       .Skill.List = .Skill.List & DMG & " (" & AttackerSpecial & "), "
840                   End If
850                   If .Skill.Low = 0 Then .Skill.Low = 10000
860                   If .Skill.High < DMG Then
870                       .Skill.High = DMG
880                       .Skill.HighSkillType = AttackerSpecial
890                   End If
900                   If .Skill.Low > DMG And DMG <> 0 Then
910                       .Skill.Low = DMG
920                   End If
930               Else
940                   .Ability.Hit = .Ability.Hit + 1
950                   .Ability.Damage = .Ability.Damage + DMG
960                   .Ability.Uses = .Ability.Uses + 1
970                   .Ability.List = .Ability.List & DMG & " (" & AttackerSpecial & "), "
980                   If .Ability.Low = 0 Then .Ability.Low = 10000
990                   If .Ability.High < DMG Then
1000                      .Ability.High = DMG
1010                      .Ability.HighSkillType = AttackerSpecial
1020                  End If
1030                  If .Ability.Low > DMG And DMG <> 0 Then
1040                      .Ability.Low = DMG
1050                  End If
1060              End If
1070          ElseIf PrevActiveLineType = 13 Then 'Critical DMG
1080              .TotalMeleeHit = .TotalMeleeHit + 1
1090              .Critical.Hit = .Critical.Hit + 1
1100              .Critical.Damage = .Critical.Damage + DMG
1110              .Critical.List = .Critical.List & DMG & ", "
1120              If .Critical.Low = 0 Then .Critical.Low = 10000
1130              If .Critical.High < DMG Then .Critical.High = DMG
1140              If .Critical.Low > DMG And DMG <> 0 Then
1150                  .Critical.Low = DMG
1160              End If
1170          ElseIf PrevActiveLineType = 15 Then 'Spell DMG
1180              .Spell.Damage = .Spell.Damage + DMG
1190              .Spell.Uses = .Spell.Uses + 1
1200              .Spell.List = .Spell.List & DMG & " (" & AttackerSpecial & "), "
1210              If .Spell.Low = 0 Then .Spell.Low = 10000
1220              If .Spell.High < DMG Then
1230                  .Spell.High = DMG
1240                  .Spell.HighSkillType = AttackerSpecial
1250              End If
1260              If .Spell.Low > DMG And DMG <> 0 Then
1270                  .Spell.Low = DMG
1280              End If
1290              For o = 0 To UBound(SpellList)
1300                  If AttackerSpecial = SpellList(o).Name Then
1310                      .Spell.MPCost = .Spell.MPCost + SpellList(o).MPCost
1320                      Exit For
1330                  End If
1340              Next
1350          End If
1360          If ActiveLineType <> 19 Then
1370              .TotalDMG = .TotalDMG + DMG
1380              If StopDPS = False And BeginDPS = True Then
1390                  For dp = 0 To UBound(DPS)
1400                      If LCase(DPS(dp, 0)) = LCase(Attacker) Then
1410                          DPS(dp, 1) = CDbl(DPS(dp, 1)) + DMG
1420                          FoundInstance = True
1430                      End If
1440                  Next
1450                  If FoundInstance = False Then
1460                      For dp = 0 To UBound(DPS)
1470                          If DPS(dp, 0) = "" Then
1480                              DPS(dp, 0) = Attacker
1490                              DPS(dp, 1) = DMG
1500                              Exit For
1510                          End If
1520                      Next
1530                  End If
1540              End If
1550          End If
1560      End With
1570  Else
1580      With BattleStats(i)
1590          If ActiveLineType = 0 Or ActiveLineType = 2 Or ActiveLineType = 3 Or ActiveLineType = 12 Or ActiveLineType = 13 Or ActiveLineType = 15 Or ActiveLineType = 1 Then 'Hit
1600              .Evasion.Hit = .Evasion.Hit + 1
1610          ElseIf ActiveLineType = 4 Or ActiveLineType = 5 Or ActiveLineType = 10 Or ActiveLineType = 18 Then  'Miss
1620              .Evasion.TotalEvasion = .Evasion.TotalEvasion + 1
1630              .Evasion.Miss = .Evasion.Miss + 1
1640          ElseIf ActiveLineType = 19 Then
1650              .Heal.Recovered = .Heal.Recovered + DMG
1660              .Heal.RecoveredList = .Heal.RecoveredList & DMG & ","
1670          End If
1680          If ActiveLineType = 11 Then
1690              .Evasion.TotalEvasion = .Evasion.TotalEvasion + 1
1700              .Evasion.Evade = .Evasion.Evade + 1
1710          ElseIf ActiveLineType = 8 Then
1720              .Evasion.TotalEvasion = .Evasion.TotalEvasion + 1
1730              .Evasion.Absorb = .Evasion.Absorb + 1
1740          ElseIf ActiveLineType = 9 Then
1750              .Evasion.TotalEvasion = .Evasion.TotalEvasion + 1
1760              .Evasion.Anticipate = .Evasion.Anticipate + 1
1770          ElseIf ActiveLineType = 6 Then
1780              .Evasion.TotalEvasion = .Evasion.TotalEvasion + 1
1790              .Evasion.Parry = .Evasion.Parry + 1
1800          ElseIf ActiveLineType = 7 Then
1810              .Evasion.TotalEvasion = .Evasion.TotalEvasion + 1
1820              .Evasion.Block = .Evasion.Block + 1
1830          End If
1840          If ActiveLineType <> 19 Then
1850              .Evasion.Damage = .Evasion.Damage + DMG
1860          End If
1870      End With
1880  End If
1890  Exit Sub
Err_Handler:
1900  HasErrors = True
1910  ErrorCount = ErrorCount + 1
1920  ReportError = "Error: " & Err.Number & vbNewLine & "Source: AddBattleStats" & vbNewLine & "Description: " & Err.Description & vbNewLine & "Line: " & Erl & vbNewLine & "FFXI Log Line: " & CurrentLine & vbNewLine & "Previous FFXI Log Line: " & PrevLineA & vbNewLine & vbNewLine
1930  Print #ErrorFile, ReportError
1940  If ErrorCount >= 25 Then
1950      lblStatus.Caption = "Errors: " & HasErrors & " - Too many errors - Parsing stopped for this log."
1960      Exit Sub
1970  Else
1980      Resume Next
1990  End If
End Sub

Private Function AlphaNumber(AN As Integer) As String
Dim i As Integer
i = AN - 1
Dim str As String
str = ""

Do While Int(i / 26) >= 1
    str = Chr(97 + (i Mod 26)) & str
    i = ((i - (i Mod 26)) / 26) - 1
Loop
AlphaNumber = Chr(97 + i) & str
End Function





Private Sub ClearBattleStats(i As Integer, ClearComplete As Boolean)
10    On Error GoTo Err_Handler
      Dim EmptyStats As udtStatistics
20    If ClearComplete = False Then
30        BattleStats(i) = EmptyStats
40        With BattleStats(i)
50            .Basic.Low = 10000
60            .Ranged.Low = 10000
70            .Counter.Low = 10000
80            .Skill.Low = 10000
90            .Critical.Low = 10000
100           .Effect.Low = 10000
110           .Spell.Low = 10000
120       End With
130   Else
140       BattleTotals = EmptyStats
150       With BattleTotals
160           .Basic.Low = 10000
170           .Ranged.Low = 10000
180           .Counter.Low = 10000
190           .Skill.Low = 10000
200           .Critical.Low = 10000
210           .Effect.Low = 10000
220           .Spell.Low = 10000
230       End With
240   End If
250   Exit Sub
Err_Handler:
260   HasErrors = True
270   ErrorCount = ErrorCount + 1
280   ReportError = "Error: " & Err.Number & vbNewLine & "Source: ClearBattleStats" & vbNewLine & "Description: " & Err.Description & vbNewLine & "Line: " & Erl & vbNewLine & "FFXI Log Line: " & CurrentLine & vbNewLine & "Previous FFXI Log Line: " & PrevLineA & vbNewLine & vbNewLine
290   Print #ErrorFile, ReportError
300   If ErrorCount >= 25 Then
310       lblStatus.Caption = "Errors: " & HasErrors & " - Too many errors - Parsing stopped for this log."
320       Exit Sub
330   Else
340       Resume Next
350   End If
End Sub

Private Function ColumnText() As String
10    On Error GoTo Err_Handler
20    ColumnText = ResizePart("Player", 1000)
30    ColumnText = ColumnText & vbTab & ResizePart("Total", 525)

40    If ReportOptions(18) = 1 Then
50        ColumnText = ColumnText & vbTab & ResizePart("DMG %", 525)
60    End If
70    If ReportOptions(0) = 1 Then
80        ColumnText = ColumnText & vbTab & ResizePart("Melee", 525)
90    End If
100   If ReportOptions(24) = 1 Then
110       ColumnText = ColumnText & vbTab & ResizePart("Ranged", 525)
120   End If
130   If ReportOptions(2) = 1 Then
140       ColumnText = ColumnText & vbTab & ResizePart("Spell", 525)
150   End If
160   If ReportOptions(1) = 1 Then
170       ColumnText = ColumnText & vbTab & ResizePart("Skill", 525)
180   End If
190   If ReportOptions(37) = 1 Then
200       ColumnText = ColumnText & vbTab & ResizePart("Ability", 525)
210   End If
220   If ReportOptions(21) = 1 Then
230       ColumnText = ColumnText & vbTab & ResizePart("Effect", 525)
240   End If
250   If ReportOptions(7) = 1 Then
260       ColumnText = ColumnText & vbTab & ResizePart("M Hit %", 525)
270   End If
280   If ReportOptions(8) = 1 Then
290       ColumnText = ColumnText & vbTab & ResizePart("M Ht/Ms", 525)
300   End If
310   If ReportOptions(3) = 1 Then
320       ColumnText = ColumnText & vbTab & ResizePart("M Hi/Lo", 525)
330   End If
340   If ReportOptions(4) = 1 Then
350       ColumnText = ColumnText & vbTab & ResizePart("M Avg", 525)
360   End If
370   If ReportOptions(26) = 1 Then
380       ColumnText = ColumnText & vbTab & ResizePart("R Hit %", 525)
390   End If
400   If ReportOptions(25) = 1 Then
410       ColumnText = ColumnText & vbTab & ResizePart("R Ht/Ms", 525)
420   End If
430   If ReportOptions(27) = 1 Then
440       ColumnText = ColumnText & vbTab & ResizePart("R Hi/Lo", 525)
450   End If
460   If ReportOptions(28) = 1 Then
470       ColumnText = ColumnText & vbTab & ResizePart("R Avg", 525)
480   End If
490   If ReportOptions(30) = 1 Then
500       ColumnText = ColumnText & vbTab & ResizePart("Sp Hi/Lo", 525)
510   End If
520   If ReportOptions(29) = 1 Then
530       ColumnText = ColumnText & vbTab & ResizePart("Sp Avg", 525)
540   End If
550   If ReportOptions(36) = 1 Then
560       ColumnText = ColumnText & vbTab & ResizePart("Sp MP", 525)
570   End If
580   If ReportOptions(32) = 1 Then
590       ColumnText = ColumnText & vbTab & ResizePart("Sk Hi/Lo", 525)
600   End If
610   If ReportOptions(31) = 1 Then
620       ColumnText = ColumnText & vbTab & ResizePart("Sk Avg", 525)
630   End If
640   If ReportOptions(22) = 1 Then
650       ColumnText = ColumnText & vbTab & ResizePart("Sk #", 525)
660   End If
670   If ReportOptions(33) = 1 Then
680       ColumnText = ColumnText & vbTab & ResizePart("Ab Hi/Lo", 525)
690   End If
700   If ReportOptions(34) = 1 Then
710       ColumnText = ColumnText & vbTab & ResizePart("Ab Avg", 525)
720   End If
730   If ReportOptions(5) = 1 Then
740       ColumnText = ColumnText & vbTab & ResizePart("Crit %", 525)
750   End If
760   If ReportOptions(6) = 1 Then
770       ColumnText = ColumnText & vbTab & ResizePart("Crit #", 525)
780   End If
790   If ReportOptions(9) = 1 Then
800       ColumnText = ColumnText & vbTab & ResizePart("Avd %", 525)
810   End If
820   If ReportOptions(10) = 1 Then
830       ColumnText = ColumnText & vbTab & ResizePart("Tk/Av", 525)
840   End If
850   If ReportOptions(11) = 1 Then
860       ColumnText = ColumnText & vbTab & ResizePart("Evade", 525)
870   End If
880   If ReportOptions(12) = 1 Then
890       ColumnText = ColumnText & vbTab & ResizePart("Parry", 525)
900   End If
910   If ReportOptions(13) = 1 Then
920       ColumnText = ColumnText & vbTab & ResizePart("Block", 525)
930   End If
940   If ReportOptions(14) = 1 Then
950       ColumnText = ColumnText & vbTab & ResizePart("Absorb", 525)
960   End If
970   If ReportOptions(15) = 1 Then
980       ColumnText = ColumnText & vbTab & ResizePart("Avoid", 525)
990   End If
1000  If ReportOptions(20) = 1 Then
1010      ColumnText = ColumnText & vbTab & ResizePart("Anti", 525)
1020  End If
1030  If ReportOptions(23) = 1 Then
1040      ColumnText = ColumnText & vbTab & ResizePart("Cnter", 525)
1050  End If
1060  If ReportOptions(16) = 1 Then
1070      ColumnText = ColumnText & vbTab & ResizePart("Taken", 525)
1080  End If
1090  If ReportOptions(17) = 1 Then
1100      ColumnText = ColumnText & vbTab & ResizePart("Rcver", 525)
1110  End If
1120  If ReportOptions(19) = 1 Then
1130      ColumnText = ColumnText & vbTab & ResizePart("Heal", 525)
1140  End If
1150  If ReportOptions(35) = 1 Then
1160      ColumnText = ColumnText & vbTab & ResizePart("HP MP", 525)
1170  End If
1180  ColumnText = ColumnText & vbTab

1190  Exit Function
Err_Handler:
1200  HasErrors = True
1210  ErrorCount = ErrorCount + 1
1220  ReportError = "Error: " & Err.Number & vbNewLine & "Source: ColumnText" & vbNewLine & "Description: " & Err.Description & vbNewLine & "Line: " & Erl & vbNewLine & "FFXI Log Line: " & CurrentLine & vbNewLine & "Previous FFXI Log Line: " & PrevLineA & vbNewLine & vbNewLine
1230  Print #ErrorFile, ReportError
1240  If ErrorCount >= 25 Then
1250      lblStatus.Caption = "Errors: " & HasErrors & " - Too many errors - Parsing stopped for this log."
1260      Exit Function
1270  Else
1280      Resume Next
1290  End If
End Function
Private Sub CombineStats(MainStats As udtStatistics, AddStats As udtStatistics)
10    On Error GoTo Err_Handler
20    With MainStats
40        .BattleID = AddStats.BattleID
50        .Attacker = AddStats.Attacker
60        .Defender = AddStats.Defender
70        .Basic.Hit = .Basic.Hit + AddStats.Basic.Hit
80        .Basic.Miss = .Basic.Miss + AddStats.Basic.Miss
90        .Basic.Damage = .Basic.Damage + AddStats.Basic.Damage
100       If AddStats.Basic.High > .Basic.High Then
110           .Basic.High = AddStats.Basic.High
120       End If
130       If .Basic.Low = 0 Then .Basic.Low = 10000
140       If AddStats.Basic.Low < .Basic.Low And AddStats.Basic.Low <> 0 Then
150           .Basic.Low = AddStats.Basic.Low
160       End If
170       .Ranged.Hit = .Ranged.Hit + AddStats.Ranged.Hit
180       .Ranged.Miss = .Ranged.Miss + AddStats.Ranged.Miss
190       .Ranged.Damage = .Ranged.Damage + AddStats.Ranged.Damage
200       If AddStats.Ranged.High > .Ranged.High Then
210           .Ranged.High = AddStats.Ranged.High
220       End If
230       If .Ranged.Low = 0 Then .Ranged.Low = 10000
240       If AddStats.Ranged.Low < .Ranged.Low Then
250           .Ranged.Low = AddStats.Ranged.Low
260       End If
270       .Counter.Hit = .Counter.Hit + AddStats.Counter.Hit
280       .Counter.Damage = .Counter.Damage + AddStats.Counter.Damage
290       If AddStats.Counter.High > .Counter.High Then
300           .Counter.High = AddStats.Counter.High
310       End If
320       If .Counter.Low = 0 Then .Counter.Low = 10000
330       If AddStats.Counter.Low < .Counter.Low Then
340           .Counter.Low = AddStats.Counter.Low
350       End If
360       .Skill.Hit = .Skill.Hit + AddStats.Skill.Hit
370       .Skill.Miss = .Skill.Miss + AddStats.Skill.Miss
380       .Skill.Uses = .Skill.Uses + AddStats.Skill.Uses
390       .Skill.Damage = .Skill.Damage + AddStats.Skill.Damage
400       If AddStats.Skill.High > .Skill.High Then
410           .Skill.High = AddStats.Skill.High
420           .Skill.HighSkillType = AddStats.Skill.HighSkillType
430       End If
440       If .Skill.Low = 0 Then .Skill.Low = 10000
450       If AddStats.Skill.Low < .Skill.Low Then
460           .Skill.Low = AddStats.Skill.Low
470       End If
    
    
480       .Ability.Hit = .Ability.Hit + AddStats.Ability.Hit
490       .Ability.Miss = .Ability.Miss + AddStats.Ability.Miss
500       .Ability.Uses = .Ability.Uses + AddStats.Ability.Uses
510       .Ability.Damage = .Ability.Damage + AddStats.Ability.Damage
520       If AddStats.Ability.High > .Ability.High Then
530           .Ability.High = AddStats.Ability.High
540           .Ability.HighSkillType = AddStats.Ability.HighSkillType
550       End If
560       If .Ability.Low = 0 Then .Ability.Low = 10000
570       If AddStats.Ability.Low < .Ability.Low Then
580           .Ability.Low = AddStats.Ability.Low
590       End If
    
600       .Critical.Hit = .Critical.Hit + AddStats.Critical.Hit
610       .Critical.Damage = .Critical.Damage + AddStats.Critical.Damage
620       If AddStats.Critical.High > .Critical.High Then
630           .Critical.High = AddStats.Critical.High
640       End If
650       If .Critical.Low = 0 Then .Critical.Low = 10000
660       If AddStats.Critical.Low < .Critical.Low Then
670           .Critical.Low = AddStats.Critical.Low
680       End If
690       .TotalMeleeHit = .TotalMeleeHit + AddStats.TotalMeleeHit
700       .TotalMeleeMiss = .TotalMeleeMiss + AddStats.TotalMeleeMiss
710       .TotalRangedHit = .TotalRangedHit + AddStats.TotalRangedHit
720       .TotalRangedMiss = .TotalRangedMiss + AddStats.TotalRangedMiss
730       .Spell.Damage = .Spell.Damage + AddStats.Spell.Damage
740       .Spell.Uses = .Spell.Uses + AddStats.Spell.Uses
750       .Spell.MPCost = .Spell.MPCost + AddStats.Spell.MPCost
760       If AddStats.Spell.High > .Spell.High Then
770           .Spell.High = AddStats.Spell.High
780           .Spell.HighSkillType = AddStats.Spell.HighSkillType
790       End If
800       If .Spell.Low = 0 Then .Spell.Low = 10000
810       If AddStats.Spell.Low < .Spell.Low Then
820           .Spell.Low = AddStats.Spell.Low
830       End If
840       .Effect.Damage = .Effect.Damage + AddStats.Effect.Damage
850       If AddStats.Effect.High > .Effect.High Then
860           .Effect.High = AddStats.Effect.High
870       End If
880       If .Effect.Low = 0 Then .Effect.Low = 10000
890       If AddStats.Effect.Low < .Effect.Low Then
900           .Effect.Low = AddStats.Effect.Low
910       End If
920       .TotalDMG = .TotalDMG + AddStats.TotalDMG
930       .Evasion.Damage = .Evasion.Damage + AddStats.Evasion.Damage
940       .Evasion.Parry = .Evasion.Parry + AddStats.Evasion.Parry
950       .Evasion.Block = .Evasion.Block + AddStats.Evasion.Block
960       .Evasion.Absorb = .Evasion.Absorb + AddStats.Evasion.Absorb
970       .Evasion.Anticipate = .Evasion.Anticipate + AddStats.Evasion.Anticipate
980       .Evasion.Evade = .Evasion.Evade + AddStats.Evasion.Evade
990       .Evasion.Miss = .Evasion.Miss + AddStats.Evasion.Miss
1000      .Evasion.Hit = .Evasion.Hit + AddStats.Evasion.Hit
1010      .Evasion.TotalEvasion = .Evasion.TotalEvasion + AddStats.Evasion.TotalEvasion
1020      .Heal.Healed = .Heal.Healed + AddStats.Heal.Healed
1030      .Heal.Recovered = .Heal.Recovered + AddStats.Heal.Recovered
1040      .Heal.MPCost = .Heal.MPCost + AddStats.Heal.MPCost
1050  End With

1060  Exit Sub
Err_Handler:
1070  HasErrors = True
1080  ErrorCount = ErrorCount + 1
1090  ReportError = "Error: " & Err.Number & vbNewLine & "Source: CombineStats" & vbNewLine & "Description: " & Err.Description & vbNewLine & "Line: " & Erl & vbNewLine & "FFXI Log Line: " & CurrentLine & vbNewLine & "Previous FFXI Log Line: " & PrevLineA & vbNewLine & vbNewLine
1100  Print #ErrorFile, ReportError
1110  If ErrorCount >= 25 Then
1120      lblStatus.Caption = "Errors: " & HasErrors & " - Too many errors - Parsing stopped for this log."
1130      Exit Sub
1140  Else
1150      Resume Next
1160  End If
End Sub

Private Sub ExportXML(FileName As String)
10    On Error GoTo Err_Handler
20    Screen.MousePointer = vbHourglass
      Dim XMLFile, XMLCode As String
      Dim XMLSummaryStats() As udtStatistics, FoundSummary As Boolean, IncludePlayer As Boolean
      Dim SummaryTotals As udtStatistics, EmptyStats As udtStatistics
      Dim i As Integer, p As Integer, c As Integer
30    ReDim XMLSummaryStats(0)

40    XMLFile = FreeFile
50    Open FileName For Output As XMLFile
    
60    For i = 0 To UBound(FullStats) - 1
70        If listResults.Selected(FullStats(i).BattleID) = True Then
80            FoundSummary = False
90            For p = 0 To UBound(XMLSummaryStats)
100               IncludePlayer = False
110               For c = 0 To listPlayers.ListCount - 1
120                   If listPlayers.List(c) = FullStats(i).Attacker Or InStr(1, FullStats(i).Attacker, "SC:") <> 0 Then
130                       If listPlayers.Selected(c) = True Then
140                           IncludePlayer = True
150                       End If
160                       Exit For
170                   End If
180               Next
190               If FullStats(i).Attacker = XMLSummaryStats(p).Attacker And FullStats(i).BattleID = XMLSummaryStats(p).BattleID And IncludePlayer Then
200                   CombineStats XMLSummaryStats(p), FullStats(i)
210                   FoundSummary = True
220               End If
230           Next
240           If FoundSummary = False Then
250               CombineStats XMLSummaryStats(UBound(XMLSummaryStats)), FullStats(i)
260               ReDim Preserve XMLSummaryStats(UBound(XMLSummaryStats) + 1)
270           End If
280       End If
290   Next

300   For i = 0 To UBound(XMLSummaryStats) - 1
310       CombineStats SummaryTotals, XMLSummaryStats(i)
320   Next

      Dim CurrentBattleID As Long
330   XMLCode = XMLCode & "<?xml version=" & """" & "1.0" & """" & " encoding=" & """" & "ISO-8859-1" & """" & "?>" & vbNewLine
340   XMLCode = XMLCode & "<!-- Created by FFXIP -->" & vbNewLine
350   XMLCode = XMLCode & "<DATA>" & vbNewLine
360   For i = 0 To UBound(XMLSummaryStats) - 1
370       With XMLSummaryStats(i)
380           CurrentBattleID = .BattleID
390           XMLCode = XMLCode & vbTab & "<BATTLE>" & vbNewLine
400           XMLCode = XMLCode & vbTab & "<ID>" & .BattleID & "</ID>" & vbNewLine
410           XMLCode = XMLCode & vbTab & "<ENEMY>" & .Defender & "</ENEMY>" & vbNewLine
420       End With
430       Do Until XMLSummaryStats(i).BattleID <> CurrentBattleID
440           With XMLSummaryStats(i)
450               XMLCode = XMLCode & vbTab & "<PLAYER>" & vbNewLine
460               XMLCode = XMLCode & vbTab & vbTab & "<NAME>" & .Attacker & "</NAME>" & vbNewLine

470               If .Basic.Hit <> 0 Or .Basic.Miss <> 0 Then
480                   XMLCode = XMLCode & vbTab & vbTab & "<BASIC>" & vbNewLine
490                   XMLCode = XMLCode & vbTab & vbTab & vbTab & "<HIT>" & .Basic.Hit & "</HIT>" & vbNewLine
500                   XMLCode = XMLCode & vbTab & vbTab & vbTab & "<MISS>" & .Basic.Miss & "</MISS>" & vbNewLine
510                   XMLCode = XMLCode & vbTab & vbTab & vbTab & "<DAMAGE>" & .Basic.Damage & "</DAMAGE>" & vbNewLine
520                   XMLCode = XMLCode & vbTab & vbTab & vbTab & "<HIGH>" & .Basic.High & "</HIGH>" & vbNewLine
530                   If .Basic.Low = 10000 Then .Basic.Low = 0
540                   XMLCode = XMLCode & vbTab & vbTab & vbTab & "<LOW>" & .Basic.Low & "</LOW>" & vbNewLine
550                   XMLCode = XMLCode & vbTab & vbTab & vbTab & "<USES>" & .Basic.Uses & "</USES>" & vbNewLine
560                   XMLCode = XMLCode & vbTab & vbTab & "</BASIC>" & vbNewLine
570               End If

580               If .Ranged.Hit <> 0 Or .Ranged.Miss <> 0 Then
590                   XMLCode = XMLCode & vbTab & vbTab & "<RANGED>" & vbNewLine
600                   XMLCode = XMLCode & vbTab & vbTab & vbTab & "<HIT>" & .Ranged.Hit & "</HIT>" & vbNewLine
610                   XMLCode = XMLCode & vbTab & vbTab & vbTab & "<MISS>" & .Ranged.Miss & "</MISS>" & vbNewLine
620                   XMLCode = XMLCode & vbTab & vbTab & vbTab & "<DAMAGE>" & .Ranged.Damage & "</DAMAGE>" & vbNewLine
630                   XMLCode = XMLCode & vbTab & vbTab & vbTab & "<HIGH>" & .Ranged.High & "</HIGH>" & vbNewLine
640                   XMLCode = XMLCode & vbTab & vbTab & vbTab & "<LOW>" & .Ranged.Low & "</LOW>" & vbNewLine
650                   If .Ranged.Low = 10000 Then .Ranged.Low = 0
660                   XMLCode = XMLCode & vbTab & vbTab & vbTab & "<USES>" & .Ranged.Uses & "</USES>" & vbNewLine
670                   XMLCode = XMLCode & vbTab & vbTab & "</RANGED>" & vbNewLine
680               End If

690               If .Counter.Damage <> 0 Then
700                   XMLCode = XMLCode & vbTab & vbTab & "<COUNTER>" & vbNewLine
710                   XMLCode = XMLCode & vbTab & vbTab & vbTab & "<HIT>" & .Counter.Hit & "</HIT>" & vbNewLine
720                   XMLCode = XMLCode & vbTab & vbTab & vbTab & "<MISS>" & .Counter.Miss & "</MISS>" & vbNewLine
730                   XMLCode = XMLCode & vbTab & vbTab & vbTab & "<DAMAGE>" & .Counter.Damage & "</DAMAGE>" & vbNewLine
740                   XMLCode = XMLCode & vbTab & vbTab & vbTab & "<HIGH>" & .Counter.High & "</HIGH>" & vbNewLine
750                   XMLCode = XMLCode & vbTab & vbTab & vbTab & "<LOW>" & .Counter.Low & "</LOW>" & vbNewLine
760                   If .Counter.Low = 10000 Then .Counter.Low = 0
770                   XMLCode = XMLCode & vbTab & vbTab & vbTab & "<USES>" & .Counter.Uses & "</USES>" & vbNewLine
780                   XMLCode = XMLCode & vbTab & vbTab & "</COUNTER>" & vbNewLine
790               End If

800               If .Skill.Uses <> 0 Then
810                   XMLCode = XMLCode & vbTab & vbTab & "<SKILL>" & vbNewLine
820                   XMLCode = XMLCode & vbTab & vbTab & vbTab & "<HIT>" & .Skill.Hit & "</HIT>" & vbNewLine
830                   XMLCode = XMLCode & vbTab & vbTab & vbTab & "<MISS>" & .Skill.Miss & "</MISS>" & vbNewLine
840                   XMLCode = XMLCode & vbTab & vbTab & vbTab & "<DAMAGE>" & .Skill.Damage & "</DAMAGE>" & vbNewLine
850                   XMLCode = XMLCode & vbTab & vbTab & vbTab & "<HIGH>" & .Skill.High & "</HIGH>" & vbNewLine
860                   XMLCode = XMLCode & vbTab & vbTab & vbTab & "<LOW>" & .Skill.Low & "</LOW>" & vbNewLine
870                   If .Skill.Low = 10000 Then .Skill.Low = 0
880                   XMLCode = XMLCode & vbTab & vbTab & vbTab & "<USES>" & .Skill.Uses & "</USES>" & vbNewLine
890                   XMLCode = XMLCode & vbTab & vbTab & "</SKILL>" & vbNewLine
900               End If

910               If .Critical.Hit <> 0 Then
920                   XMLCode = XMLCode & vbTab & vbTab & "<CRITICAL>" & vbNewLine
930                   XMLCode = XMLCode & vbTab & vbTab & vbTab & "<HIT>" & .Critical.Hit & "</HIT>" & vbNewLine
940                   XMLCode = XMLCode & vbTab & vbTab & vbTab & "<MISS>" & .Critical.Miss & "</MISS>" & vbNewLine
950                   XMLCode = XMLCode & vbTab & vbTab & vbTab & "<DAMAGE>" & .Critical.Damage & "</DAMAGE>" & vbNewLine
960                   XMLCode = XMLCode & vbTab & vbTab & vbTab & "<HIGH>" & .Critical.High & "</HIGH>" & vbNewLine
970                   XMLCode = XMLCode & vbTab & vbTab & vbTab & "<LOW>" & .Critical.Low & "</LOW>" & vbNewLine
980                   If .Critical.Low = 10000 Then .Critical.Low = 0
990                   XMLCode = XMLCode & vbTab & vbTab & vbTab & "<USES>" & .Critical.Uses & "</USES>" & vbNewLine
1000                  XMLCode = XMLCode & vbTab & vbTab & "</CRITICAL>" & vbNewLine
1010              End If

1020              If .Effect.Hit <> 0 Or .Effect.Damage <> 0 Then
1030                  XMLCode = XMLCode & vbTab & vbTab & "<EFFECT>" & vbNewLine
1040                  XMLCode = XMLCode & vbTab & vbTab & vbTab & "<HIT>" & .Effect.Hit & "</HIT>" & vbNewLine
1050                  XMLCode = XMLCode & vbTab & vbTab & vbTab & "<MISS>" & .Effect.Miss & "</MISS>" & vbNewLine
1060                  XMLCode = XMLCode & vbTab & vbTab & vbTab & "<DAMAGE>" & .Effect.Damage & "</DAMAGE>" & vbNewLine
1070                  XMLCode = XMLCode & vbTab & vbTab & vbTab & "<HIGH>" & .Effect.High & "</HIGH>" & vbNewLine
1080                  XMLCode = XMLCode & vbTab & vbTab & vbTab & "<LOW>" & .Effect.Low & "</LOW>" & vbNewLine
1090                  If .Effect.Low = 10000 Then .Effect.Low = 0
1100                  XMLCode = XMLCode & vbTab & vbTab & vbTab & "<USES>" & .Effect.Uses & "</USES>" & vbNewLine
1110                  XMLCode = XMLCode & vbTab & vbTab & "</EFFECT>" & vbNewLine
1120              End If

1130              If .Spell.Damage <> 0 Or .Spell.Hit <> 0 Then
1140                  XMLCode = XMLCode & vbTab & vbTab & "<SPELL>" & vbNewLine
1150                  XMLCode = XMLCode & vbTab & vbTab & vbTab & "<HIT>" & .Spell.Hit & "</HIT>" & vbNewLine
1160                  XMLCode = XMLCode & vbTab & vbTab & vbTab & "<MISS>" & .Spell.Miss & "</MISS>" & vbNewLine
1170                  XMLCode = XMLCode & vbTab & vbTab & vbTab & "<DAMAGE>" & .Spell.Damage & "</DAMAGE>" & vbNewLine
1180                  XMLCode = XMLCode & vbTab & vbTab & vbTab & "<HIGH>" & .Spell.High & "</HIGH>" & vbNewLine
1190                  XMLCode = XMLCode & vbTab & vbTab & vbTab & "<LOW>" & .Spell.Low & "</LOW>" & vbNewLine
1200                  If .Spell.Low = 10000 Then .Spell.Low = 0
1210                  XMLCode = XMLCode & vbTab & vbTab & vbTab & "<USES>" & .Spell.Uses & "</USES>" & vbNewLine
1220                  XMLCode = XMLCode & vbTab & vbTab & "</SPELL>" & vbNewLine
1230              End If

1240              If .Evasion.TotalEvasion <> 0 Or .Evasion.Damage <> 0 Or .Evasion.Hit <> 0 Then
1250                  XMLCode = XMLCode & vbTab & vbTab & "<EVASION>" & vbNewLine
1260                  XMLCode = XMLCode & vbTab & vbTab & vbTab & "<HIT>" & .Evasion.Hit & "</HIT>" & vbNewLine
1270                  XMLCode = XMLCode & vbTab & vbTab & vbTab & "<MISS>" & .Evasion.Miss & "</MISS>" & vbNewLine
1280                  XMLCode = XMLCode & vbTab & vbTab & vbTab & "<DAMAGE>" & .Evasion.Damage & "</DAMAGE>" & vbNewLine
1290                  XMLCode = XMLCode & vbTab & vbTab & vbTab & "<ABSORB>" & .Evasion.Absorb & "</ABSORB>" & vbNewLine
1300                  XMLCode = XMLCode & vbTab & vbTab & vbTab & "<ANTICIPATE>" & .Evasion.Anticipate & "</ANTICIPATE>" & vbNewLine
1310                  XMLCode = XMLCode & vbTab & vbTab & vbTab & "<BLOCK>" & .Evasion.Block & "</BLOCK>" & vbNewLine
1320                  XMLCode = XMLCode & vbTab & vbTab & vbTab & "<EVADE>" & .Evasion.Evade & "</EVADE>" & vbNewLine
1330                  XMLCode = XMLCode & vbTab & vbTab & vbTab & "<PARRY>" & .Evasion.Parry & "</PARRY>" & vbNewLine
1340                  XMLCode = XMLCode & vbTab & vbTab & "</EVASION>" & vbNewLine
1350              End If

1360              If .Heal.Healed <> 0 Or .Heal.Recovered <> 0 Then
1370                  XMLCode = XMLCode & vbTab & vbTab & "<HEAL>" & vbNewLine
1380                  XMLCode = XMLCode & vbTab & vbTab & vbTab & "<HEALED>" & .Heal.Healed & "</HEALED>" & vbNewLine
1390                  XMLCode = XMLCode & vbTab & vbTab & vbTab & "<RECOVERED>" & .Heal.Recovered & "</RECOVERED>" & vbNewLine
1400                  XMLCode = XMLCode & vbTab & vbTab & "</HEAL>" & vbNewLine
1410              End If

1420              XMLCode = XMLCode & vbTab & "</PLAYER>" & vbNewLine
1430              i = i + 1
1440          End With
1450      Loop
1460      i = i - 1
1470      XMLCode = XMLCode & vbTab & "</BATTLE>" & vbNewLine
1480  Next

1490  XMLCode = XMLCode & "</DATA>" & vbNewLine

1500  Print #XMLFile, XMLCode
1510  Close #XMLFile
1520  Screen.MousePointer = vbDefault
1530  Exit Sub
Err_Handler:
1540  HasErrors = True
1550  ErrorCount = ErrorCount + 1
1560  ReportError = "Error: " & Err.Number & vbNewLine & "Source: ExportXML" & vbNewLine & "Description: " & Err.Description & vbNewLine & "Line: " & Erl & vbNewLine & "FFXI Log Line: " & CurrentLine & vbNewLine & "Previous FFXI Log Line: " & PrevLineA & vbNewLine & vbNewLine
1570  Print #ErrorFile, ReportError
1580  If ErrorCount >= 25 Then
1590      lblStatus.Caption = "Errors: " & HasErrors & " - Too many errors - Parsing stopped for this log."
1600      Exit Sub
1610  Else
1620      Resume Next
1630  End If
End Sub

Private Sub ExportCSV(FileName As String)
10    On Error GoTo Err_Handler
20    Screen.MousePointer = vbHourglass
      Dim CSVFile, CSVCode As String, i

30    CSVFile = FreeFile
40    Open FileName For Output As CSVFile
    
50    For i = 0 To UBound(CraftingCSV)
60        CSVCode = CSVCode & CraftingCSV(i).Count & "," & CraftingCSV(i).Result & "," & CraftingCSV(i).DayType & "," & CraftingCSV(i).MoonPhase & "," & CraftingCSV(i).MoonPerc & "," & CraftingCSV(i).CurrentTime & "," & CraftingCSV(i).Direction & vbNewLine
70    Next
80    Print #CSVFile, CSVCode
90    Close #CSVFile
100   Screen.MousePointer = vbDefault
110   Exit Sub
Err_Handler:
120   HasErrors = True
130   ErrorCount = ErrorCount + 1
140   ReportError = "Error: " & Err.Number & vbNewLine & "Source: ExportCSV" & vbNewLine & "Description: " & Err.Description & vbNewLine & "Line: " & Erl & vbNewLine & "FFXI Log Line: " & CurrentLine & vbNewLine & "Previous FFXI Log Line: " & PrevLineA & vbNewLine & vbNewLine
150   Print #ErrorFile, ReportError
160   If ErrorCount >= 25 Then
170       lblStatus.Caption = "Errors: " & HasErrors & " - Too many errors - Parsing stopped for this log."
180       Exit Sub
190   Else
200       Resume Next
210   End If
End Sub


Private Function FooterText(SummaryArray As udtStatistics, UserRPT As Boolean) As String
10    On Error GoTo Err_Handler
      Dim dHigh As Integer, dLow As Integer
20    With SummaryArray
30        .Defender = .Defender
40        FooterText = ResizePart("Totals", 1000)
50        FooterText = FooterText & vbTab & ResizePart(CStr(.TotalDMG), 525)
60        If ReportOptions(18) = 1 Then 'DAMAGE PERCENT
70            If SummaryArray.TotalDMG <> 0 Then
80                FooterText = FooterText & vbTab & ResizePart(CStr(Round((.TotalDMG / SummaryArray.TotalDMG) * 100, 2)), 525)
90            Else
100               FooterText = FooterText & vbTab & ResizePart("0.00", 525)
110           End If
120       End If
130       If ReportOptions(0) = 1 Then 'MELEE DAMAGE
140           FooterText = FooterText & vbTab & ResizePart(CStr(.Basic.Damage + .Critical.Damage), 525)
150       End If
160       If ReportOptions(24) = 1 Then
170           FooterText = FooterText & vbTab & ResizePart(CStr(.Ranged.Damage), 525)
180       End If
190       If ReportOptions(2) = 1 Then
200           FooterText = FooterText & vbTab & ResizePart(CStr(.Spell.Damage), 525)
210       End If
220       If ReportOptions(1) = 1 Then
230           FooterText = FooterText & vbTab & ResizePart(CStr(.Skill.Damage), 525)
240       End If
250       If ReportOptions(37) = 1 Then
260           FooterText = FooterText & vbTab & ResizePart(CStr(.Ability.Damage), 525)
270       End If
280       If ReportOptions(21) = 1 Then
290           FooterText = FooterText & vbTab & ResizePart(CStr(.Effect.Damage), 525)
300       End If
    
310       If ReportOptions(7) = 1 Then 'MELEE HIT %
320           If .Basic.Hit + .Basic.Miss <> 0 Then
330               FooterText = FooterText & vbTab & ResizePart(CStr(Round((.TotalMeleeHit / (.TotalMeleeHit + .TotalMeleeMiss)) * 100, 2)), 525)
340           Else
350               FooterText = FooterText & vbTab & ResizePart("0.00", 525)
360           End If
370       End If
380       If ReportOptions(8) = 1 Then 'MELEE HIT/MISS
390           FooterText = FooterText & vbTab & ResizePart(CStr(.TotalMeleeHit) & "/" & CStr(.TotalMeleeMiss), 525)
400       End If
410       If ReportOptions(3) = 1 Then 'MELEE HIGH/LOW
420           If .Basic.Low = 10000 Then .Basic.Low = 0
430           dHigh = .Basic.High
440           If .Basic.Low <> 0 Then
450               dLow = .Basic.Low
460           Else
470               dLow = dHigh
480           End If
490           FooterText = FooterText & vbTab & ResizePart(CStr(dHigh) & "/" & CStr(dLow), 525)
500       End If
510       If ReportOptions(4) = 1 Then 'MELEE AVERAGE
520           If InStr(1, .Attacker, "SC:") Then
530               If .Skill.Damage <> 0 Then
540                   FooterText = FooterText & vbTab & ResizePart(CStr(Round((.Skill.Damage / .Skill.Hit), 2)), 525)
550               Else
560                   FooterText = FooterText & vbTab & ResizePart("00.00", 525)
570               End If
580           Else
590               If .Basic.Damage <> 0 Then
600                   FooterText = FooterText & vbTab & ResizePart(CStr(Round((.Basic.Damage / .Basic.Hit), 2)), 525)
610               Else
620                   FooterText = FooterText & vbTab & ResizePart("00.00", 525)
630               End If
640           End If
650       End If

660       If ReportOptions(26) = 1 Then 'RANGED HIT %
670           If .Ranged.Hit + .Ranged.Miss <> 0 Then
680               FooterText = FooterText & vbTab & ResizePart(CStr(Round((.TotalRangedHit / (.TotalRangedHit + .TotalRangedMiss)) * 100, 2)), 525)
690           Else
700               FooterText = FooterText & vbTab & ResizePart("0.00", 525)
710           End If
720       End If
730       If ReportOptions(25) = 1 Then 'RANGED HIT/MISS
740           FooterText = FooterText & vbTab & ResizePart(CStr(.TotalRangedHit) & "/" & CStr(.TotalRangedMiss), 525)
750       End If
760       If ReportOptions(27) = 1 Then 'RANGED HIGH/LOW
770           If .Ranged.Low = 10000 Then .Ranged.Low = 0
780           dHigh = .Ranged.High
790           If .Ranged.Low <> 0 Then
800               dLow = .Ranged.Low
810           Else
820               dLow = dHigh
830           End If
840           FooterText = FooterText & vbTab & ResizePart(CStr(dHigh) & "/" & CStr(dLow), 525)
850       End If
860       If ReportOptions(28) = 1 Then 'RANGED AVERAGE
870           If .Ranged.Damage <> 0 Then
880               FooterText = FooterText & vbTab & ResizePart(CStr(Round((.Ranged.Damage / .Ranged.Hit), 2)), 525)
890           Else
900               FooterText = FooterText & vbTab & ResizePart("00.00", 525)
910           End If
920       End If
930       If ReportOptions(30) = 1 Then 'SPELL HIGH/LOW
940           If .Spell.Low = 10000 Then .Spell.Low = 0
950           dHigh = .Spell.High
960           If .Spell.Low <> 0 Then
970               dLow = .Spell.Low
980           Else
990               dLow = dHigh
1000          End If
1010          FooterText = FooterText & vbTab & ResizePart(CStr(dHigh) & "/" & CStr(dLow), 525)
1020      End If
1030      If ReportOptions(29) = 1 Then 'SPELL AVERAGE
1040          If .Spell.Damage <> 0 Then
1050              FooterText = FooterText & vbTab & ResizePart(CStr(Round((.Spell.Damage / .Spell.Uses), 2)), 525)
1060          Else
1070              FooterText = FooterText & vbTab & ResizePart("00.00", 525)
1080          End If
1090      End If
1100      If ReportOptions(36) = 1 Then 'SPELL MP
1110          FooterText = FooterText & vbTab & ResizePart(CStr(.Spell.MPCost), 525)
1120      End If
1130      If ReportOptions(32) = 1 Then 'SKILL HIGH/LOW
1140          If .Skill.Low = 10000 Then .Skill.Low = 0
1150          dHigh = .Skill.High
1160          If .Skill.Low <> 0 Then
1170              dLow = .Skill.Low
1180          Else
1190              dLow = dHigh
1200          End If
1210          FooterText = FooterText & vbTab & ResizePart(CStr(dHigh) & "/" & CStr(dLow), 525)
1220      End If
1230      If ReportOptions(31) = 1 Then 'SKILL AVERAGE
1240          If .Skill.Damage <> 0 Then
1250              FooterText = FooterText & vbTab & ResizePart(CStr(Round((.Skill.Damage / .Skill.Hit), 2)), 525)
1260          Else
1270              FooterText = FooterText & vbTab & ResizePart("00.00", 525)
1280          End If
1290      End If
1300      If ReportOptions(22) = 1 Then 'WS COUNT
1310          FooterText = FooterText & vbTab & ResizePart(CStr(.Skill.Uses), 525)
1320      End If
1330      If ReportOptions(33) = 1 Then 'ABILITY HIGH/LOW
1340          If .Ability.Low = 10000 Then .Ability.Low = 0
1350          dHigh = .Ability.High
1360          If .Ability.Low <> 0 Then
1370              dLow = .Ability.Low
1380          Else
1390              dLow = dHigh
1400          End If
1410          FooterText = FooterText & vbTab & ResizePart(CStr(dHigh) & "/" & CStr(dLow), 525)
1420      End If
1430      If ReportOptions(34) = 1 Then 'ABILITY AVERAGE
1440          If .Ability.Damage <> 0 Then
1450              FooterText = FooterText & vbTab & ResizePart(CStr(Round((.Ability.Damage / .Ability.Hit), 2)), 525)
1460          Else
1470              FooterText = FooterText & vbTab & ResizePart("00.00", 525)
1480          End If
1490      End If
1500      If ReportOptions(5) = 1 Then 'CRIT PERCENT
1510          If .Critical.Hit <> 0 Then
1520              FooterText = FooterText & vbTab & ResizePart(CStr(Round((.Critical.Hit / (.TotalMeleeHit + .TotalRangedHit)) * 100, 2)), 525)
1530          Else
1540              FooterText = FooterText & vbTab & ResizePart("0.00", 525)
1550          End If
1560      End If
1570      If ReportOptions(6) = 1 Then 'CRIT COUNT
1580          FooterText = FooterText & vbTab & ResizePart(CStr(.Critical.Hit), 525)
1590      End If
1600      If ReportOptions(9) = 1 Then
1610          If .Evasion.TotalEvasion <> 0 Then
1620              FooterText = FooterText & vbTab & ResizePart(CStr(Round((.Evasion.TotalEvasion / (.Evasion.TotalEvasion + .Evasion.Hit)) * 100, 2)), 525)
1630          Else
1640              FooterText = FooterText & vbTab & ResizePart("0.00", 525)
1650          End If
1660      End If
1670      If ReportOptions(10) = 1 Then
1680          FooterText = FooterText & vbTab & ResizePart(CStr(.Evasion.Hit & "/" & .Evasion.TotalEvasion), 525)
1690      End If
1700      If ReportOptions(11) = 1 Then
1710          FooterText = FooterText & vbTab & ResizePart(CStr(.Evasion.Evade), 525)
1720      End If
1730      If ReportOptions(12) = 1 Then
1740          FooterText = FooterText & vbTab & ResizePart(CStr(.Evasion.Parry), 525)
1750      End If
1760      If ReportOptions(13) = 1 Then
1770          FooterText = FooterText & vbTab & ResizePart(CStr(.Evasion.Block), 525)
1780      End If
1790      If ReportOptions(14) = 1 Then
1800          FooterText = FooterText & vbTab & ResizePart(CStr(.Evasion.Absorb), 525)
1810      End If
1820      If ReportOptions(15) = 1 Then
1830          FooterText = FooterText & vbTab & ResizePart(CStr(.Evasion.Miss), 525)
1840      End If
1850      If ReportOptions(20) = 1 Then
1860          FooterText = FooterText & vbTab & ResizePart(CStr(.Evasion.Anticipate), 525)
1870      End If
1880      If ReportOptions(23) = 1 Then
1890          FooterText = FooterText & vbTab & ResizePart(CStr(.Counter.Hit), 525)
1900      End If
1910      If ReportOptions(16) = 1 Then
1920          FooterText = FooterText & vbTab & ResizePart(CStr(.Evasion.Damage), 525)
1930      End If
1940      If ReportOptions(17) = 1 Then
1950          FooterText = FooterText & vbTab & ResizePart(CStr(.Heal.Recovered), 525)
1960      End If
1970      If ReportOptions(19) = 1 Then
1980          FooterText = FooterText & vbTab & ResizePart(CStr(.Heal.Healed), 525)
1990      End If
2000      If ReportOptions(35) = 1 Then
2010          FooterText = FooterText & vbTab & ResizePart(CStr(.Heal.MPCost), 525)
2020      End If
2030      FooterText = FooterText & vbTab
2040  End With

2050  Exit Function
Err_Handler:
2060  HasErrors = True
2070  ErrorCount = ErrorCount + 1
2080  ReportError = "Error: " & Err.Number & vbNewLine & "Source: FooterText" & vbNewLine & "Description: " & Err.Description & vbNewLine & "Line: " & Erl & vbNewLine & "FFXI Log Line: " & CurrentLine & vbNewLine & "Previous FFXI Log Line: " & PrevLineA & vbNewLine & vbNewLine
2090  Print #ErrorFile, ReportError
2100  If ErrorCount >= 25 Then
2110      lblStatus.Caption = "Errors: " & HasErrors & " - Too many errors - Parsing stopped for this log."
2120      Exit Function
2130  Else
2140      Resume Next
2150  End If
End Function
Private Function GenerateCode(Source() As udtStatistics, SourceTotals As udtStatistics, Summary As Boolean) As String
10    On Error GoTo Err_Handler
      Dim HTMLCode As String, p As Long, i

20    If Summary Then
          Dim Job As String, SubJob As String, Level As String, Player As String, UserName As String
          Dim FoundPlayer As Boolean, PlayerCount As Integer
30        HTMLCode = HTMLCode & "<CENTER><TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0 style=""BORDER-COLLAPSE:collapse;font-family:verdana;font-size:7pt;color:black;""><TR>"
40        For p = 0 To UBound(Source) - 1
50            Player = Source(p).Attacker
60            If InStr(1, Player, ":") = 0 Then
70                HTMLCode = HTMLCode & "<TD><TABLE BORDER=1 CELLPADDING=0 CELLSPACING=0 style=""PADDING-LEFT: 3px;PADDING-RIGHT: 3px;BORDER-COLLAPSE:collapse;font-family:verdana;font-size:7pt;color:black;BORDER-RIGHT: #7CB1CB 1px solid; BORDER-TOP: #7CB1CB 1px solid; BORDER-LEFT: #7CB1CB 1px solid; BORDER-BOTTOM: #7CB1CB 1px solid""><TR>"
80                HTMLCode = HTMLCode & "<TD COLSPAN=2 BGColor=""7CB1CB""><b>" & Player & "</b></TD></TR>"
90                PlayerCount = GetSetting(App.Title, "Online_Setup", "Count", "5")
100               FoundPlayer = False
110               Job = ""
120               SubJob = ""
130               Level = ""
140               For i = 0 To PlayerCount
150                   UserName = GetSetting(App.Title, "Online_Setup", "Name" & i, "")
160                   If LCase(UserName) = LCase(Player) Then
170                       Job = GetSetting(App.Title, "Online_Setup", "Job" & i, "")
180                       SubJob = GetSetting(App.Title, "Online_Setup", "SubJob" & i, "")
190                       Level = GetSetting(App.Title, "Online_Setup", "Level" & i, "")
200                       FoundPlayer = True
210                       Exit For
220                   End If
230               Next
240               HTMLCode = HTMLCode & "<TR><TD BGCOLOR=""#b8ced9"">Level:</TD><TD>" & Level & "</TD></TR><TR><TD BGCOLOR=""#b8ced9"">Job:</TD><TD>" & Job & "</TD></TR><TR><TD BGCOLOR=""#b8ced9"">Sub:</TD><TD>" & SubJob & "</TD></TR></TABLE></TD>"
250           End If
260       Next
270       HTMLCode = HTMLCode & "</TR></TABLE></CENTER><p>"
280   End If

290   If Summary Then
300       HTMLCode = HTMLCode & "<CENTER><TABLE BORDER=1 CELLPADDING=0 CELLSPACING=0 style=""PADDING-LEFT: 3px;PADDING-RIGHT: 3px;BORDER-COLLAPSE:collapse;font-family:verdana;font-size:7pt;color:black;BORDER-RIGHT: #7CB1CB 1px solid; BORDER-TOP: #7CB1CB 1px solid; BORDER-LEFT: #7CB1CB 1px solid; BORDER-BOTTOM: #7CB1CB 1px solid"">" & vbNewLine
310       HTMLCode = HTMLCode & "<TR><TH colSpan=40 align=""Left"" BGColor=""7CB1CB"">Summary - " & SourceTotals.Battles & " battles.</font></TH></TR>" & vbNewLine
320       HTMLCode = HTMLCode & HTMLCodeHeader(Summary)
330   Else
340       HTMLCode = HTMLCode & "<CENTER><TABLE BORDER=1 CELLPADDING=0 CELLSPACING=0 style=""PADDING-LEFT: 3px;PADDING-RIGHT: 3px;BORDER-COLLAPSE:collapse;font-family:verdana;font-size:7pt;color:black;BORDER-RIGHT: #7CB1CB 1px solid; BORDER-TOP: #7CB1CB 1px solid; BORDER-LEFT: #7CB1CB 1px solid; BORDER-BOTTOM: #7CB1CB 1px solid"">" & vbNewLine
350       HTMLCode = HTMLCode & "<TR><TH colSpan=40 align=""Left"" BGColor=""7CB1CB"">" & SourceTotals.Defender & "</font></TH></TR>" & vbNewLine
360       HTMLCode = HTMLCode & HTMLCodeHeader(Summary)
370   End If
380   For p = 0 To UBound(Source) - 1
390       HTMLCode = HTMLCode & "<TR style=""BACKGROUND-COLOR:#dae6ec"">" & vbNewLine
400       HTMLCode = HTMLCode & "<TD BGCOLOR=""#b8ced9""><b>" & Source(p).Attacker & "</b></TD>" & vbNewLine 'PLAYER NAME
410       If ReportOptions(0) = 1 Then HTMLCode = HTMLCode & "<TD>" & Source(p).Basic.Damage + Source(p).Critical.Damage + Source(p).Counter.Damage & "</TD>" & vbNewLine              'BASIC DMG
420       If ReportOptions(24) = 1 Then HTMLCode = HTMLCode & "<TD>" & Source(p).Ranged.Damage & "</TD>" & vbNewLine                'RANGED DMG
430       If ReportOptions(2) = 1 Then HTMLCode = HTMLCode & "<TD>" & Source(p).Spell.Damage & "</TD>" & vbNewLine                'SPELL DMG
440       If ReportOptions(1) = 1 Then HTMLCode = HTMLCode & "<TD>" & Source(p).Skill.Damage & "</TD>" & vbNewLine                'SKILL DMG
450       If ReportOptions(37) = 1 Then HTMLCode = HTMLCode & "<TD>" & Source(p).Ability.Damage & "</TD>" & vbNewLine                'ABILITY DMG
460       If ReportOptions(21) = 1 Then HTMLCode = HTMLCode & "<TD>" & Source(p).Effect.Damage & "</TD>" & vbNewLine                'EFFECT
    
470       If ReportOptions(7) = 1 Then
480     If Source(p).TotalMeleeHit <> 0 Then 'MELEE HIT %
490         HTMLCode = HTMLCode & "<TD>" & Round((Source(p).TotalMeleeHit / (Source(p).TotalMeleeHit + Source(p).TotalMeleeMiss)) * 100, 2) & "%</TD>" & vbNewLine
500     Else
510         HTMLCode = HTMLCode & "<TD>0%</TD>" & vbNewLine
520     End If
530       End If
540       If ReportOptions(8) = 1 Then 'MELEE HIT/MISS
550     HTMLCode = HTMLCode & "<TD>" & Source(p).TotalMeleeHit & "/" & Source(p).TotalMeleeMiss & "</TD>" & vbNewLine
560       End If
570       If ReportOptions(3) = 1 Then 'MELEE HIGH/LOW
580     If Source(p).Basic.Low = 10000 Then Source(p).Basic.Low = 0
590     HTMLCode = HTMLCode & "<TD>" & Source(p).Basic.High & "/" & Source(p).Basic.Low & "</TD>" & vbNewLine                'High/Low
600       End If
610       If ReportOptions(4) = 1 Then 'MELEE AVERAGE
620     If Source(p).Basic.Damage And Source(p).TotalMeleeHit <> 0 Then
630         HTMLCode = HTMLCode & "<TD>" & Round(Source(p).Basic.Damage / Source(p).Basic.Hit, 2) & "</TD>" & vbNewLine
640     Else
650         HTMLCode = HTMLCode & "<TD>0</TD>" & vbNewLine
660     End If
670       End If

680       If ReportOptions(26) = 1 Then
690     If Source(p).TotalRangedHit <> 0 Then 'RANGED HIT %
700         HTMLCode = HTMLCode & "<TD>" & Round((Source(p).TotalRangedHit / (Source(p).TotalRangedHit + Source(p).TotalRangedMiss)) * 100, 2) & "%</TD>" & vbNewLine
710     Else
720         HTMLCode = HTMLCode & "<TD>0%</TD>" & vbNewLine
730     End If
740       End If
750       If ReportOptions(25) = 1 Then 'RANGED HIT/MISS
760     HTMLCode = HTMLCode & "<TD>" & Source(p).TotalRangedHit & "/" & Source(p).TotalRangedMiss & "</TD>" & vbNewLine
770       End If
780       If ReportOptions(27) = 1 Then 'RANGED HIGH/LOW
790     If Source(p).Ranged.Low = 10000 Then Source(p).Ranged.Low = 0
800     HTMLCode = HTMLCode & "<TD>" & Source(p).Ranged.High & "/" & Source(p).Ranged.Low & "</TD>" & vbNewLine                'High/Low
810       End If
820       If ReportOptions(28) = 1 Then 'RANGED AVERAGE
830     If Source(p).Ranged.Damage And Source(p).TotalRangedHit <> 0 Then
840         HTMLCode = HTMLCode & "<TD>" & Round(Source(p).Ranged.Damage / Source(p).TotalRangedHit, 2) & "</TD>" & vbNewLine
850     Else
860         HTMLCode = HTMLCode & "<TD>0</TD>" & vbNewLine
870     End If
880       End If

890       If ReportOptions(30) = 1 Then 'SPELL HIGH/LOW
900     If Source(p).Spell.Low = 10000 Then Source(p).Spell.Low = 0
910     HTMLCode = HTMLCode & "<TD>" & Source(p).Spell.High & "/" & Source(p).Spell.Low & "</TD>" & vbNewLine                'High/Low
920       End If
930       If ReportOptions(29) = 1 Then 'SPELL AVERAGE
940     If Source(p).Skill.Damage And Source(p).Spell.Uses <> 0 Then
950         HTMLCode = HTMLCode & "<TD>" & Round(Source(p).Spell.Damage / Source(p).Spell.Uses, 2) & "</TD>" & vbNewLine
960     Else
970         HTMLCode = HTMLCode & "<TD>0</TD>" & vbNewLine
980     End If
990       End If
1000      If ReportOptions(36) = 1 Then HTMLCode = HTMLCode & "<TD>" & Source(p).Spell.MPCost & "</TD>" & vbNewLine   'Spell MP
1010      If ReportOptions(32) = 1 Then 'SKILL HIGH/LOW
1020    If Source(p).Skill.Low = 10000 Then Source(p).Skill.Low = 0
1030    HTMLCode = HTMLCode & "<TD>" & Source(p).Skill.High & "/" & Source(p).Skill.Low & "</TD>" & vbNewLine                'High/Low
1040      End If
1050      If ReportOptions(31) = 1 Then 'SKILL AVERAGE
1060    If Source(p).Skill.Damage And Source(p).TotalMeleeHit <> 0 Then
1070        HTMLCode = HTMLCode & "<TD>" & Round(Source(p).Skill.Damage / Source(p).Skill.Hit, 2) & "</TD>" & vbNewLine
1080    Else
1090        HTMLCode = HTMLCode & "<TD>0</TD>" & vbNewLine
1100    End If
1110      End If
1120      If ReportOptions(22) = 1 Then HTMLCode = HTMLCode & "<TD>" & Source(p).Skill.Uses & "</TD>" & vbNewLine  'WS #
1130      If ReportOptions(33) = 1 Then 'ABILITY HIGH/LOW
1140    If Source(p).Ability.Low = 10000 Then Source(p).Ability.Low = 0
1150    HTMLCode = HTMLCode & "<TD>" & Source(p).Ability.High & "/" & Source(p).Ability.Low & "</TD>" & vbNewLine                'High/Low
1160      End If
1170      If ReportOptions(34) = 1 Then 'ABILITY AVERAGE
1180    If Source(p).Ability.Damage And Source(p).TotalMeleeHit <> 0 Then
1190        HTMLCode = HTMLCode & "<TD>" & Round(Source(p).Ability.Damage / Source(p).Ability.Hit, 2) & "</TD>" & vbNewLine
1200    Else
1210        HTMLCode = HTMLCode & "<TD>0</TD>" & vbNewLine
1220    End If
1230      End If
    
1240      If ReportOptions(5) = 1 Then 'CRIT %
1250    If Source(p).Critical.Hit <> 0 Then
1260        HTMLCode = HTMLCode & "<TD>" & Round((Source(p).Critical.Hit / (Source(p).TotalMeleeHit + Source(p).TotalRangedHit)) * 100, 2) & "%</TD>" & vbNewLine
1270    Else
1280        HTMLCode = HTMLCode & "<TD>0%</TD>" & vbNewLine
1290    End If
1300      End If 'CRIT COUNT
1310      If ReportOptions(6) = 1 Then HTMLCode = HTMLCode & "<TD>" & Source(p).Critical.Hit & "</TD>" & vbNewLine

    
1320      If ReportOptions(9) = 1 Then
1330    If Source(p).Evasion.TotalEvasion <> "0" Then 'Avoid %
1340        HTMLCode = HTMLCode & "<TD>" & Round((Source(p).Evasion.TotalEvasion / (Source(p).Evasion.Hit + Source(p).Evasion.TotalEvasion)) * 100, 2) & "%</TD>" & vbNewLine
1350    Else
1360        HTMLCode = HTMLCode & "<TD>0%</TD>" & vbNewLine
1370    End If
1380      End If
1390      If ReportOptions(10) = 1 Then HTMLCode = HTMLCode & "<TD>" & Source(p).Evasion.Hit & "/" & Source(p).Evasion.TotalEvasion & "</TD>" & vbNewLine                  'TAKE/Avoid
1400      If ReportOptions(11) = 1 Then HTMLCode = HTMLCode & "<TD>" & Source(p).Evasion.Evade & "</TD>" & vbNewLine                 'Evades
1410      If ReportOptions(12) = 1 Then HTMLCode = HTMLCode & "<TD>" & Source(p).Evasion.Parry & "</TD>" & vbNewLine                 'Parries
1420      If ReportOptions(13) = 1 Then HTMLCode = HTMLCode & "<TD>" & Source(p).Evasion.Block & "</TD>" & vbNewLine                 'Blocks
1430      If ReportOptions(14) = 1 Then HTMLCode = HTMLCode & "<TD>" & Source(p).Evasion.Absorb & "</TD>" & vbNewLine                 'Absorbs
1440      If ReportOptions(15) = 1 Then HTMLCode = HTMLCode & "<TD>" & Source(p).Evasion.Miss & "</TD>" & vbNewLine                 'Avoids
1450      If ReportOptions(20) = 1 Then HTMLCode = HTMLCode & "<TD>" & Source(p).Evasion.Anticipate & "</TD>" & vbNewLine                 'Anti
1460      If ReportOptions(23) = 1 Then HTMLCode = HTMLCode & "<TD>" & Source(p).Counter.Hit & "</TD>" & vbNewLine                 'Counters
1470      If ReportOptions(16) = 1 Then HTMLCode = HTMLCode & "<TD>" & Source(p).Evasion.Damage & "</TD>" & vbNewLine                 'DMG TAKEN
1480      If ReportOptions(17) = 1 Then HTMLCode = HTMLCode & "<TD>" & Source(p).Heal.Recovered & "</TD>" & vbNewLine                 'HP REC'D
1490      If ReportOptions(19) = 1 Then HTMLCode = HTMLCode & "<TD>" & Source(p).Heal.Healed & "</TD>" & vbNewLine                 'HP healed
1500      If ReportOptions(35) = 1 Then HTMLCode = HTMLCode & "<TD>" & Source(p).Heal.MPCost & "</TD>" & vbNewLine                 'HP MP
1510      If Summary Then
1520          For i = 0 To UBound(SummaryStats)
1530              If SummaryStats(i).Attacker = Source(p).Attacker Then
1540                  HTMLCode = HTMLCode & "<TD>" & SummaryStats(i).Battles & " / " & Round(((Source(p).TotalDMG / SummaryStats(i).Battles))) & "</TD>" & vbNewLine   'Fight Count and Average
1550                  Exit For
1560              End If
1570          Next
1580      End If
1590      HTMLCode = HTMLCode & "<TD BGCOLOR=""#b8ced9""><B>" & Source(p).TotalDMG & "</b> <FONT style=""font-family:small fonts;font-size:6pt"">(" & Round(((Source(p).TotalDMG / SourceTotals.TotalDMG) * 100), 2) & "%)</TD></TR>" & vbNewLine   'TOTAL AND % OF DMG
1600  Next
1610  HTMLCode = HTMLCode & "<TR style=""BACKGROUND-COLOR:#7CB1CB"">" & vbNewLine
1620  HTMLCode = HTMLCode & "<TD><B>Totals</B></TD>" & vbNewLine
1630  If ReportOptions(0) = 1 Then HTMLCode = HTMLCode & "<TD><B>" & Format(SourceTotals.Basic.Damage + SourceTotals.Critical.Damage, "#,###") & "</B> <FONT style=""font-family:small fonts;font-size:6pt""><br>(" & Round(((SourceTotals.Basic.Damage + SourceTotals.Critical.Damage) / SourceTotals.TotalDMG) * 100, 2) & "%)</TD>" & vbNewLine    'TOTAL BASIC

1640  If ReportOptions(24) = 1 Then ' RANGED DMG
1650      If SourceTotals.Ranged.Damage <> 0 Then
1660          HTMLCode = HTMLCode & "<TD><B>" & Format(SourceTotals.Ranged.Damage, "#,###") & "</B> <FONT style=""font-family:small fonts;font-size:6pt""><br>(" & Round((SourceTotals.Ranged.Damage / SourceTotals.TotalDMG) * 100, 2) & "%)</TD>" & vbNewLine        'TOTAL RANGED
1670      Else
1680          HTMLCode = HTMLCode & "<TD><B>0</B> <FONT style=""font-family:small fonts;font-size:6pt""><br>(0%)</TD>" & vbNewLine
1690      End If
1700  End If
1710  If ReportOptions(2) = 1 Then 'SPELL DMG
1720      If SourceTotals.Spell.Damage <> 0 Then
1730          HTMLCode = HTMLCode & "<TD><B>" & Format(SourceTotals.Spell.Damage, "#,###") & "</B> <FONT style=""font-family:small fonts;font-size:6pt""><br>(" & Round((SourceTotals.Spell.Damage / SourceTotals.TotalDMG) * 100, 2) & "%)</TD>" & vbNewLine        'TOTAL SPELL
1740      Else
1750          HTMLCode = HTMLCode & "<TD><B>0</B> <FONT style=""font-family:small fonts;font-size:6pt""><br>(0%)</TD>" & vbNewLine
1760      End If
1770  End If
1780  If ReportOptions(1) = 1 Then ' SKILL DMG
1790      If SourceTotals.Skill.Damage <> 0 Then
1800          HTMLCode = HTMLCode & "<TD><B>" & Format(SourceTotals.Skill.Damage, "#,###") & "</B> <FONT style=""font-family:small fonts;font-size:6pt""><br>(" & Round((SourceTotals.Skill.Damage / SourceTotals.TotalDMG) * 100, 2) & "%)</TD>" & vbNewLine        'TOTAL SKILL
1810      Else
1820          HTMLCode = HTMLCode & "<TD><B>0</B> <FONT style=""font-family:small fonts;font-size:6pt""><br>(0%)</TD>" & vbNewLine        'TOTAL SKILL
1830      End If
1840  End If
1850  If ReportOptions(37) = 1 Then 'ABILITY DMG
1860      If SourceTotals.Ability.Damage <> 0 Then
1870          HTMLCode = HTMLCode & "<TD><B>" & Format(SourceTotals.Ability.Damage, "#,###") & "</B> <FONT style=""font-family:small fonts;font-size:6pt""><br>(" & Round((SourceTotals.Ability.Damage / SourceTotals.TotalDMG) * 100, 2) & "%)</TD>" & vbNewLine        'TOTAL ABILITY
1880      Else
1890          HTMLCode = HTMLCode & "<TD><B>0</B> <FONT style=""font-family:small fonts;font-size:6pt""><br>(0%)</TD>" & vbNewLine        'TOTAL ABILITY
1900      End If
1910  End If
1920  If ReportOptions(21) = 1 Then 'EFFECT DMG
1930      If SourceTotals.Effect.Damage <> 0 Then
1940          HTMLCode = HTMLCode & "<TD><B>" & Format(SourceTotals.Effect.Damage, "#,###") & "</B> <FONT style=""font-family:small fonts;font-size:6pt""><br>(" & Round((SourceTotals.Effect.Damage / SourceTotals.TotalDMG) * 100, 2) & "%)</TD>" & vbNewLine        'TOTAL EFFECT
1950      Else
1960          HTMLCode = HTMLCode & "<TD><B>0</B> <FONT style=""font-family:small fonts;font-size:6pt""><br>(0%)</TD>" & vbNewLine
1970      End If
1980  End If
1990  If ReportOptions(7) = 1 Then HTMLCode = HTMLCode & "<TD></TD>" & vbNewLine 'MELEE HIT %
2000  If ReportOptions(8) = 1 Then HTMLCode = HTMLCode & "<TD></TD>" & vbNewLine 'MELEE HIT/MISS
2010  If ReportOptions(3) = 1 Then HTMLCode = HTMLCode & "<TD></TD>" & vbNewLine 'MELEE HIGH/LOW
2020  If ReportOptions(4) = 1 Then HTMLCode = HTMLCode & "<TD></TD>" & vbNewLine 'MELEE AVERAGE
2030  If ReportOptions(26) = 1 Then HTMLCode = HTMLCode & "<TD></TD>" & vbNewLine 'RANGED HIT %
2040  If ReportOptions(25) = 1 Then HTMLCode = HTMLCode & "<TD></TD>" & vbNewLine 'RANGED HIT/MISS
2050  If ReportOptions(27) = 1 Then HTMLCode = HTMLCode & "<TD></TD>" & vbNewLine  'RANGED HIGH/LOW
2060  If ReportOptions(28) = 1 Then HTMLCode = HTMLCode & "<TD></TD>" & vbNewLine 'RANGED AVERAGE
2070  If ReportOptions(30) = 1 Then HTMLCode = HTMLCode & "<TD></TD>" & vbNewLine 'SPELL HIGH/LOW
2080  If ReportOptions(29) = 1 Then HTMLCode = HTMLCode & "<TD></TD>" & vbNewLine 'SPELL AVERAGE
2090  If ReportOptions(36) = 1 Then HTMLCode = HTMLCode & "<TD></TD>" & vbNewLine 'SPELL MP USED
2100  If ReportOptions(32) = 1 Then HTMLCode = HTMLCode & "<TD></TD>" & vbNewLine  'SKILL HIGH/LOW
2110  If ReportOptions(31) = 1 Then HTMLCode = HTMLCode & "<TD></TD>" & vbNewLine  'SKILL AVERAGE
2120  If ReportOptions(33) = 1 Then HTMLCode = HTMLCode & "<TD></TD>" & vbNewLine  'ABILITY HIGH/LOW
2130  If ReportOptions(34) = 1 Then HTMLCode = HTMLCode & "<TD></TD>" & vbNewLine 'ABILITY AVERAGE
2140  If ReportOptions(5) = 1 Then HTMLCode = HTMLCode & "<TD></TD>" & vbNewLine 'CRIT %
2150  If ReportOptions(6) = 1 Then HTMLCode = HTMLCode & "<TD></TD>" & vbNewLine 'CRIT COUNT
2160  If ReportOptions(22) = 1 Then HTMLCode = HTMLCode & "<TD></TD>" & vbNewLine 'WS COUNT
2170  If ReportOptions(9) = 1 Then HTMLCode = HTMLCode & "<TD></TD>" & vbNewLine 'AVOID %
2180  If ReportOptions(10) = 1 Then HTMLCode = HTMLCode & "<TD></TD>" & vbNewLine 'TAKE AVOID
2190  If ReportOptions(11) = 1 Then HTMLCode = HTMLCode & "<TD></TD>" & vbNewLine 'EVADE
2200  If ReportOptions(12) = 1 Then HTMLCode = HTMLCode & "<TD></TD>" & vbNewLine 'PARRY
2210  If ReportOptions(13) = 1 Then HTMLCode = HTMLCode & "<TD></TD>" & vbNewLine 'BLOCK
2220  If ReportOptions(14) = 1 Then HTMLCode = HTMLCode & "<TD></TD>" & vbNewLine 'ABSORB
2230  If ReportOptions(15) = 1 Then HTMLCode = HTMLCode & "<TD></TD>" & vbNewLine 'AVOIDS
2240  If ReportOptions(20) = 1 Then HTMLCode = HTMLCode & "<TD></TD>" & vbNewLine 'ANTICIPATES
2250  If ReportOptions(23) = 1 Then HTMLCode = HTMLCode & "<TD></TD>" & vbNewLine 'COUNTERS
2260  If ReportOptions(16) = 1 Then HTMLCode = HTMLCode & "<TD><B>" & Format(SourceTotals.Evasion.Damage, "#,###") & "</B></TD>" & vbNewLine        'TOTAL TAKEN
2270  If ReportOptions(17) = 1 Then HTMLCode = HTMLCode & "<TD><B>" & Format(SourceTotals.Heal.Recovered, "#,###") & "</B></TD>" & vbNewLine        'TOTAL HP REC'D
2280  If ReportOptions(19) = 1 Then HTMLCode = HTMLCode & "<TD><B>" & Format(SourceTotals.Heal.Healed, "#,###") & "</B></TD>" & vbNewLine        'TOTAL given
2290  If ReportOptions(35) = 1 Then HTMLCode = HTMLCode & "<TD><B>" & Format(SourceTotals.Heal.MPCost, "#,###") & "</B></TD>" & vbNewLine        'TOTAL given
2300  If Summary Then HTMLCode = HTMLCode & "<TD>&nbsp;</TD>" & vbNewLine
2310  HTMLCode = HTMLCode & "<TD><B>" & Format(SourceTotals.TotalDMG, "#,###") & "</B></TD>" & vbNewLine 'TOTAL DMG DEALT
2320  HTMLCode = HTMLCode & "</TR>" & vbNewLine
2330  HTMLCode = HTMLCode & "</TABLE><P></CENTER>"
2340  GenerateCode = HTMLCode

2350  Exit Function
Err_Handler:
2360  HasErrors = True
2370  ErrorCount = ErrorCount + 1
2380  ReportError = "Error: " & Err.Number & vbNewLine & "Source: GenerateCode" & vbNewLine & "Description: " & Err.Description & vbNewLine & "Line: " & Erl & vbNewLine & "FFXI Log Line: " & CurrentLine & vbNewLine & "Previous FFXI Log Line: " & PrevLineA & vbNewLine & vbNewLine
2390  Print #ErrorFile, ReportError
2400  If ErrorCount >= 25 Then
2410      lblStatus.Caption = "Errors: " & HasErrors & " - Too many errors - Parsing stopped for this log."
2420      Exit Function
2430  Else
2440      Resume Next
2450  End If
End Function

Public Sub PlayerReport(Name As String)
10    On Error GoTo Err_Handler
      Dim PlayerStats() As udtStatistics, FoundBattle As Boolean, p As Integer, i As Integer
      Dim SummaryTotals As udtStatistics, EmptyStats As udtStatistics, LastLen As Long
20    ReDim PlayerStats(0)
30    RTB_User.Text = ""
40    RTB_User.SelBold = True
50    RTB_User.SelUnderline = True
60    RTB_User.SelText = ColumnText & vbNewLine
70    RTB_User.SelBold = False
80    RTB_User.SelUnderline = False

90    For i = 0 To UBound(FullStats) - 1
100       If listResults.Selected(FullStats(i).BattleID) = True And FullStats(i).Attacker = Name Then
110           FoundBattle = False
120           For p = 0 To UBound(PlayerStats)
130               If FullStats(i).Attacker = PlayerStats(p).Attacker Then
140                   CombineStats PlayerStats(p), FullStats(i)
150                   FoundBattle = True
160               End If
170           Next
180           If FoundBattle = False Then
190               CombineStats PlayerStats(UBound(PlayerStats)), FullStats(i)
200               ReDim Preserve PlayerStats(UBound(PlayerStats) + 1)
210           End If
220           RTB_User.SelBold = False
230           RTB_User.SelUnderline = False
240           RTB_User.SelText = PlayerText(PlayerStats, UBound(PlayerStats) - 1, "0") & vbNewLine
250           LastLen = Len(PlayerText(PlayerStats, UBound(PlayerStats) - 1, "0") & vbNewLine)
260       End If
270   Next

280   For i = 0 To UBound(PlayerStats) - 1
290       CombineStats SummaryTotals, PlayerStats(i)
300   Next
310   RTB_User.SelStart = Len(RTB_User.Text) - LastLen
320   RTB_User.SelLength = LastLen
330   RTB_User.SelUnderline = True
340   RTB_User.SelStart = Len(RTB_User.Text)
350   RTB_User.SelBold = True
360   RTB_User.SelUnderline = False
370   RTB_User.SelText = FooterText(SummaryTotals, False)

380   Exit Sub
Err_Handler:
390   HasErrors = True
400   ErrorCount = ErrorCount + 1
410   ReportError = "Error: " & Err.Number & vbNewLine & "Source: PlayerReport" & vbNewLine & "Description: " & Err.Description & vbNewLine & "Line: " & Erl & vbNewLine & "FFXI Log Line: " & CurrentLine & vbNewLine & "Previous FFXI Log Line: " & PrevLineA & vbNewLine & vbNewLine
420   Print #ErrorFile, ReportError
430   If ErrorCount >= 25 Then
440       lblStatus.Caption = "Errors: " & HasErrors & " - Too many errors - Parsing stopped for this log."
450       Exit Sub
460   Else
470       Resume Next
480   End If
End Sub

Private Function PlayerText(ArrayType() As udtStatistics, i As Integer, TotalDMG As Long) As String
10    On Error GoTo Err_Handler
      Dim dHigh As Integer, dLow As Integer
20    With ArrayType(i)
30        .Defender = .Defender
40        PlayerText = ResizePart(.Attacker, 1000)
50        PlayerText = PlayerText & vbTab & ResizePart(CStr(.TotalDMG), 525)
60        If ReportOptions(18) = 1 Then 'DAMAGE PERCENT
70            If TotalDMG <> 0 Then
80                PlayerText = PlayerText & vbTab & ResizePart(CStr(Round((.TotalDMG / TotalDMG) * 100, 2)), 525)
90            Else
100               PlayerText = PlayerText & vbTab & ResizePart("0.00", 525)
110           End If
120       End If
130       If .TotalDMG <> 0 Then 'Add text to AltHome feature
140         AltHome = AltHome & Trim(Left(.Attacker, 8)) & ":" & .TotalDMG & ","
150       End If
160       If ReportOptions(0) = 1 Then 'MELEE DAMAGE
170           PlayerText = PlayerText & vbTab & ResizePart(CStr(.Basic.Damage + .Critical.Damage), 525)
180       End If
190       If ReportOptions(24) = 1 Then
200           PlayerText = PlayerText & vbTab & ResizePart(CStr(.Ranged.Damage), 525)
210       End If
220       If ReportOptions(2) = 1 Then
230           PlayerText = PlayerText & vbTab & ResizePart(CStr(.Spell.Damage), 525)
240       End If
250       If ReportOptions(1) = 1 Then
260           PlayerText = PlayerText & vbTab & ResizePart(CStr(.Skill.Damage), 525)
270       End If
280       If ReportOptions(37) = 1 Then
290           PlayerText = PlayerText & vbTab & ResizePart(CStr(.Ability.Damage), 525)
300       End If
310       If ReportOptions(21) = 1 Then
320           PlayerText = PlayerText & vbTab & ResizePart(CStr(.Effect.Damage), 525)
330       End If
    
340       If ReportOptions(7) = 1 Then 'MELEE HIT %
350           If .Basic.Hit + .Basic.Miss <> 0 Then
360               PlayerText = PlayerText & vbTab & ResizePart(CStr(Round((.TotalMeleeHit / (.TotalMeleeHit + .TotalMeleeMiss)) * 100, 2)), 525)
370           Else
380               PlayerText = PlayerText & vbTab & ResizePart("0.00", 525)
390           End If
400       End If
410       If ReportOptions(8) = 1 Then 'MELEE HIT/MISS
420           PlayerText = PlayerText & vbTab & ResizePart(CStr(.TotalMeleeHit) & "/" & CStr(.TotalMeleeMiss), 525)
430       End If
440       If ReportOptions(3) = 1 Then 'MELEE HIGH/LOW
450           If .Basic.Low = 10000 Then .Basic.Low = 0
460           dHigh = .Basic.High
470           If .Basic.Low <> 0 Then
480               dLow = .Basic.Low
490           Else
500               dLow = dHigh
510           End If
520           PlayerText = PlayerText & vbTab & ResizePart(CStr(dHigh) & "/" & CStr(dLow), 525)
530       End If
540       If ReportOptions(4) = 1 Then 'MELEE AVERAGE
550           If InStr(1, .Attacker, "SC:") Then
560               If .Skill.Damage <> 0 Then
570                   PlayerText = PlayerText & vbTab & ResizePart(CStr(Round((.Skill.Damage / .Skill.Hit), 2)), 525)
580               Else
590                   PlayerText = PlayerText & vbTab & ResizePart("00.00", 525)
600               End If
610           Else
620               If .Basic.Damage <> 0 Then
630                   PlayerText = PlayerText & vbTab & ResizePart(CStr(Round((.Basic.Damage / .Basic.Hit), 2)), 525)
640               Else
650                   PlayerText = PlayerText & vbTab & ResizePart("00.00", 525)
660               End If
670           End If
680       End If

690       If ReportOptions(26) = 1 Then 'RANGED HIT %
700           If .Ranged.Hit + .Ranged.Miss <> 0 Then
710               PlayerText = PlayerText & vbTab & ResizePart(CStr(Round((.TotalRangedHit / (.TotalRangedHit + .TotalRangedMiss)) * 100, 2)), 525)
720           Else
730               PlayerText = PlayerText & vbTab & ResizePart("0.00", 525)
740           End If
750       End If
760       If ReportOptions(25) = 1 Then 'RANGED HIT/MISS
770           PlayerText = PlayerText & vbTab & ResizePart(CStr(.TotalRangedHit) & "/" & CStr(.TotalRangedMiss), 525)
780       End If
790       If ReportOptions(27) = 1 Then 'RANGED HIGH/LOW
800           If .Ranged.Low = 10000 Then .Ranged.Low = 0
810           dHigh = .Ranged.High
820           If .Ranged.Low <> 0 Then
830               dLow = .Ranged.Low
840           Else
850               dLow = dHigh
860           End If
870           PlayerText = PlayerText & vbTab & ResizePart(CStr(dHigh) & "/" & CStr(dLow), 525)
880       End If
890       If ReportOptions(28) = 1 Then 'RANGED AVERAGE
900           If .Ranged.Damage <> 0 Then
910               PlayerText = PlayerText & vbTab & ResizePart(CStr(Round((.Ranged.Damage / .Ranged.Hit), 2)), 525)
920           Else
930               PlayerText = PlayerText & vbTab & ResizePart("00.00", 525)
940           End If
950       End If
960       If ReportOptions(30) = 1 Then 'SPELL HIGH/LOW
970           If .Spell.Low = 10000 Then .Spell.Low = 0
980           dHigh = .Spell.High
990           If .Spell.Low <> 0 Then
1000              dLow = .Spell.Low
1010          Else
1020              dLow = dHigh
1030          End If
1040          PlayerText = PlayerText & vbTab & ResizePart(CStr(dHigh) & "/" & CStr(dLow), 525)
1050      End If
1060      If ReportOptions(29) = 1 Then 'SPELL AVERAGE
1070          If .Spell.Damage <> 0 Then
1080              PlayerText = PlayerText & vbTab & ResizePart(CStr(Round((.Spell.Damage / .Spell.Uses), 2)), 525)
1090          Else
1100              PlayerText = PlayerText & vbTab & ResizePart("00.00", 525)
1110          End If
1120      End If
1130      If ReportOptions(36) = 1 Then 'SPELL MP
1140          PlayerText = PlayerText & vbTab & ResizePart(CStr(.Spell.MPCost), 525)
1150      End If
1160      If ReportOptions(32) = 1 Then 'SKILL HIGH/LOW
1170          If .Skill.Low = 10000 Then .Skill.Low = 0
1180          dHigh = .Skill.High
1190          If .Skill.Low <> 0 Then
1200              dLow = .Skill.Low
1210          Else
1220              dLow = dHigh
1230          End If
1240          PlayerText = PlayerText & vbTab & ResizePart(CStr(dHigh) & "/" & CStr(dLow), 525)
1250      End If
1260      If ReportOptions(31) = 1 Then 'SKILL AVERAGE
1270          If .Skill.Damage <> 0 Then
1280              PlayerText = PlayerText & vbTab & ResizePart(CStr(Round((.Skill.Damage / .Skill.Hit), 2)), 525)
1290          Else
1300              PlayerText = PlayerText & vbTab & ResizePart("00.00", 525)
1310          End If
1320      End If
1330      If ReportOptions(22) = 1 Then 'WS COUNT
1340          PlayerText = PlayerText & vbTab & ResizePart(CStr(.Skill.Uses), 525)
1350      End If
1360      If ReportOptions(33) = 1 Then 'Ability HIGH/LOW
1370          If .Ability.Low = 10000 Then .Ability.Low = 0
1380          dHigh = .Ability.High
1390          If .Ability.Low <> 0 Then
1400              dLow = .Ability.Low
1410          Else
1420              dLow = dHigh
1430          End If
1440          PlayerText = PlayerText & vbTab & ResizePart(CStr(dHigh) & "/" & CStr(dLow), 525)
1450      End If
1460      If ReportOptions(34) = 1 Then 'Ability AVERAGE
1470          If .Ability.Damage <> 0 Then
1480              PlayerText = PlayerText & vbTab & ResizePart(CStr(Round((.Ability.Damage / .Ability.Hit), 2)), 525)
1490          Else
1500              PlayerText = PlayerText & vbTab & ResizePart("00.00", 525)
1510          End If
1520      End If
1530      If ReportOptions(5) = 1 Then 'CRIT PERCENT
1540          If .Critical.Hit <> 0 Then
1550              PlayerText = PlayerText & vbTab & ResizePart(CStr(Round((.Critical.Hit / (.TotalMeleeHit + .TotalRangedHit)) * 100, 2)), 525)
1560          Else
1570              PlayerText = PlayerText & vbTab & ResizePart("0.00", 525)
1580          End If
1590      End If
1600      If ReportOptions(6) = 1 Then 'CRIT COUNT
1610          PlayerText = PlayerText & vbTab & ResizePart(CStr(.Critical.Hit), 525)
1620      End If
1630      If ReportOptions(9) = 1 Then
1640          If .Evasion.TotalEvasion <> 0 Then
1650              PlayerText = PlayerText & vbTab & ResizePart(CStr(Round((.Evasion.TotalEvasion / (.Evasion.TotalEvasion + .Evasion.Hit)) * 100, 2)), 525)
1660          Else
1670              PlayerText = PlayerText & vbTab & ResizePart("0.00", 525)
1680          End If
1690      End If
1700      If ReportOptions(10) = 1 Then
1710          PlayerText = PlayerText & vbTab & ResizePart(CStr(.Evasion.Hit & "/" & .Evasion.TotalEvasion), 525)
1720      End If
1730      If ReportOptions(11) = 1 Then
1740          PlayerText = PlayerText & vbTab & ResizePart(CStr(.Evasion.Evade), 525)
1750      End If
1760      If ReportOptions(12) = 1 Then
1770          PlayerText = PlayerText & vbTab & ResizePart(CStr(.Evasion.Parry), 525)
1780      End If
1790      If ReportOptions(13) = 1 Then
1800          PlayerText = PlayerText & vbTab & ResizePart(CStr(.Evasion.Block), 525)
1810      End If
1820      If ReportOptions(14) = 1 Then
1830          PlayerText = PlayerText & vbTab & ResizePart(CStr(.Evasion.Absorb), 525)
1840      End If
1850      If ReportOptions(15) = 1 Then
1860          PlayerText = PlayerText & vbTab & ResizePart(CStr(.Evasion.Miss), 525)
1870      End If
1880      If ReportOptions(20) = 1 Then
1890          PlayerText = PlayerText & vbTab & ResizePart(CStr(.Evasion.Anticipate), 525)
1900      End If
1910      If ReportOptions(23) = 1 Then
1920          PlayerText = PlayerText & vbTab & ResizePart(CStr(.Counter.Hit), 525)
1930      End If
1940      If ReportOptions(16) = 1 Then
1950          PlayerText = PlayerText & vbTab & ResizePart(CStr(.Evasion.Damage), 525)
1960      End If
1970      If ReportOptions(17) = 1 Then
1980          PlayerText = PlayerText & vbTab & ResizePart(CStr(.Heal.Recovered), 525)
1990      End If
2000      If ReportOptions(19) = 1 Then
2010          PlayerText = PlayerText & vbTab & ResizePart(CStr(.Heal.Healed), 525)
2020      End If
2030      If ReportOptions(35) = 1 Then
2040          PlayerText = PlayerText & vbTab & ResizePart(CStr(.Heal.MPCost), 525)
2050      End If
2060      PlayerText = PlayerText & vbTab
2070  End With

2080  Exit Function
Err_Handler:
2090  HasErrors = True
2100  ErrorCount = ErrorCount + 1
2110  ReportError = "Error: " & Err.Number & vbNewLine & "Source: PlayerText" & vbNewLine & "Description: " & Err.Description & vbNewLine & "Line: " & Erl & vbNewLine & "FFXI Log Line: " & CurrentLine & vbNewLine & "Previous FFXI Log Line: " & PrevLineA & vbNewLine & vbNewLine
2120  Print #ErrorFile, ReportError
2130  If ErrorCount >= 25 Then
2140      lblStatus.Caption = "Errors: " & HasErrors & " - Too many errors - Parsing stopped for this log."
2150      Exit Function
2160  Else
2170      Resume Next
2180  End If
End Function
Private Sub CreateSingleFile()
Dim i As Integer
Dim FSO As FileSystemObject
Set FSO = New FileSystemObject

If InStr(1, SingleFile, Format(Date, "MM-DD-YYYY"), vbTextCompare) = 0 Then
    For i = 1 To 100
        If FSO.FileExists(App.Path & "\" & Format(Date, "MM-DD-YYYY") & AlphaNumber(i) & ".prs") = False Then
            SingleFile = App.Path & "\" & Format(Date, "MM-DD-YYYY") & AlphaNumber(i) & ".prs"
            Exit For
        End If
    Next
    If SingleFile = "" Then
        SingleFile = App.Path & "\" & Format(Date, "MM-DD-YYYY") & ".prs"
        FSO.DeleteFile SingleFile, True
    End If
End If
Set FSO = Nothing
End Sub

Private Function FindNumber() As Integer
10    On Error GoTo Err_Handler
      'This is used to retrieve the amount of damage/heal.
      Dim LineText As String, MyPos As Integer
20    MyPos = InStrRev(CurrentLine, ".")
30    If MyPos = 0 Then
40        MyPos = Len(CurrentLine) - 2
50    End If
60    LineText = Left(CurrentLine, MyPos)

      Dim i As Integer, FullNumber As String, FoundNumber As Boolean
70    For i = 1 To Len(LineText)
80        If IsNumeric(Mid(LineText, i, 1)) Then
90            FullNumber = FullNumber & Mid(LineText, i, 1)
100           FoundNumber = True
110       ElseIf FoundNumber Then
120           Exit For
130       End If
140   Next
150   If FullNumber <> "" Then
160       FindNumber = CDbl(FullNumber)
170   Else
180       FindNumber = 0
190   End If

200   Exit Function
Err_Handler:
210   HasErrors = True
220   ErrorCount = ErrorCount + 1
230   ReportError = "Error: " & Err.Number & vbNewLine & "Source: FindNumber" & vbNewLine & "Description: " & Err.Description & vbNewLine & "Line: " & Erl & vbNewLine & "FFXI Log Line: " & CurrentLine & vbNewLine & "Previous FFXI Log Line: " & PrevLineA & vbNewLine & vbNewLine
240   Print #ErrorFile, ReportError
250   If ErrorCount >= 25 Then
260       lblStatus.Caption = "Errors: " & HasErrors & " - Too many errors - Parsing stopped for this log."
270       Exit Function
280   Else
290       Resume Next
300   End If
End Function

Private Sub FishRPT()
10    On Error GoTo Err_Handler
      Dim lf, MyPos, AddLoot As String
20    If FishFound(0) <> "" Then
30        If FishHeader = "" Then FishHeader = "Unknown Time"
40        RTB_Fish.SelBold = True
50        RTB_Fish.SelText = FishHeader & vbNewLine
60        RTB_Fish.SelBold = False
70        For lf = 0 To UBound(FishFound)
80            RTB_Fish.SelBold = False
90            If FishFound(lf) <> "" Then
100               AddLoot = FishFound(lf)
110               MyPos = InStr(1, AddLoot, " - ")
120               AddLoot = Left$(AddLoot, MyPos - 1) & " - " & UCase(Mid$(AddLoot, MyPos + 3, 1)) & Mid$(AddLoot, MyPos + 4)
130               RTB_Fish.SelBold = False
140               RTB_Fish.SelColor = vbBlack
150               RTB_Fish.SelText = vbTab & AddLoot & vbNewLine
160           End If
170       Next
180       If FishComment <> "" Then
190           RTB_Fish.SelBold = False
200           RTB_Fish.SelColor = vbBlack
210           RTB_Fish.SelText = vbTab & "Comment: " & FishComment
220           FishComment = ""
230       End If
240       RTB_Fish.SelText = vbNewLine & vbNewLine
250   End If

260   Exit Sub
Err_Handler:
270   HasErrors = True
280   ErrorCount = ErrorCount + 1
290   ReportError = "Error: " & Err.Number & vbNewLine & "Source: FishRPT" & vbNewLine & "Description: " & Err.Description & vbNewLine & "Line: " & Erl & vbNewLine & "FFXI Log Line: " & CurrentLine & vbNewLine & "Previous FFXI Log Line: " & PrevLineA & vbNewLine & vbNewLine
300   Print #ErrorFile, ReportError
310   If ErrorCount >= 25 Then
320       lblStatus.Caption = "Errors: " & HasErrors & " - Too many errors - Parsing stopped for this log."
330       Exit Sub
340   Else
350       Resume Next
360   End If
End Sub



Private Sub GenerateReports(Recalculation As Boolean)
10    On Error GoTo Err_Handler
      Dim i As Integer, s As Integer, p As Integer, ThisText As String, LastTextLen As Long, FoundSummary As Boolean
20    AltHome = ""

30    With RTB_Report
40        .SelBold = True
50        .SelColor = &H40097
60        If FightComment <> "" Then
70            .SelText = BattleID & " - " & Defender & MonsterCheck & " - " & FightComment & vbNewLine
80        Else
90            .SelText = BattleID & " - " & Defender & MonsterCheck & vbNewLine
100       End If
110       MonsterCheck = ""
120       .SelBold = True
130       .SelColor = vbBlack
140       .SelUnderline = True
150       .SelText = ColumnText & vbNewLine
160   End With

      'Add statistics to BattleTotals and FullStats array
170   For i = 0 To UBound(BattleStats)
180       If BattleStats(i).Defender = Defender Then
190           CombineStats BattleTotals, BattleStats(i)
200           If Recalculation = False Then
210               CombineStats FullStats(UBound(FullStats)), BattleStats(i)
220               ReDim Preserve FullStats(UBound(FullStats) + 1)
230           End If
240       End If
250   Next
260   EffTotals.ATK = EffTotals.ATK + BattleTotals.TotalMeleeHit + BattleTotals.TotalMeleeMiss
270   EffTotals.ATKTaken = EffTotals.ATKTaken + BattleTotals.Evasion.Hit + BattleTotals.Evasion.TotalEvasion
280   EffTotals.BasicDMG = EffTotals.BasicDMG + BattleTotals.Basic.Damage + BattleTotals.Ranged.Damage + BattleTotals.Critical.Damage + BattleTotals.Counter.Damage
290   EffTotals.TotalDMG = EffTotals.TotalDMG + BattleTotals.TotalDMG
300   EffTotals.DMGTaken = EffTotals.DMGTaken + BattleTotals.Evasion.Damage

      'Add statistics to summary array
310   For i = 0 To UBound(BattleStats)
320       If BattleStats(i).Defender = Defender Then
330           FoundSummary = False
340           For s = 0 To UBound(SummaryStats)
350               If BattleStats(i).Attacker = SummaryStats(s).Attacker Then
360                   With SummaryStats(s)
370                       .BattleID = BattleStats(i).BattleID
380                       If BattleStats(i).TotalDMG <> 0 Then
390                           .Battles = .Battles + 1
400                       End If
410                       .Basic.Hit = .Basic.Hit + BattleStats(i).Basic.Hit
420                       .Basic.Miss = .Basic.Miss + BattleStats(i).Basic.Miss
430                       .Basic.Damage = .Basic.Damage + BattleStats(i).Basic.Damage
440                       .Ranged.Hit = .Ranged.Hit + BattleStats(i).Ranged.Hit
450                       .Ranged.Miss = .Ranged.Miss + BattleStats(i).Ranged.Miss
460                       .Ranged.Damage = .Ranged.Damage + BattleStats(i).Ranged.Damage
470                       .Counter.Hit = .Counter.Hit + BattleStats(i).Counter.Hit
480                       .Counter.Damage = .Counter.Damage + BattleStats(i).Counter.Damage
490                       .Skill.Hit = .Skill.Hit + BattleStats(i).Skill.Hit
500                       .Skill.Miss = .Skill.Miss + BattleStats(i).Skill.Miss
510                       .Skill.Uses = .Skill.Uses + BattleStats(i).Skill.Uses
520                       .Skill.Damage = .Skill.Damage + BattleStats(i).Skill.Damage
530                       .Ability.Hit = .Ability.Hit + BattleStats(i).Ability.Hit
540                       .Ability.Miss = .Ability.Miss + BattleStats(i).Ability.Miss
550                       .Ability.Uses = .Ability.Uses + BattleStats(i).Ability.Uses
560                       .Ability.Damage = .Ability.Damage + BattleStats(i).Ability.Damage
570                       .Critical.Hit = .Critical.Hit + BattleStats(i).Critical.Hit
580                       .Critical.Damage = .Critical.Damage + BattleStats(i).Critical.Damage
590                       .TotalMeleeHit = .TotalMeleeHit + BattleStats(i).TotalMeleeHit
600                       .TotalMeleeMiss = .TotalMeleeMiss + BattleStats(i).TotalMeleeMiss
610                       .TotalRangedHit = .TotalRangedHit + BattleStats(i).TotalRangedHit
620                       .TotalRangedMiss = .TotalRangedMiss + BattleStats(i).TotalRangedMiss
630                       .Spell.Damage = .Spell.Damage + BattleStats(i).Spell.Damage
640                       .Spell.Uses = .Spell.Uses + BattleStats(i).Spell.Uses
650                       .Effect.Damage = .Effect.Damage + BattleStats(i).Effect.Damage
660                       .TotalDMG = .TotalDMG + BattleStats(i).TotalDMG
670                       .Evasion.Damage = .Evasion.Damage + BattleStats(i).Evasion.Damage
680                       .Evasion.Parry = .Evasion.Parry + BattleStats(i).Evasion.Parry
690                       .Evasion.Block = .Evasion.Block + BattleStats(i).Evasion.Block
700                       .Evasion.Absorb = .Evasion.Absorb + BattleStats(i).Evasion.Absorb
710                       .Evasion.Anticipate = .Evasion.Anticipate + BattleStats(i).Evasion.Anticipate
720                       .Evasion.Evade = .Evasion.Evade + BattleStats(i).Evasion.Evade
730                       .Evasion.Miss = .Evasion.Miss + BattleStats(i).Evasion.Miss
740                       .Evasion.Hit = .Evasion.Hit + BattleStats(i).Evasion.Hit
750                       .Evasion.TotalEvasion = .Evasion.TotalEvasion + BattleStats(i).Evasion.TotalEvasion
760                       .Heal.Healed = .Heal.Healed + BattleStats(i).Heal.Healed
770                       .Heal.Recovered = .Heal.Recovered + BattleStats(i).Heal.Recovered
780                       .Heal.MPCost = .Heal.MPCost + BattleStats(i).Heal.MPCost
790                       .Spell.MPCost = .Spell.MPCost + BattleStats(i).Spell.MPCost
800                       If BattleStats(i).TotalDMG <> 0 Then
810                           .Percent = .Percent + Round((BattleStats(i).TotalDMG / BattleTotals.TotalDMG) * 100, 2)
820                       End If
830                       If (BattleStats(i).TotalMeleeHit + BattleStats(i).TotalRangedHit) <> 0 Then
840                           .Accuracy = .Accuracy + Round(((BattleStats(i).TotalMeleeHit + BattleStats(i).TotalRangedHit) / ((BattleStats(i).TotalMeleeHit + BattleStats(i).TotalRangedHit) + (BattleStats(i).TotalMeleeMiss + BattleStats(i).TotalRangedMiss))) * 100, 2)
850                       End If
860                   End With
870                   FoundSummary = True
880                   Exit For
890               End If
900           Next
910           If FoundSummary = False Then
920               ReDim Preserve SummaryStats(UBound(SummaryStats) + 1)
930               With SummaryStats(UBound(SummaryStats) - 1)
940                   .BattleID = BattleStats(i).BattleID
950                   .Attacker = BattleStats(i).Attacker
960                   .Defender = "Summary"
970                   If BattleStats(i).TotalDMG <> 0 Then
980                       .Battles = .Battles + 1
990                   End If
1000                  .Basic.Hit = .Basic.Hit + BattleStats(i).Basic.Hit
1010                  .Basic.Miss = .Basic.Miss + BattleStats(i).Basic.Miss
1020                  .Basic.Damage = .Basic.Damage + BattleStats(i).Basic.Damage
1030                  .Ranged.Hit = .Ranged.Hit + BattleStats(i).Ranged.Hit
1040                  .Ranged.Miss = .Ranged.Miss + BattleStats(i).Ranged.Miss
1050                  .Ranged.Damage = .Ranged.Damage + BattleStats(i).Ranged.Damage
1060                  .Counter.Hit = .Counter.Hit + BattleStats(i).Counter.Hit
1070                  .Counter.Damage = .Counter.Damage + BattleStats(i).Counter.Damage
1080                  .Skill.Hit = .Skill.Hit + BattleStats(i).Skill.Hit
1090                  .Skill.Miss = .Skill.Miss + BattleStats(i).Skill.Miss
1100                  .Skill.Uses = .Skill.Uses + BattleStats(i).Skill.Uses
1110                  .Skill.Damage = .Skill.Damage + BattleStats(i).Skill.Damage
1120                  .Ability.Hit = .Ability.Hit + BattleStats(i).Ability.Hit
1130                  .Ability.Miss = .Ability.Miss + BattleStats(i).Ability.Miss
1140                  .Ability.Uses = .Ability.Uses + BattleStats(i).Ability.Uses
1150                  .Ability.Damage = .Ability.Damage + BattleStats(i).Ability.Damage
1160                  .Critical.Hit = .Critical.Hit + BattleStats(i).Critical.Hit
1170                  .Critical.Damage = .Critical.Damage + BattleStats(i).Critical.Damage
1180                  .TotalMeleeHit = .TotalMeleeHit + BattleStats(i).TotalMeleeHit
1190                  .TotalMeleeMiss = .TotalMeleeMiss + BattleStats(i).TotalMeleeMiss
1200                  .TotalRangedHit = .TotalRangedHit + BattleStats(i).TotalRangedHit
1210                  .TotalRangedMiss = .TotalRangedMiss + BattleStats(i).TotalRangedMiss
1220                  .Spell.Damage = .Spell.Damage + BattleStats(i).Spell.Damage
1230                  .Spell.Uses = .Spell.Uses + BattleStats(i).Spell.Uses
1240                  .Effect.Damage = .Effect.Damage + BattleStats(i).Effect.Damage
1250                  .TotalDMG = .TotalDMG + BattleStats(i).TotalDMG
1260                  .Evasion.Damage = .Evasion.Damage + BattleStats(i).Evasion.Damage
1270                  .Evasion.Parry = .Evasion.Parry + BattleStats(i).Evasion.Parry
1280                  .Evasion.Block = .Evasion.Block + BattleStats(i).Evasion.Block
1290                  .Evasion.Absorb = .Evasion.Absorb + BattleStats(i).Evasion.Absorb
1300                  .Evasion.Anticipate = .Evasion.Anticipate + BattleStats(i).Evasion.Anticipate
1310                  .Evasion.Evade = .Evasion.Evade + BattleStats(i).Evasion.Evade
1320                  .Evasion.Miss = .Evasion.Miss + BattleStats(i).Evasion.Miss
1330                  .Evasion.Hit = .Evasion.Hit + BattleStats(i).Evasion.Hit
1340                  .Evasion.TotalEvasion = .Evasion.TotalEvasion + BattleStats(i).Evasion.TotalEvasion
1350                  .Heal.Healed = .Heal.Healed + BattleStats(i).Heal.Healed
1360                  .Heal.Recovered = .Heal.Recovered + BattleStats(i).Heal.Recovered
1370                  .Heal.MPCost = .Heal.MPCost + BattleStats(i).Heal.MPCost
1380                  .Spell.MPCost = .Spell.MPCost + BattleStats(i).Spell.MPCost
1390                  If BattleStats(i).TotalDMG <> 0 Then
1400                      .Percent = .Percent + Round((BattleStats(i).TotalDMG / BattleTotals.TotalDMG) * 100, 2)
1410                  End If
1420                  If (BattleStats(i).TotalMeleeHit + BattleStats(i).TotalRangedHit) <> 0 Then
1430                      .Accuracy = .Accuracy + Round(((BattleStats(i).TotalMeleeHit + BattleStats(i).TotalRangedHit) / ((BattleStats(i).TotalMeleeHit + BattleStats(i).TotalRangedHit) + (BattleStats(i).TotalMeleeMiss + BattleStats(i).TotalRangedMiss))) * 100, 2)
1440                  End If
1450              End With
1460          End If
1470      End If
1480  Next

      Dim HighDMG As Long, LowDMG As Long
1490  LowDMG = 99999
1500  For i = 0 To UBound(BattleStats)
1510      With BattleStats(i)
1520          If .Defender = Defender And Left(.Attacker, 2) <> "SC" And .TotalDMG > 10 Then
1530              If .TotalDMG > HighDMG Then
1540                  HighDMG = .TotalDMG
1550              End If
1560              If .TotalDMG < LowDMG Then
1570                  LowDMG = .TotalDMG
1580              End If
1590          End If
1600      End With
1610  Next

1620  For i = 0 To UBound(BattleStats)
1630      With BattleStats(i)
1640          If .Defender = Defender Then
1650              If Recalculation = False Then
1660                  RTB_Details.SelBold = True
1670                  RTB_Details.SelColor = vbBlue
1680                  RTB_Details.SelText = .Attacker & " - " & .Defender & "(" & BattleID & ")" & vbNewLine
1690                  RTB_Details.SelColor = vbBlack
1700                  RTB_Details.SelBold = False
    
1710                  If .Basic.List <> "" Then
1720                      RTB_Details.SelBold = True
1730                      RTB_Details.SelText = vbTab & "Melee Damage: "
1740                      RTB_Details.SelBold = False
1750                      RTB_Details.SelText = .Basic.List & vbNewLine
1760                  End If
1770                  If .Critical.List <> "" Then
1780                      RTB_Details.SelBold = True
1790                      RTB_Details.SelText = vbTab & "Critical Damage: "
1800                      RTB_Details.SelBold = False
1810                      RTB_Details.SelText = .Critical.List & vbNewLine
1820                  End If
1830                  If .Ranged.List <> "" Then
1840                      RTB_Details.SelBold = True
1850                      RTB_Details.SelText = vbTab & "Ranged Damage: "
1860                      RTB_Details.SelBold = False
1870                      RTB_Details.SelText = .Ranged.List & vbNewLine
1880                  End If
1890                  If .Skill.List <> "" Then
1900                      RTB_Details.SelBold = True
1910                      RTB_Details.SelText = vbTab & "WeaponSkills: "
1920                      RTB_Details.SelBold = False
1930                      RTB_Details.SelText = .Skill.List & vbNewLine
1940                  End If
1950                  If .Ability.List <> "" Then
1960                      RTB_Details.SelBold = True
1970                      RTB_Details.SelText = vbTab & "Abilities: "
1980                      RTB_Details.SelBold = False
1990                      RTB_Details.SelText = .Ability.List & vbNewLine
2000                  End If
2010                  If .Effect.List <> "" Then
2020                      RTB_Details.SelBold = True
2030                      RTB_Details.SelText = vbTab & "Additional Effects: "
2040                      RTB_Details.SelBold = False
2050                      RTB_Details.SelText = .Effect.List & vbNewLine
2060                  End If
2070                  If .Spell.List <> "" Then
2080                      RTB_Details.SelBold = True
2090                      RTB_Details.SelText = vbTab & "Spells: "
2100                      RTB_Details.SelBold = False
2110                      RTB_Details.SelText = .Spell.List & vbNewLine
2120                  End If
2130                  If .Heal.HealedList <> "" Then
2140                      RTB_Details.SelBold = True
2150                      RTB_Details.SelText = vbTab & "Heals: "
2160                      RTB_Details.SelBold = False
2170                      RTB_Details.SelText = .Heal.HealedList & vbNewLine
2180                  End If
2190                  If .Heal.RecoveredList <> "" Then
2200                      RTB_Details.SelBold = True
2210                      RTB_Details.SelText = vbTab & "Recovered: "
2220                      RTB_Details.SelBold = False
2230                      RTB_Details.SelText = .Heal.RecoveredList & vbNewLine
2240                  End If
2250                  RTB_Details.SelBold = True
2260                  RTB_Details.SelText = vbTab & "Total Damage: "
2270                  RTB_Details.SelBold = False
2280                  RTB_Details.SelColor = vbRed
2290                  RTB_Details.SelText = .TotalDMG & vbNewLine
2300                  RTB_Details.SelColor = vbBlack
2310                  RTB_Details.SelText = vbNewLine
2320              End If
  
  
2330              RTB_Report.SelUnderline = False
2340              ThisText = PlayerText(BattleStats, i, BattleTotals.TotalDMG)

2350              If .TotalDMG = HighDMG Then
2360                  RTB_Report.SelColor = vbBlue
2370              ElseIf .TotalDMG = LowDMG Then
2380                  RTB_Report.SelColor = vbRed
2390              Else
2400                  RTB_Report.SelColor = vbBlack
2410              End If
2420              RTB_Report.SelBold = False
2430              RTB_Report.SelText = ThisText & vbNewLine
2440              LastTextLen = Len(ThisText) + 2
                  'ClearBattleStats i, False
2450          End If
2460      End With
2470  Next
2480  RTB_Report.SelStart = Len(RTB_Report.Text) - LastTextLen
2490  RTB_Report.SelLength = LastTextLen
2500  RTB_Report.SelUnderline = True

2510  RTB_Report.SelStart = Len(RTB_Report.Text)
2520  RTB_Report.SelColor = vbBlack
2530  RTB_Report.SelBold = True
2540  RTB_Report.SelText = FooterText(BattleTotals, True) & vbNewLine & vbNewLine

      Dim EstDPS As String, dp As Integer
2550  If optionSummary(0).Value = True Then
2560      With RTB_Averages
2570          .Text = ""
2580          .SelBold = True
2590          .SelText = "Experience" & vbNewLine
2600          .SelBold = False
  
2610          If TotalExp <> 0 And StartTime <> Empty And StopTime <> Empty Then
2620            .SelText = "Start: " & StartTime & " / Stop: " & StopTime & vbNewLine & "Total Exp: " & TotalExp & vbNewLine & "Per Hour: " & Round(TotalExp / DateDiff("n", StartTime, StopTime), 2) * 60 & vbNewLine & "Per Minute: " & Round(TotalExp / DateDiff("n", StartTime, StopTime), 2) & vbNewLine & vbNewLine
2630          ElseIf TotalExp <> 0 And StartTime <> Empty Then
2640            .SelText = "Start: " & StartTime & " / Stop: " & Now & vbNewLine & "Total Exp: " & TotalExp & vbNewLine & "Per Hour: " & Round(TotalExp / DateDiff("n", StartTime, Now), 2) * 60 & vbNewLine & "Per Minute: " & Round(TotalExp / DateDiff("n", StartTime, Now), 2) & vbNewLine & vbNewLine
2650          Else
2660            .SelText = "Start: " & StartTime & vbNewLine & "Total Exp: " & TotalExp & vbNewLine & "Per Hour: 0" & vbNewLine & "Per Minute: 0" & vbNewLine & vbNewLine
2670          End If
  
2680          .SelBold = True
2690          .SelText = "Experience Chains" & vbNewLine
2700          .SelBold = False
2710          For p = 0 To UBound(ChainExp)
2720              If ChainExp(p, 0) <> 0 Then
2730                  .SelText = "EXP Chain #" & p & ": " & CStr(ChainExp(p, 0)) & " - " & CStr(ChainExp(p, 1)) & " times, " & Round(ChainExp(p, 0) / ChainExp(p, 1), 2) & " average." & vbNewLine
2740              End If
2750          Next
2760          .SelText = vbNewLine
2770      End With
2780      For i = 0 To UBound(SummaryStats) - 1
2790          With SummaryStats(i)
2800              RTB_Averages.SelBold = True
2810              RTB_Averages.SelText = .Attacker & vbNewLine
2820              RTB_Averages.SelBold = False
2830              RTB_Averages.SelText = ResizePart("Total Fights: ", 1500) & vbTab & .Battles & vbNewLine
2840              If .Battles <> 0 Then
2850                  RTB_Averages.SelText = ResizePart("Average Damage: ", 1500) & vbTab & Round(.TotalDMG / .Battles, 2) & vbNewLine
2860                  RTB_Averages.SelText = ResizePart("Average Percent: ", 1500) & vbTab & Round((.Percent / .Battles), 2) & vbNewLine
2870                  RTB_Averages.SelText = ResizePart("Average Accuracy: ", 1500) & vbTab & Round((.Accuracy / .Battles), 2) & vbNewLine
2880              Else
2890                  RTB_Averages.SelText = ResizePart("Average Damage: ", 1500) & vbTab & "0.00" & vbNewLine
2900                  RTB_Averages.SelText = ResizePart("Average Percent: ", 1500) & vbTab & "0.00" & vbNewLine
2910                  RTB_Averages.SelText = ResizePart("Average Accuracy: ", 1500) & vbTab & "0.00" & vbNewLine
2920              End If

2930              EstDPS = ""
2940              For dp = 0 To UBound(DPS)
2950                If Trim(DPS(dp, 0)) = .Attacker Then
2960                    If DPS(dp, 0) <> "" Then
2970                          If DPS(dp, 1) <> "0" And DPS(dp, 2) <> "0" And DPS(dp, 2) <> "" And DPS(dp, 1) <> "" Then
2980                              EstDPS = Round(CDbl(DPS(dp, 1)) / CDbl(DPS(dp, 2)), 2) & " (" & DPS(dp, 2) & " seconds / " & DPS(dp, 1) & " dmg)"
2990                          Else
3000                              EstDPS = "0.00"
3010                          End If
3020                        Exit For
3030                    End If
3040                End If
3050              Next
3060              RTB_Averages.SelText = ResizePart("Estimated DPS: ", 1500) & vbTab & EstDPS & vbNewLine & vbNewLine
3070          End With
3080      Next
3090      RTB_Averages.SelStart = 0
3100  End If
3110  If AltHome <> "" Then
3120    AltHome = "TTL:" & BattleTotals.TotalDMG & "," & Left(AltHome, Len(AltHome) - 1)
3130    If mnuEnableSounds.Checked = True Then
3140      If OpenSingle = False And (Gather = False Or ParseGather = True) Then
3150          If BeepNotWave Then
3160              Call Beep(100, 100)
3170          Else
3180              If BeepSounds(10) <> "" Then
3190                  PlaySound BeepSounds(10), 0&, SND_FILENAME Or SND_ASYNC
3200              Else
3210                  Call Beep(100, 100)
3220              End If
3230          End If
3240      End If
3250    End If
3260  End If

3270  Exit Sub
Err_Handler:
3280  HasErrors = True
3290  ErrorCount = ErrorCount + 1
3300  ReportError = "Error: " & Err.Number & vbNewLine & "Source: GenerateReports" & vbNewLine & "Description: " & Err.Description & vbNewLine & "Line: " & Erl & vbNewLine & "FFXI Log Line: " & CurrentLine & vbNewLine & "Previous FFXI Log Line: " & PrevLineA & vbNewLine & vbNewLine
3310  Print #ErrorFile, ReportError
3320  If ErrorCount >= 25 Then
3330      lblStatus.Caption = "Errors: " & HasErrors & " - Too many errors - Parsing stopped for this log."
3340      Exit Sub
3350  Else
3360      Resume Next
3370  End If
End Sub

Private Sub GetMonsterCheck()
10    On Error GoTo Err_Handler
20    If InStr(1, LCase(CurrentLine), "decent") Then
30        MonsterCheck = "(DC)"
40    ElseIf InStr(1, LCase(CurrentLine), "impossible to guage") Then
50        MonsterCheck = "(IG)"
    ElseIf InStr(1, LCase(CurrentLine), "very tough") Then
        MonsterCheck = "(VT)"
60    ElseIf InStr(1, LCase(CurrentLine), "incredibly tough") Then
70        MonsterCheck = "(IT)"
80    ElseIf InStr(1, LCase(CurrentLine), "tough") Then
90        MonsterCheck = "(T)"
100   ElseIf InStr(1, LCase(CurrentLine), "weak") Then
110       MonsterCheck = "(TW)"
120   ElseIf InStr(1, LCase(CurrentLine), "easy") Then
130       MonsterCheck = "(EP)"
140   ElseIf InStr(1, LCase(CurrentLine), "even") Then
150       MonsterCheck = "(EM)"
160   End If

170   Exit Sub
Err_Handler:
180   HasErrors = True
190   ErrorCount = ErrorCount + 1
200   ReportError = "Error: " & Err.Number & vbNewLine & "Source: GetMonsterCheck" & vbNewLine & "Description: " & Err.Description & vbNewLine & "Line: " & Erl & vbNewLine & "FFXI Log Line: " & CurrentLine & vbNewLine & "Previous FFXI Log Line: " & PrevLineA & vbNewLine & vbNewLine
210   Print #ErrorFile, ReportError
220   If ErrorCount >= 25 Then
230       lblStatus.Caption = "Errors: " & HasErrors & " - Too many errors - Parsing stopped for this log."
240       Exit Sub
250   Else
260       Resume Next
270   End If
End Sub

Private Function HTMLCodeHeader(Summary As Boolean) As String
10    On Error GoTo Err_Handler
20    HTMLCodeHeader = ""
30    HTMLCodeHeader = HTMLCodeHeader & "<TR style=""FONT-WEIGHT:bold;BACKGROUND-COLOR:#b8ced9"">"
40    HTMLCodeHeader = HTMLCodeHeader & "<TD WIDTH=75></TD>"
50    If ReportOptions(0) = 1 Then HTMLCodeHeader = HTMLCodeHeader & "<TD WIDTH=50>Melee</TD>"
60    If ReportOptions(24) = 1 Then HTMLCodeHeader = HTMLCodeHeader & "<TD WIDTH=50>Ranged</TD>"
70    If ReportOptions(2) = 1 Then HTMLCodeHeader = HTMLCodeHeader & "<TD WIDTH=50>Spell</TD>"
80    If ReportOptions(1) = 1 Then HTMLCodeHeader = HTMLCodeHeader & "<TD WIDTH=50>Skill</TD>"
90    If ReportOptions(37) = 1 Then HTMLCodeHeader = HTMLCodeHeader & "<TD WIDTH=50>Ability</TD>"
100   If ReportOptions(21) = 1 Then HTMLCodeHeader = HTMLCodeHeader & "<TD WIDTH=50>Effects</TD>"


110   If ReportOptions(7) = 1 Then HTMLCodeHeader = HTMLCodeHeader & "<TD WIDTH=50>M Hit%</TD>" & vbNewLine 'MELEE HIT %
120   If ReportOptions(8) = 1 Then HTMLCodeHeader = HTMLCodeHeader & "<TD WIDTH=50>M Ht/Ms</TD>" & vbNewLine 'MELEE HIT/MISS
130   If ReportOptions(3) = 1 Then HTMLCodeHeader = HTMLCodeHeader & "<TD WIDTH=50>M Hi/Lo</TD>" & vbNewLine 'MELEE HIGH/LOW
140   If ReportOptions(4) = 1 Then HTMLCodeHeader = HTMLCodeHeader & "<TD WIDTH=50>M Avg</TD>" & vbNewLine 'MELEE AVERAGE
150   If ReportOptions(26) = 1 Then HTMLCodeHeader = HTMLCodeHeader & "<TD WIDTH=50>R Hit%</TD>" & vbNewLine 'RANGED HIT %
160   If ReportOptions(25) = 1 Then HTMLCodeHeader = HTMLCodeHeader & "<TD WIDTH=50>R Ht/Ms</TD>" & vbNewLine 'RANGED HIT/MISS
170   If ReportOptions(27) = 1 Then HTMLCodeHeader = HTMLCodeHeader & "<TD WIDTH=50>R Hi/Lo</TD>" & vbNewLine  'RANGED HIGH/LOW
180   If ReportOptions(28) = 1 Then HTMLCodeHeader = HTMLCodeHeader & "<TD WIDTH=50>R Avg</TD>" & vbNewLine 'RANGED AVERAGE
190   If ReportOptions(30) = 1 Then HTMLCodeHeader = HTMLCodeHeader & "<TD WIDTH=50>Sp Hi/Lo</TD>" & vbNewLine 'SPELL HIGH/LOW
200   If ReportOptions(29) = 1 Then HTMLCodeHeader = HTMLCodeHeader & "<TD WIDTH=50>Sp Avg</TD>" & vbNewLine 'SPELL AVERAGE
210   If ReportOptions(36) = 1 Then HTMLCodeHeader = HTMLCodeHeader & "<TD WIDTH=50>Sp MP</TD>" & vbNewLine 'SPELL MP USED
220   If ReportOptions(32) = 1 Then HTMLCodeHeader = HTMLCodeHeader & "<TD WIDTH=50>Sk Hi/Lo</TD>" & vbNewLine  'SKILL HIGH/LOW
230   If ReportOptions(31) = 1 Then HTMLCodeHeader = HTMLCodeHeader & "<TD WIDTH=50>Sk Avg</TD>" & vbNewLine  'SKILL AVERAGE
240   If ReportOptions(22) = 1 Then HTMLCodeHeader = HTMLCodeHeader & "<TD WIDTH=50>Sk #</TD>" & vbNewLine 'WS COUNT
250   If ReportOptions(33) = 1 Then HTMLCodeHeader = HTMLCodeHeader & "<TD WIDTH=50>Ab Hi/Lo</TD>" & vbNewLine  'ABILITY HIGH/LOW
260   If ReportOptions(34) = 1 Then HTMLCodeHeader = HTMLCodeHeader & "<TD WIDTH=50>Ab Avg</TD>" & vbNewLine 'ABILITY AVERAGE
270   If ReportOptions(5) = 1 Then HTMLCodeHeader = HTMLCodeHeader & "<TD WIDTH=50>Crit %</TD>" & vbNewLine 'CRIT %
280   If ReportOptions(6) = 1 Then HTMLCodeHeader = HTMLCodeHeader & "<TD WIDTH=50>Crit #</TD>" & vbNewLine 'CRIT COUNT
290   If ReportOptions(9) = 1 Then HTMLCodeHeader = HTMLCodeHeader & "<TD WIDTH=50>Avoid %</TD>" & vbNewLine 'AVOID %
300   If ReportOptions(10) = 1 Then HTMLCodeHeader = HTMLCodeHeader & "<TD WIDTH=50>Take/Avoid</TD>" & vbNewLine 'TAKE AVOID
310   If ReportOptions(11) = 1 Then HTMLCodeHeader = HTMLCodeHeader & "<TD WIDTH=50>Evades</TD>" & vbNewLine 'EVADE
320   If ReportOptions(12) = 1 Then HTMLCodeHeader = HTMLCodeHeader & "<TD WIDTH=50>Parries</TD>" & vbNewLine 'PARRY
330   If ReportOptions(13) = 1 Then HTMLCodeHeader = HTMLCodeHeader & "<TD WIDTH=50>Blocks</TD>" & vbNewLine 'BLOCK
340   If ReportOptions(14) = 1 Then HTMLCodeHeader = HTMLCodeHeader & "<TD WIDTH=50>Absorbs</TD>" & vbNewLine 'ABSORB
350   If ReportOptions(15) = 1 Then HTMLCodeHeader = HTMLCodeHeader & "<TD WIDTH=50>Avoids</TD>" & vbNewLine 'AVOIDS
360   If ReportOptions(20) = 1 Then HTMLCodeHeader = HTMLCodeHeader & "<TD WIDTH=50>Anticipates</TD>" & vbNewLine 'ANTICIPATES
370   If ReportOptions(23) = 1 Then HTMLCodeHeader = HTMLCodeHeader & "<TD WIDTH=50>Counters</TD>" & vbNewLine 'COUNTERS
380   If ReportOptions(16) = 1 Then HTMLCodeHeader = HTMLCodeHeader & "<TD WIDTH=75>DMG Taken</TD>"
390   If ReportOptions(17) = 1 Then HTMLCodeHeader = HTMLCodeHeader & "<TD WIDTH=75>HP Rec'd</TD>"
400   If ReportOptions(19) = 1 Then HTMLCodeHeader = HTMLCodeHeader & "<TD WIDTH=75>HP Healed</TD>"
410   If ReportOptions(35) = 1 Then HTMLCodeHeader = HTMLCodeHeader & "<TD WIDTH=75>HP MP</TD>"
420   If Summary Then HTMLCodeHeader = HTMLCodeHeader & "<TD WIDTH=75>Fights / Avg</TD>" & vbNewLine
430   HTMLCodeHeader = HTMLCodeHeader & "<TD WIDTH=75>TTL DMG</TD></TR>" & vbNewLine
440   Exit Function

Err_Handler:
450   HasErrors = True
460   ErrorCount = ErrorCount + 1
470   ReportError = "Error: " & Err.Number & vbNewLine & "Source: HTMLCodeHeader" & vbNewLine & "Description: " & Err.Description & vbNewLine & "Line: " & Erl & vbNewLine & "FFXI Log Line: " & CurrentLine & vbNewLine & "Previous FFXI Log Line: " & PrevLineA & vbNewLine & vbNewLine
480   Print #ErrorFile, ReportError
490   If ErrorCount >= 25 Then
500       lblStatus.Caption = "Errors: " & HasErrors & " - Too many errors - Parsing stopped for this log."
510       Exit Function
520   Else
530       Resume Next
540   End If
End Function



Public Sub ParseLog(FullFile() As String)
10    On Error GoTo Err_Handler
      Dim ff As Long, i As Integer
      Dim FoundInstance As Boolean

20    For ff = 0 To UBound(FullFile)
30        PrevLineB = PrevLineA
40        PrevLineA = CurrentLine
50        CurrentLine = FullFile(ff)
60        LineType = Trim(Right(CurrentLine, 3))
70        SetActiveLineType
80        If ActiveLineType <= 19 Then
90            RetrieveUsers
100           If Attacker <> "" And Defender <> "" Then
110               FoundInstance = False
120               For i = 0 To UBound(BattleStats)
130                   If BattleStats(i).Attacker = Attacker And BattleStats(i).Defender = Defender Then
140                       FoundInstance = True
150                       AddBattleStats i, False
160                       Exit For
170                   End If
180               Next
190               If FoundInstance = False Then
200                   For i = 0 To UBound(BattleStats)
210                       With BattleStats(i)
220                           If .Attacker = "" Then
230                               .BattleID = BattleID
240                               .Attacker = Attacker
250                               .Defender = Defender
260                               AddBattleStats i, False
270                               Exit For
280                           End If
290                       End With
300                   Next
310               End If

                  'Reverse Attacker/Defender to add to Evasion/HP Recovered
320               FoundInstance = False
330               For i = 0 To UBound(BattleStats)
340                   If BattleStats(i).Attacker = Defender And BattleStats(i).Defender = Attacker Then
350                       FoundInstance = True
360                       AddBattleStats i, True
370                       Exit For
380                   End If
390               Next
400               If FoundInstance = False Then
410                   For i = 0 To UBound(BattleStats)
420                       With BattleStats(i)
430                           If .Attacker = "" Then
440                               .BattleID = BattleID
450                               .Attacker = Defender
460                               .Defender = Attacker
470                               AddBattleStats i, True
480                               Exit For
490                           End If
500                       End With
510                   Next
520               End If
530           End If
540       ElseIf ActiveLineType = 20 Then
550           ReadFishing
560       ElseIf ActiveLineType = 30 Then
570           ReadLoot
580       ElseIf ActiveLineType = 31 Then
590           ReadPlayerLoot
600       ElseIf ActiveLineType = 32 Then
610           ReadGil
620       ElseIf ActiveLineType = 40 Then
630           ReadExp
640       ElseIf ActiveLineType = 50 Then
650           ReadTimes
660       ElseIf ActiveLineType = 51 Then
670           ReadCraft
680       ElseIf ActiveLineType = 52 Then
690           SetTimeA
700       ElseIf ActiveLineType = 53 Then
710           SetTimeB
720       ElseIf ActiveLineType = 54 Then
730           ReadCraftB
740       ElseIf ActiveLineType = 60 Then
750           GetMonsterCheck
760       ElseIf ActiveLineType = 70 Then
770           ChatText(UBound(ChatText)) = CurrentLine
780           ReDim Preserve ChatText(UBound(ChatText) + 1)
790       ElseIf ActiveLineType = 80 Then
800           ParserCommand
810       ElseIf ActiveLineType = 90 Then
820           CurrentFight = ""
830           RetrieveUsers
  
              'Read existing battles with this enemy and correct their BattleID
840           For i = 0 To UBound(BattleStats)
850               If BattleStats(i).Defender = Defender And BattleStats(i).BattleID <> BattleID Then
860                   BattleStats(i).BattleID = BattleID
870               End If
880           Next
  
890           GenerateReports False
  
              'Add battles to listing for editing
900           If FightComment = "" Then
910               listResults.AddItem BattleID & "-" & Defender
920           Else
930               listResults.AddItem BattleID & "-" & Defender & " - " & FightComment
940           End If
950           listResults.Selected(listResults.ListCount - 1) = True
960           FoundInstance = False
970           For i = 0 To comboMOB.ListCount - 1
980               If comboMOB.List(i) = Defender Then
990                   FoundInstance = True
1000                  Exit For
1010              End If
1020          Next
1030          If FoundInstance = False Then
1040              comboMOB.AddItem Defender
1050          End If
  
              'Clear all battle stats for this fight
1060          FightComment = ""
1070          For i = 0 To UBound(BattleStats)
1080              If BattleStats(i).Defender = Defender Then
1090                  ClearBattleStats i, False
1100                  ClearBattleStats i, True
1110              End If
1120          Next
1130          BattleID = BattleID + 1
  
              'Clear out any fights that had no death line and are at least 5 fights old.
1140          For i = 0 To UBound(BattleStats)
1150              With BattleStats(i)
1160                  .Battles = .Battles + 1
1170                  If .Battles >= 6 Then
1180                      ClearBattleStats i, False
1190                      ClearBattleStats i, True
1200                  End If
1210              End With
1220          Next
1230      End If
1240  Next
1250  Exit Sub
Err_Handler:
1260  HasErrors = True
1270  ErrorCount = ErrorCount + 1
1280  ReportError = "Error: " & Err.Number & vbNewLine & "Source: ParseLog" & vbNewLine & "Description: " & Err.Description & vbNewLine & "Line: " & Erl & vbNewLine & "FFXI Log Line: " & CurrentLine & vbNewLine & "Previous FFXI Log Line: " & PrevLineA & vbNewLine & vbNewLine
1290  Print #ErrorFile, ReportError
1300  If ErrorCount >= 25 Then
1310      lblStatus.Caption = "Errors: " & HasErrors & " - Too many errors - Parsing stopped for this log."
1320      Exit Sub
1330  Else
1340      Resume Next
1350  End If
End Sub
Private Sub ParserCommand()
10    On Error GoTo Err_Handler
      Dim MyPos As Integer, MyPos2 As Integer, i As Integer
      Dim SaveFileName As String, JobList As String
      Dim SaveReportTXT As Boolean

      Dim CraftFailed As String

20    If (Left$(LCase(CurrentLine), 18) = "parser start dps") Then
30        ReadDPS_Start = True
40        BeginDPS = True
50    ElseIf (Left$(LCase(CurrentLine), 18) = "parser direction") Then 'Crafting Direction' /echo parser direction NE
60        MyPos = InStrRev(CurrentLine, " ")
70        CraftDirection = Mid(CurrentLine, 20, MyPos - 20)
80    ElseIf (Left$(LCase(CurrentLine), 15) = "parser failed") Then 'Failed Item' /echo parser failed blah
90        MyPos = InStrRev(CurrentLine, " ")
100       CraftFailed = Mid(CurrentLine, 17, MyPos - 17)
110       If CraftingCSV(UBound(CraftingCSV)).Result <> "" Then
120           CraftingCSV(UBound(CraftingCSV)).Result = "Failure-" & CraftFailed
130       End If
140   ElseIf (Left$(LCase(CurrentLine), 17) = "parser critical") Then 'Critical Failure' /echo parser critical
150       CraftingCSV(UBound(CraftingCSV)).CriticalFailure = True
160   ElseIf (Left$(LCase(CurrentLine), 18) = "parser start exp") Then
170       ReadEXP_Start = True
180   ElseIf (Left$(LCase(CurrentLine), 20) = "parser start fight") Then
190       Read_Start = True
200   ElseIf (Left$(LCase(CurrentLine), 19) = "parser stop fight") Then
210       Read_Stop = True
220   ElseIf (Left$(LCase(CurrentLine), 17) = "parser stop dps") Then
230       BeginDPS = False
240       ReadDPS_Stop = True
250   ElseIf (Left$(LCase(CurrentLine), 18) = "parser clear dps") Then
260       BeginDPS = False
270       ReadDPS_Stop = False
280       ReadDPS_Start = False
290       Erase DPS
300   ElseIf (Left$(LCase(CurrentLine), 17) = "parser stop exp") Then
310       ReadEXP_Stop = True
320   ElseIf (Left$(LCase(CurrentLine), 19) = "parser start fish") Then
330       FishRPT
340       Erase FishFound
350       ReDim FishFound(0)
360       ReadFISH_Start = True
370   ElseIf (Left$(LCase(CurrentLine), 18) = "parser stop fish") Then
380       FishRPT
390       Erase FishFound
400       ReDim FishFound(0)
410   ElseIf (Left$(LCase(CurrentLine), 13) = "parser beep") Then
420       Call Beep(100, 100)
430       Call Beep(100, 100)
440   ElseIf (Left$(LCase(CurrentLine), 14) = "parser timer") Then
450       MyPos = InStr(1, CurrentLine, "'")
460       MyPos2 = InStr(MyPos + 1, CurrentLine, "'")
470       timerBeepAmt = Mid$(CurrentLine, MyPos + 1, MyPos2 - (MyPos + 1))
480       timerLength = Mid(CurrentLine, 16, 8)
490       ReadTimer = True
500   ElseIf (Left$(LCase(CurrentLine), 13) = "parser save") Then
510       MyPos = InStr(1, CurrentLine, ".rtf")
520       SaveReportTXT = False
530       If MyPos = 0 Then
540           MyPos = InStr(1, CurrentLine, ".txt")
550           SaveReportTXT = True
560       End If
570       MyPos2 = InStrRev(CurrentLine, " ", MyPos)
580       SaveFileName = Mid$(CurrentLine, MyPos2 + 1, (MyPos + 4) - (MyPos2 + 1))
590       If InStr(1, LCase(CurrentLine), "save report") Then
600           If SaveReportTXT Then
610             RTB_Report.SaveFile SaveFileName, rtfText
620           Else
630             RTB_Report.SaveFile SaveFileName, rtfRTF
640           End If
650       ElseIf InStr(1, LCase(CurrentLine), "save player") Then
660           If InStr(1, LCase(CurrentLine), "save player1") Then
670               comboUser.ListIndex = 0
680           ElseIf InStr(1, LCase(CurrentLine), "save player2") Then
690               comboUser.ListIndex = 1
700           ElseIf InStr(1, LCase(CurrentLine), "save player3") Then
710               comboUser.ListIndex = 2
720           ElseIf InStr(1, LCase(CurrentLine), "save player4") Then
730               comboUser.ListIndex = 3
740           ElseIf InStr(1, LCase(CurrentLine), "save player5") Then
750               comboUser.ListIndex = 4
760           ElseIf InStr(1, LCase(CurrentLine), "save player6") Then
770               comboUser.ListIndex = 5
780           End If
790           comboUser_Click
800           If SaveReportTXT Then
810               RTB_User.SaveFile SaveFileName, rtfText
820           Else
830               RTB_User.SaveFile SaveFileName, rtfRTF
840           End If
850       ElseIf InStr(1, LCase(CurrentLine), "save summary") Then
860           If SaveReportTXT Then
870               RTB_Averages.SaveFile SaveFileName, rtfText
880           Else
890               RTB_Averages.SaveFile SaveFileName, rtfRTF
900           End If
910       ElseIf InStr(1, LCase(CurrentLine), "save details") Then
920           If SaveReportTXT Then
930               RTB_Details.SaveFile SaveFileName, rtfText
940           Else
950               RTB_Details.SaveFile SaveFileName, rtfRTF
960           End If
970       End If
980   ElseIf (Left$(LCase(CurrentLine), 16) = "parser player ") Then
          Dim Player As String, Job As String, SubJob As String, Level As String, UserName As String, PlayerCount As Integer, FoundPlayer As Boolean
          'parser player spyle warrior ninja 75-71 ce'
990       MyPos = InStr(17, CurrentLine, " ")
1000      Player = Mid(CurrentLine, 17, MyPos - 17)
1010      Player = UCase(Left(Player, 1)) & LCase(Mid(Player, 2))
1020      MyPos = InStr(MyPos, CurrentLine, " ")
1030      MyPos2 = InStr(MyPos + 1, CurrentLine, " ")
1040      Job = Mid(CurrentLine, MyPos + 1, MyPos2 - (MyPos + 1))
1050      MyPos = InStr(MyPos2, CurrentLine, " ")
1060      MyPos2 = InStr(MyPos + 1, CurrentLine, " ")
1070      SubJob = Mid(CurrentLine, MyPos + 1, MyPos2 - (MyPos + 1))
1080      MyPos = InStr(MyPos2, CurrentLine, " ")
1090      MyPos2 = InStr(MyPos + 1, CurrentLine, " ")
1100      Level = Mid(CurrentLine, MyPos + 1, MyPos2 - (MyPos + 1))
    
1110      If LCase(Job) = "whm" Or LCase(Job) = "whitemage" Then
1120        Job = "White Mage"
1130      ElseIf LCase(Job) = "brd" Or LCase(Job) = "bard" Then
1140        Job = "Bard"
1150      ElseIf LCase(Job) = "blm" Or LCase(Job) = "blackmage" Then
1160        Job = "Black Mage"
1170      ElseIf LCase(Job) = "drk" Or LCase(Job) = "darkknight" Then
1180        Job = "Dark Knight"
1190      ElseIf LCase(Job) = "drg" Or LCase(Job) = "dragoon" Then
1200        Job = "Dragoon"
1210      ElseIf LCase(Job) = "mnk" Or LCase(Job) = "monk" Then
1220        Job = "Monk"
1230      ElseIf LCase(Job) = "nin" Or LCase(Job) = "ninja" Then
1240        Job = "Ninja"
1250      ElseIf LCase(Job) = "pld" Or LCase(Job) = "pal" Or LCase(Job) = "paladin" Then
1260        Job = "Paladin"
1270      ElseIf LCase(Job) = "rng" Or LCase(Job) = "ran" Or LCase(Job) = "ranger" Then
1280        Job = "Ranger"
1290      ElseIf LCase(Job) = "rdm" Or LCase(Job) = "red" Or LCase(Job) = "redmage" Then
1300        Job = "Red Mage"
1310      ElseIf LCase(Job) = "sam" Then
1320        Job = "Samurai"
1330      ElseIf LCase(Job) = "smn" Or LCase(Job) = "sum" Or LCase(Job) = "summoner" Then
1340        Job = "Summoner"
1350      ElseIf LCase(Job) = "thf" Or LCase(Job) = "thief" Then
1360        Job = "Thief"
1370      ElseIf LCase(Job) = "war" Or LCase(Job) = "warrior" Then
1380        Job = "Warrior"
1390      End If
    

1400      If LCase(SubJob) = "whm" Or LCase(SubJob) = "whitemage" Then
1410        SubJob = "White Mage"
1420      ElseIf LCase(SubJob) = "brd" Or LCase(SubJob) = "bard" Then
1430        SubJob = "Bard"
1440      ElseIf LCase(SubJob) = "blm" Or LCase(SubJob) = "blackmage" Then
1450        SubJob = "Black Mage"
1460      ElseIf LCase(SubJob) = "drk" Or LCase(SubJob) = "darkknight" Then
1470        SubJob = "Dark Knight"
1480      ElseIf LCase(SubJob) = "drg" Or LCase(SubJob) = "dragoon" Then
1490        SubJob = "Dragoon"
1500      ElseIf LCase(SubJob) = "mnk" Or LCase(SubJob) = "monk" Then
1510        SubJob = "Monk"
1520      ElseIf LCase(SubJob) = "nin" Or LCase(SubJob) = "ninja" Then
1530        SubJob = "Ninja"
1540      ElseIf LCase(SubJob) = "pld" Or LCase(SubJob) = "pal" Or LCase(SubJob) = "paladin" Then
1550        SubJob = "Paladin"
1560      ElseIf LCase(SubJob) = "rng" Or LCase(SubJob) = "ran" Or LCase(SubJob) = "ranger" Then
1570        SubJob = "Ranger"
1580      ElseIf LCase(SubJob) = "rdm" Or LCase(SubJob) = "red" Or LCase(SubJob) = "redmage" Then
1590        SubJob = "Red Mage"
1600      ElseIf LCase(SubJob) = "sam" Then
1610        SubJob = "Samurai"
1620      ElseIf LCase(SubJob) = "smn" Or LCase(SubJob) = "sum" Or LCase(SubJob) = "summoner" Then
1630        SubJob = "Summoner"
1640      ElseIf LCase(SubJob) = "thf" Or LCase(SubJob) = "thief" Then
1650        SubJob = "Thief"
1660      ElseIf LCase(SubJob) = "war" Or LCase(SubJob) = "warrior" Then
1670        SubJob = "Warrior"
1680      End If

1690      JobList = "White Mage, Bard, Black Mage, Dark Knight, Dragoon, Monk, Ninja, Paladin, Ranger, Red Mage, Samurai, Summoner, Thief, Warrior"
1700      If InStr(1, JobList, Job) <> 0 And InStr(1, JobList, SubJob) <> 0 Then
1710          PlayerCount = GetSetting(App.Title, "Online_Setup", "Count", "5")
1720          For i = 0 To PlayerCount
1730              UserName = GetSetting(App.Title, "Online_Setup", "Name" & i, "")
1740              If LCase(UserName) = LCase(Player) Then
1750                  SaveSetting App.Title, "Online_Setup", "Job" & i, Job
1760                  SaveSetting App.Title, "Online_Setup", "SubJob" & i, SubJob
1770                  SaveSetting App.Title, "Online_Setup", "Level" & i, Level
1780                  FoundPlayer = True
1790                  Exit For
1800              Else
1810                  FoundPlayer = False
1820              End If
1830          Next
  
1840          If FoundPlayer = False Then
1850              For i = 0 To PlayerCount
1860                  UserName = GetSetting(App.Title, "Online_Setup", "Name" & i, "")
1870                  If UserName = "" Then
1880                      SaveSetting App.Title, "Online_Setup", "Name" & i, Player
1890                      SaveSetting App.Title, "Online_Setup", "Job" & i, Job
1900                      SaveSetting App.Title, "Online_Setup", "SubJob" & i, SubJob
1910                      SaveSetting App.Title, "Online_Setup", "Level" & i, Level
1920                      FoundPlayer = True
1930                      Exit For
1940                  Else
1950                      FoundPlayer = False
1960                  End If
1970              Next
1980          End If
  
1990          If FoundPlayer = False Then
2000              SaveSetting App.Title, "Online_Setup", "Count", PlayerCount + 1
2010              SaveSetting App.Title, "Online_Setup", "Name" & PlayerCount + 1, Player
2020              SaveSetting App.Title, "Online_Setup", "Job" & PlayerCount + 1, Job
2030              SaveSetting App.Title, "Online_Setup", "SubJob" & PlayerCount + 1, SubJob
2040              SaveSetting App.Title, "Online_Setup", "Level" & PlayerCount + 1, Level
2050          End If
2060      End If
2070  ElseIf (Left$(LCase(CurrentLine), 15) = "parser window") Then
2080      If IsNumeric(Mid(CurrentLine, 17, 1)) = True Then
2090          If Mid(CurrentLine, 17, 1) < 8 Then
2100              comboDisplay.ListIndex = CDbl(Mid(CurrentLine, 17, 1) - 1)
2110              comboDisplay_Click
2120          End If
2130      End If
2140  ElseIf (Left$(LCase(CurrentLine), 18) = "parser comment '") Then
2150      MyPos = InStr(1, CurrentLine, "'")
2160      MyPos2 = InStr(MyPos + 1, CurrentLine, "'")
2170      If MyPos2 <> 0 Then
2180          FightComment = Mid$(CurrentLine, MyPos + 1, MyPos2 - (MyPos + 1))
2190      Else
2200          FightComment = "Invalid format."
2210      End If
2220  ElseIf (Left$(LCase(CurrentLine), 23) = "parser fish comment '") Then
2230      MyPos = InStr(1, CurrentLine, "'")
2240      MyPos2 = InStr(MyPos + 1, CurrentLine, "'")
2250      If MyPos2 <> 0 Then
2260          FishComment = Mid$(CurrentLine, MyPos + 1, MyPos2 - (MyPos + 1))
2270      Else
2280          FishComment = "Invalid format."
2290      End If
2300  End If

2310  Exit Sub
Err_Handler:
2320  HasErrors = True
2330  ErrorCount = ErrorCount + 1
2340  ReportError = "Error: " & Err.Number & vbNewLine & "Source: ParserCommand" & vbNewLine & "Description: " & Err.Description & vbNewLine & "Line: " & Erl & vbNewLine & "FFXI Log Line: " & CurrentLine & vbNewLine & "Previous FFXI Log Line: " & PrevLineA & vbNewLine & vbNewLine
2350  Print #ErrorFile, ReportError
2360  If ErrorCount >= 25 Then
2370      lblStatus.Caption = "Errors: " & HasErrors & " - Too many errors - Parsing stopped for this log."
2380      Exit Sub
2390  Else
2400      Resume Next
2410  End If
End Sub

Private Sub ReadCraft()
10    On Error GoTo Err_Handler

      Dim i, MyPos, MyPos2, MyResult As String, FoundIt As Boolean
20    MyPos = InStr(1, CurrentLine, " ")
30    MyPos2 = InStrRev(CurrentLine, ".")
40    If MyPos = 0 Then
50        MyResult = "Failure-" & LCase(PrevResult)
60    Else
70        MyResult = LCase(Mid(CurrentLine, MyPos + 3, MyPos2 - (MyPos + 3)))
80        PrevResult = MyResult
90    End If
100   If Right(MyResult, 1) = "1" Then
110       MyResult = Mid(MyResult, Len(MyResult) - 1)
120   End If

130   FoundIt = False
140   For i = 0 To UBound(CraftingCSV)
150       With CraftingCSV(i)
160           If .Result = MyResult And .DayType = CurrentDay And .MoonPerc = CurrentPerc And .MoonPhase = CurrentMoon And .CurrentTime = CurrentTime And .Direction = CraftDirection Then
170               .Count = .Count + 1
180               FoundIt = True
190               Exit For
200           End If
210       End With
220   Next

230   If Not FoundIt Then
240       If CraftingCSV(0).Result <> "" Then
250           ReDim Preserve CraftingCSV(UBound(CraftingCSV) + 1)
260       End If
270       i = UBound(CraftingCSV)
280       With CraftingCSV(i)
290           .Result = MyResult
300           .DayType = CurrentDay
310           .MoonPerc = CurrentPerc
320           .MoonPhase = CurrentMoon
330           .Direction = CraftDirection
340           .CurrentTime = CurrentTime
350           .Count = 1
360       End With
370   End If
    
380   Exit Sub
Err_Handler:
390   HasErrors = True
400   ErrorCount = ErrorCount + 1
410   ReportError = "Error: " & Err.Number & vbNewLine & "Source: ReadCraft" & vbNewLine & "Description: " & Err.Description & vbNewLine & "Line: " & Erl & vbNewLine & "FFXI Log Line: " & CurrentLine & vbNewLine & "Previous FFXI Log Line: " & PrevLineA & vbNewLine & vbNewLine
420   Print #ErrorFile, ReportError
430   If ErrorCount >= 25 Then
440       lblStatus.Caption = "Errors: " & HasErrors & " - Too many errors - Parsing stopped for this log."
450       Exit Sub
460   Else
470       Resume Next
480   End If
End Sub

Private Sub ReadCraftB()
10    On Error GoTo Err_Handler

      Dim i, MyPos, MyPos2, MyResult As String, FoundIt As Boolean
20    If InStr(1, CurrentLine, "was lost.") Then
30        MyPos = InStr(1, CurrentLine, "y")
40        MyPos2 = InStrRev(CurrentLine, " ")
50        MyResult = Mid(CurrentLine, MyPos + 3, MyPos2 - (MyPos + 3))
60        PrevResult = MyResult
70        If Right(MyResult, 1) = "1" Then
80      MyResult = Mid(MyResult, Len(MyResult) - 1)
90        End If
100       MyResult = "Failure-" & LCase(MyResult)
110   Else
120       MyPos = InStr(1, CurrentLine, " ")
130       MyPos2 = InStrRev(CurrentLine, ".")
140       MyResult = LCase(Mid(CurrentLine, MyPos + 3, MyPos2 - (MyPos + 3)))
150       PrevResult = MyResult
160       If Right(MyResult, 1) = "1" Then
170     MyResult = Mid(MyResult, Len(MyResult) - 1)
180       End If
190   End If

200   FoundIt = False
210   For i = 0 To UBound(CraftingCSV)
220       With CraftingCSV(i)
230           If .Result = MyResult And .DayType = CurrentDay And .MoonPerc = CurrentPerc And .MoonPhase = CurrentMoon And .CurrentTime = CurrentTime And .Direction = CraftDirection Then
240               .Count = .Count + 1
250               FoundIt = True
260               Exit For
270           End If
280       End With
290   Next

300   If Not FoundIt Then
310       If CraftingCSV(0).Result <> "" Then
320           ReDim Preserve CraftingCSV(UBound(CraftingCSV) + 1)
330       End If
340       i = UBound(CraftingCSV)
350       With CraftingCSV(i)
360           .Result = MyResult
370           .DayType = CurrentDay
380           .MoonPerc = CurrentPerc
390           .MoonPhase = CurrentMoon
400           .Direction = CraftDirection
410           .CurrentTime = CurrentTime
420           .Count = 1
430       End With
440   End If
    
450   Exit Sub
Err_Handler:
460   HasErrors = True
470   ErrorCount = ErrorCount + 1
480   ReportError = "Error: " & Err.Number & vbNewLine & "Source: ReadCraft" & vbNewLine & "Description: " & Err.Description & vbNewLine & "Line: " & Erl & vbNewLine & "FFXI Log Line: " & CurrentLine & vbNewLine & "Previous FFXI Log Line: " & PrevLineA & vbNewLine & vbNewLine
490   Print #ErrorFile, ReportError
500   If ErrorCount >= 25 Then
510       lblStatus.Caption = "Errors: " & HasErrors & " - Too many errors - Parsing stopped for this log."
520       Exit Sub
530   Else
540       Resume Next
550   End If
End Sub


Private Sub ReadExp()
10    On Error GoTo Err_Handler
      Dim MyPos As Integer, MyPos2 As Integer
20    If InStr(1, CurrentLine, "experience points.") Or InStr(1, CurrentLine, "limit points.") Then
30        MyPos = InStr(1, CurrentLine, "gains ")
40        MyPos2 = InStr(1, CurrentLine, "exp")
50        If MyPos2 = 0 Then
60            MyPos2 = InStr(1, CurrentLine, "limit")
70        End If
80        If StopEXP = False Then
90            TotalExp = TotalExp + CDbl(Mid$(CurrentLine, MyPos + 6, MyPos2 - (MyPos + 7)))
100           ChainExp(ExpType, 0) = ChainExp(ExpType, 0) + CDbl(Mid$(CurrentLine, MyPos + 6, MyPos2 - (MyPos + 7)))
110           ChainExp(ExpType, 1) = ChainExp(ExpType, 1) + 1
120           ExpType = 0
130       End If
140   ElseIf InStr(1, CurrentLine, "EXP chain #") Then
150       MyPos = InStr(1, CurrentLine, "#")
160       ExpType = CDbl(Mid(CurrentLine, MyPos + 1, 1))
170   ElseIf InStr(1, CurrentLine, "Limit chain #") Then
180       MyPos = InStr(1, CurrentLine, "#")
190       ExpType = CDbl(Mid(CurrentLine, MyPos + 1, 1))
200   End If


210   Exit Sub
Err_Handler:
220   HasErrors = True
230   ErrorCount = ErrorCount + 1
240   ReportError = "Error: " & Err.Number & vbNewLine & "Source: ReadExp" & vbNewLine & "Description: " & Err.Description & vbNewLine & "Line: " & Erl & vbNewLine & "FFXI Log Line: " & CurrentLine & vbNewLine & "Previous FFXI Log Line: " & PrevLineA & vbNewLine & vbNewLine
250   Print #ErrorFile, ReportError
260   If ErrorCount >= 25 Then
270       lblStatus.Caption = "Errors: " & HasErrors & " - Too many errors - Parsing stopped for this log."
280       Exit Sub
290   Else
300       Resume Next
310   End If
End Sub

Private Sub ReadFishing()
10    On Error GoTo Err_Handler
      Dim MyPos As Integer, MyPos2 As Integer
      Dim FishItem As String, PrevFish As Boolean, lf
20    If InStr(1, LCase(CurrentLine), "obtained: ") And LineType = "94" Then
30        MyPos = InStr(1, CurrentLine, "obtained: ")
40        MyPos2 = InStr(1, CurrentLine, ".")
50        FishItem = Mid$(CurrentLine, MyPos + 15, MyPos2 - (MyPos + 17))
60        PrevFish = False
70        For lf = 0 To UBound(FishFound)
80          If InStr(1, FishFound(lf), FishItem) Then
90              PrevFish = True
100             Exit For
110         End If
120       Next
130       If PrevFish Then
140         MyPos = InStr(1, FishFound(lf), " - ")
150         FishFound(lf) = CDbl(Left$(FishFound(lf), MyPos)) + 1 & " - " & FishItem
160       Else
170         FishFound(UBound(FishFound)) = "1 - " & FishItem
180         ReDim Preserve FishFound(UBound(FishFound) + 1)
190       End If
200   ElseIf InStr(1, LCase(CurrentLine), "you lost your catch.") Then
210       FishItem = "catches lost"
220       PrevFish = False
230       For lf = 0 To UBound(FishFound)
240         If InStr(1, FishFound(lf), FishItem) Then
250             PrevFish = True
260             Exit For
270         End If
280       Next
290       If PrevFish Then
300         MyPos = InStr(1, FishFound(lf), " - ")
310         FishFound(lf) = CDbl(Left$(FishFound(lf), MyPos)) + 1 & " - " & FishItem
320       Else
330         FishFound(UBound(FishFound)) = "1 - " & FishItem
340         ReDim Preserve FishFound(UBound(FishFound) + 1)
350       End If
360   ElseIf InStr(1, LCase(CurrentLine), "you didn't catch anything.") Then
370       FishItem = "didn't catch anything"
380       PrevFish = False
390       For lf = 0 To UBound(FishFound)
400         If InStr(1, FishFound(lf), FishItem) Then
410             PrevFish = True
420             Exit For
430         End If
440       Next
450       If PrevFish Then
460         MyPos = InStr(1, FishFound(lf), " - ")
470         FishFound(lf) = CDbl(Left$(FishFound(lf), MyPos)) + 1 & " - " & FishItem
480       Else
490         FishFound(UBound(FishFound)) = "1 - " & FishItem
500         ReDim Preserve FishFound(UBound(FishFound) + 1)
510       End If
520   End If

530   Exit Sub
Err_Handler:
540   HasErrors = True
550   ErrorCount = ErrorCount + 1
560   ReportError = "Error: " & Err.Number & vbNewLine & "Source: ReadFishing" & vbNewLine & "Description: " & Err.Description & vbNewLine & "Line: " & Erl & vbNewLine & "FFXI Log Line: " & CurrentLine & vbNewLine & "Previous FFXI Log Line: " & PrevLineA & vbNewLine & vbNewLine
570   Print #ErrorFile, ReportError
580   If ErrorCount >= 25 Then
590       lblStatus.Caption = "Errors: " & HasErrors & " - Too many errors - Parsing stopped for this log."
600       Exit Sub
610   Else
620       Resume Next
630   End If
End Sub

Private Sub ReadGil()
10    On Error GoTo Err_Handler
      Dim MyPos As Integer, MyPos2 As Integer
      Dim LootItem As String, GilAmt As Long, PrevLoot As Boolean, lf As Integer
20    MyPos = InStr(1, CurrentLine, " obtains ")
30    MyPos2 = InStr(1, CurrentLine, " gil.")
40    LootItem = "Gil"
50    GilAmt = CDbl(Mid$(CurrentLine, MyPos + 9, MyPos2 - (MyPos + 9)))
60    PrevLoot = False
70    For lf = 0 To UBound(LootFound)
80      If InStr(1, LootFound(lf), LootItem) Then
90          PrevLoot = True
100         Exit For
110     End If
120   Next
130   If PrevLoot Then
140     MyPos = InStr(1, LootFound(lf), " - ")
150     LootFound(lf) = CDbl(Left$(LootFound(lf), MyPos)) + GilAmt & " - " & LootItem
160   Else
170     LootFound(UBound(LootFound)) = GilAmt & " - " & LootItem
180     ReDim Preserve LootFound(UBound(LootFound) + 1)
190   End If

200   Exit Sub
Err_Handler:
210   HasErrors = True
220   ErrorCount = ErrorCount + 1
230   ReportError = "Error: " & Err.Number & vbNewLine & "Source: ReadGil" & vbNewLine & "Description: " & Err.Description & vbNewLine & "Line: " & Erl & vbNewLine & "FFXI Log Line: " & CurrentLine & vbNewLine & "Previous FFXI Log Line: " & PrevLineA & vbNewLine & vbNewLine
240   Print #ErrorFile, ReportError
250   If ErrorCount >= 25 Then
260       lblStatus.Caption = "Errors: " & HasErrors & " - Too many errors - Parsing stopped for this log."
270       Exit Sub
280   Else
290       Resume Next
300   End If
End Sub

Private Sub ReadLoot()
10    On Error GoTo Err_Handler
      Dim LootItem As String, PrevLoot As Boolean, lf
      Dim MyPos As Integer, MyPos2 As Integer
20    MyPos = InStr(1, CurrentLine, " ")
30    MyPos2 = InStr(MyPos + 1, CurrentLine, " ")
40    LootItem = Mid$(CurrentLine, MyPos + 3, MyPos2 - (MyPos + 3))
50    PrevLoot = False
60    For lf = 0 To UBound(LootFound)
70      If InStr(1, LootFound(lf), LootItem) Then
80          PrevLoot = True
90          Exit For
100     End If
110   Next
120   If PrevLoot Then
130     MyPos = InStr(1, LootFound(lf), " - ")
140     LootFound(lf) = CDbl(Left$(LootFound(lf), MyPos)) + 1 & " - " & LootItem
150   Else
160     LootFound(UBound(LootFound)) = "1 - " & LootItem
170     ReDim Preserve LootFound(UBound(LootFound) + 1)
180   End If

190   Exit Sub
Err_Handler:
200   HasErrors = True
210   ErrorCount = ErrorCount + 1
220   ReportError = "Error: " & Err.Number & vbNewLine & "Source: ReadLoot" & vbNewLine & "Description: " & Err.Description & vbNewLine & "Line: " & Erl & vbNewLine & "FFXI Log Line: " & CurrentLine & vbNewLine & "Previous FFXI Log Line: " & PrevLineA & vbNewLine & vbNewLine
230   Print #ErrorFile, ReportError
240   If ErrorCount >= 25 Then
250       lblStatus.Caption = "Errors: " & HasErrors & " - Too many errors - Parsing stopped for this log."
260       Exit Sub
270   Else
280       Resume Next
290   End If
End Sub

Private Sub ReadPlayerLoot()
10    On Error GoTo Err_Handler
      Dim PrevLoot As Boolean
      Dim LootPlayer As String, LootItem As String
      Dim MyPos As Integer, MyPos2 As Integer, lf As Integer
20    MyPos = InStr(1, CurrentLine, " ")
30    LootPlayer = Mid(CurrentLine, 5, MyPos - 5)
40    MyPos = InStr(1, CurrentLine, " ")
50    MyPos2 = InStr(MyPos + 1, CurrentLine, ".")
60    LootItem = Mid$(CurrentLine, MyPos + 3, MyPos2 - (MyPos + 3))
70    PrevLoot = False
80    For lf = 0 To UBound(PlayerLoot)
90      If InStr(1, PlayerLoot(lf), LootItem & ";" & LootPlayer) Then
100         PrevLoot = True
110         Exit For
120     End If
130   Next
140   If PrevLoot Then
150     MyPos = InStr(1, PlayerLoot(lf), " - ")
160     PlayerLoot(lf) = CDbl(Left$(PlayerLoot(lf), MyPos)) + 1 & " - " & LootItem & ";" & LootPlayer
170   Else
180     PlayerLoot(UBound(PlayerLoot)) = "1 - " & LootItem & ";" & LootPlayer
190     ReDim Preserve PlayerLoot(UBound(PlayerLoot) + 1)
200   End If

210   Exit Sub
Err_Handler:
220   HasErrors = True
230   ErrorCount = ErrorCount + 1
240   ReportError = "Error: " & Err.Number & vbNewLine & "Source: ReadPlayerLoot" & vbNewLine & "Description: " & Err.Description & vbNewLine & "Line: " & Erl & vbNewLine & "FFXI Log Line: " & CurrentLine & vbNewLine & "Previous FFXI Log Line: " & PrevLineA & vbNewLine & vbNewLine
250   Print #ErrorFile, ReportError
260   If ErrorCount >= 25 Then
270       lblStatus.Caption = "Errors: " & HasErrors & " - Too many errors - Parsing stopped for this log."
280       Exit Sub
290   Else
300       Resume Next
310   End If
End Sub

Private Sub ReadTimes()
10    On Error GoTo Err_Handler
      Dim i As Integer, MyPos As Integer
      Dim TotalSeconds
20    If ReadTimer Then
30        For i = 0 To UBound(TimerStart)
40            If TimerStart(i, 0) = "" Then
50                TimerStart(i, 0) = Replace(Mid(CurrentLine, 10, Len(CurrentLine) - 13), ".", "")
60                TimerStart(i, 1) = timerLength
70                TimerStart(i, 2) = timerBeepAmt
80                Exit For
90            End If
100       Next
110       ReadTimer = False
120   End If
130   If Read_Start Then
140       EffTotals.ATK = 0
150       EffTotals.ATKTaken = 0
160       EffTotals.BasicDMG = 0
170       EffTotals.DMGTaken = 0
180       EffTotals.TotalDMG = 0
190       BeginDPS = True
200       StartTimeDPS = Replace(Mid(CurrentLine, 10, Len(CurrentLine) - 13), ".", "")
210       ReadDPS_Start = False
220       StopDPS = False
230       FightStartTime = Replace(Mid(CurrentLine, 10, Len(CurrentLine) - 13), ".", "")
240       Read_Start = False
250   ElseIf Read_Stop = True And FightStartTime <> Empty Then
260       BeginDPS = False
270       StopTimeDPS = Replace(Mid(CurrentLine, 10, Len(CurrentLine) - 13), ".", "")
280       If StartTimeDPS <> Empty Then
290           For i = 0 To UBound(DPS)
300               If DPS(i, 2) = "" Then DPS(i, 2) = "0"
310               DPS(i, 2) = CDbl(DPS(i, 2)) + CDbl(DateDiff("s", StartTimeDPS, StopTimeDPS))
320               DPS(i, 0) = DPS(i, 0)
330               DPS(i, 1) = DPS(i, 1)
340               DPS(i, 2) = DPS(i, 2)
350           Next
360       End If
370       ReadDPS_Stop = False
380       StopDPS = True
390       FightStopTime = Replace(Mid(CurrentLine, 10, Len(CurrentLine) - 13), ".", "")
400       With RTB_Report
410           TotalSeconds = DateDiff("s", FightStartTime, FightStopTime)
420           If TotalSeconds <> 0 Then
430               .SelStart = Len(.Text) - 2
440               .SelBold = True
450               .SelText = vbNewLine & "Efficiency" & vbNewLine
460               .SelBold = False
470               .SelColor = &HC00000
480               .SelBold = False
490               .SelText = "Time Taken: "
500               .SelColor = &HC0&
510               .SelText = Format(FightStopTime - FightStartTime, "HH:MM:SS") & vbNewLine
520               .SelColor = &HC00000
530               .SelBold = False
540               .SelText = "Total Damage Per Second: "
550               .SelColor = &HC0&
560               .SelText = Round(EffTotals.TotalDMG / TotalSeconds, 2) & vbNewLine
570               .SelColor = &HC00000
580               .SelBold = False
590               .SelText = "Basic Damage Per Second: "
600               .SelColor = &HC0&
610               .SelText = Round(EffTotals.BasicDMG / TotalSeconds, 2) & vbNewLine
620               .SelColor = &HC00000
630               .SelBold = False
640               .SelText = "Attacks Per Second: "
650               .SelColor = &HC0&
660               .SelText = Round((EffTotals.ATK) / TotalSeconds, 2) & vbNewLine
670               .SelColor = &HC00000
680               .SelBold = False
690               .SelText = "Damage Taken Per Second: "
700               .SelColor = &HC0&
710               .SelText = Round(EffTotals.DMGTaken / TotalSeconds, 2) & vbNewLine
720               .SelColor = &HC00000
730               .SelBold = False
740               .SelText = "Attacks Taken Per Second: "
750               .SelColor = &HC0&
760               .SelText = Round((EffTotals.ATKTaken) / TotalSeconds, 2) & vbNewLine
770               .SelStart = Len(.Text)
780           End If
790       End With
800       FightStartTime = Empty
810       FightStopTime = Empty
820       Read_Stop = False
830   End If
840   If ReadEXP_Start Then
850       Erase ChainExp
860       ExpType = 0
870       TotalExp = 0
880       StartTime = Replace(Mid(CurrentLine, 10, Len(CurrentLine) - 13), ".", "")
890       StopTime = Empty
900       ReadEXP_Start = False
910       StopEXP = False
920   ElseIf ReadEXP_Stop Then
930       StopTime = Replace(Mid(CurrentLine, 10, Len(CurrentLine) - 13), ".", "")
940       ReadEXP_Stop = False
950       StopEXP = True
960   End If
970   If ReadFISH_Start Then
980       MyPos = InStr(1, PrevLineB, ",")
990       If MyPos <> 0 Then
1000          FishHeader = Trim(Left(Mid(PrevLineB, MyPos + 2), Len(Mid(PrevLineB, MyPos + 2)) - 2))
1010          FishHeader = FishHeader & " - " & Left(Trim(Mid(PrevLineA, 3)), Len(Trim(Mid(PrevLineA, 3))) - 2)
1020          FishHeader = FishHeader & " - Earth: " & Trim(Mid(CurrentLine, 24, Len(CurrentLine) - 26))
1030      End If
1040      ReadFISH_Start = False
1050  End If
1060  If ReadDPS_Start Then
1070      StartTimeDPS = Replace(Mid(CurrentLine, 10, Len(CurrentLine) - 13), ".", "")
1080      ReadDPS_Start = False
1090      StopDPS = False
1100  ElseIf ReadDPS_Stop Then
1110      StopTimeDPS = Replace(Mid(CurrentLine, 10, Len(CurrentLine) - 13), ".", "")
1120      If StartTimeDPS <> Empty Then
1130          For i = 0 To UBound(DPS)
1140              If DPS(i, 2) = "" Then DPS(i, 2) = "0"
1150              DPS(i, 2) = CDbl(DPS(i, 2)) + CDbl(DateDiff("s", StartTimeDPS, StopTimeDPS))
1160              DPS(i, 0) = DPS(i, 0)
1170              DPS(i, 1) = DPS(i, 1)
1180              DPS(i, 2) = DPS(i, 2)
1190          Next
1200      End If
1210      ReadDPS_Stop = False
1220      StopDPS = True
1230      StartTimeDPS = Empty
1240      StopTimeDPS = Empty
1250  End If

1260  Exit Sub
Err_Handler:
1270  HasErrors = True
1280  ErrorCount = ErrorCount + 1
1290  ReportError = "Error: " & Err.Number & vbNewLine & "Source: ReadTimes" & vbNewLine & "Description: " & Err.Description & vbNewLine & "Line: " & Erl & vbNewLine & "FFXI Log Line: " & CurrentLine & vbNewLine & "Previous FFXI Log Line: " & PrevLineA & vbNewLine & vbNewLine
1300  Print #ErrorFile, ReportError
1310  If ErrorCount >= 25 Then
1320      lblStatus.Caption = "Errors: " & HasErrors & " - Too many errors - Parsing stopped for this log."
1330      Exit Sub
1340  Else
1350      Resume Next
1360  End If
End Sub


Private Sub RecalculateData()
10    On Error GoTo Err_Handler
20    mnuClear_Click
30    Screen.MousePointer = vbHourglass
      Dim i As Integer, p As Integer, b As Integer, c As Integer
      Dim FoundInstance As Boolean, IncludePlayer As Boolean

      Dim FullPercent As Long, PartPercent As Long
40    For i = 0 To listResults.ListCount - 1
50        If listResults.Selected(i) = True Then
60            FullPercent = FullPercent + 1
70        End If
80    Next

90    FullPercent = FullPercent * UBound(FullStats) - 1
If FullPercent = 0 Then Exit Sub
100   For i = 0 To listResults.ListCount - 1
110       If listResults.Selected(i) = True Then
120           For p = 0 To UBound(FullStats) - 1
130               PartPercent = PartPercent + 1
140                 lblStatus.Caption = CStr(Round((PartPercent / FullPercent) * 100, 2)) & "%"
150               DoEvents
160               IncludePlayer = False
170               For c = 0 To listPlayers.ListCount - 1
180                   If listPlayers.List(c) = FullStats(p).Attacker Or InStr(1, FullStats(p).Attacker, "SC:") <> 0 Then
190                       If listPlayers.Selected(c) = True Then
200                           IncludePlayer = True
210                       End If
220                       Exit For
230                   End If
240               Next
250               If FullStats(p).BattleID = i And IncludePlayer Then
260                   For b = 0 To UBound(BattleStats)
270                       If BattleStats(b).Attacker = FullStats(p).Attacker And BattleStats(b).Defender = FullStats(p).Defender Then
280                           CombineStats BattleStats(b), FullStats(p)
290                           BattleStats(b).Basic.List = BattleStats(b).Basic.List & ", " & FullStats(p).Basic.Damage
300                           Defender = FullStats(p).Defender
310                           Attacker = FullStats(p).Attacker
320                           CurrentFight = FullStats(p).Defender
330                       End If
340                   Next
350                   If FoundInstance = False Then
360                       For b = 0 To UBound(BattleStats)
370                           With BattleStats(b)
380                               If .Attacker = "" Then
390                                   CombineStats BattleStats(b), FullStats(p)
400                                   BattleStats(b).Basic.List = BattleStats(b).Basic.List & ", " & FullStats(p).Basic.Damage
410                                   Defender = FullStats(p).Defender
420                                   Attacker = FullStats(p).Attacker
430                                   CurrentFight = FullStats(p).Defender
440                                   Exit For
450                               End If
460                           End With
470                       Next
480                   End If
490               End If
500           Next
510           BattleID = i
520           If CurrentFight <> "" Then
530               GenerateReports True
540               CurrentFight = ""
550           End If
560           For b = 0 To UBound(BattleStats)
570               If BattleStats(b).BattleID = i Then
580                   ClearBattleStats b, False
590                   ClearBattleStats b, True
600               End If
610           Next
620       End If
630   Next
640   Screen.MousePointer = vbDefault

650   MsgBox "Done!", vbInformation, "Recalculate"

660   Exit Sub
Err_Handler:
670   HasErrors = True
680   ErrorCount = ErrorCount + 1
690   ReportError = "Error: " & Err.Number & vbNewLine & "Source: RecalculateData" & vbNewLine & "Description: " & Err.Description & vbNewLine & "Line: " & Erl & vbNewLine & "FFXI Log Line: " & CurrentLine & vbNewLine & "Previous FFXI Log Line: " & PrevLineA & vbNewLine & vbNewLine
700   Print #ErrorFile, ReportError
710   If ErrorCount >= 25 Then
720       lblStatus.Caption = "Errors: " & HasErrors & " - Too many errors - Parsing stopped for this log."
730       Exit Sub
740   Else
750       Resume Next
760   End If
End Sub

Public Sub ResetTimeFile(TheFile As String, m_Date As Date)
    Dim lngHandle As Long
    Dim udtFileTime As FILETIME
    Dim udtLocalTime As FILETIME
    Dim udtSystemTime As SYSTEMTIME

    udtSystemTime.wYear = Year(m_Date)
    udtSystemTime.wMonth = Month(m_Date)
    udtSystemTime.wDay = Day(m_Date)
    udtSystemTime.wDayOfWeek = Weekday(m_Date) - 1
    udtSystemTime.wHour = Hour(m_Date)
    udtSystemTime.wMinute = Minute(m_Date)
    udtSystemTime.wSecond = Second(m_Date)
    udtSystemTime.wMilliseconds = 0

    ' convert system time to local time
    SystemTimeToFileTime udtSystemTime, udtLocalTime
    ' convert local time to GMT
    LocalFileTimeToFileTime udtLocalTime, udtFileTime
    ' open the file to get the filehandle
    lngHandle = CreateFile(TheFile, GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)
    ' change date/time property of the file
    SetFileTime lngHandle, udtFileTime, udtFileTime, udtFileTime
    ' close the handle
    CloseHandle lngHandle
End Sub



Public Function ResizePart(Part As String, Size As Long) As String
Do Until TextWidth(Part) >= Size
    Part = Part & " "
Loop
ResizePart = Part
End Function


Private Sub RetrieveUsers()
10    On Error GoTo Err_Handler
20    Attacker = ""
30    Defender = ""
      Dim MyPos As Integer, MyPos2 As Integer
      Dim PreP1 As String

40    If ActiveLineType <= 3 Then
50        If ActiveLineType = 0 Then 'attacker Hits defender For
60            MyPos = InStr(1, CurrentLine, " hit")
70            Attacker = Mid(CurrentLine, 3, MyPos - 3)
80            MyPos = InStr(MyPos + 3, CurrentLine, " ")
90            MyPos2 = InStr(MyPos, CurrentLine, " for ")
100           Defender = Mid(CurrentLine, MyPos + 1, MyPos2 - (MyPos + 1))
110       ElseIf ActiveLineType = 1 Then 'additional effect: defender Takes
120           Attacker = PreviousAttacker
130           MyPos = InStr(21, CurrentLine, " takes ")
140           If MyPos <> 0 Then
150               Defender = Mid(CurrentLine, 22, MyPos - 22)
160           Else
170               Defender = CurrentFight
180           End If
190       ElseIf ActiveLineType = 2 Then 'attacker's ranged attack Hits defender For
200           MyPos = InStr(1, CurrentLine, " ranged ")
210           If MyPos = 0 Then Exit Sub
220           Attacker = Mid(CurrentLine, 3, MyPos - 5)
  
230           MyPos = InStr(1, CurrentLine, " hits ")
240           If MyPos = 0 Then
250               MyPos = InStr(1, CurrentLine, " on ")
260               MyPos2 = InStr(1, CurrentLine, " for ")
270               If MyPos2 = 0 Then MyPos2 = InStr(1, CurrentLine, ".")
280               Defender = Mid(CurrentLine, MyPos + 4, MyPos2 - (MyPos + 4))
290           Else
300               MyPos2 = InStr(1, CurrentLine, " for ")
310               If MyPos2 = 0 Then MyPos2 = InStr(1, CurrentLine, ".")
320               Defender = Mid(CurrentLine, MyPos + 6, MyPos2 - (MyPos + 6))
330           End If
340       ElseIf ActiveLineType = 3 Then 'attacker's attack is countered by defender, attacker takes
350           MyPos = InStr(1, CurrentLine, " by ")
360           MyPos2 = InStr(1, CurrentLine, ". ")
370           Attacker = Mid(CurrentLine, MyPos + 4, MyPos2 - (MyPos + 4))
      '290           MyPos = InStr(1, CurrentLine, ". ")
380           MyPos2 = InStr(1, CurrentLine, "'s")
390           Defender = Mid(CurrentLine, 3, MyPos2 - 3)
400       End If
  
410       If InStr(1, "14,19", LineType) Then
420           CurrentFight = Defender
430       ElseIf InStr(1, "1c,20", LineType) Then
440           CurrentFight = Attacker
450       End If
460       CurrentFight = Replace(Replace(CurrentFight, "the ", "The "), "Cover!", "")
470       If InStr(1, CurrentFight, "'s") Then CurrentFight = Trim(Replace(CurrentFight, "'s", ""))
480   ElseIf ActiveLineType <= 11 Then
490       MyPos = InStr(1, CurrentLine, " misses ")
500       Select Case ActiveLineType
          Case 4
510           MyPos = InStr(1, CurrentLine, " miss ")
520       Case 5
530           MyPos = InStr(1, CurrentLine, " misses ")
540       Case 6
550           MyPos = InStr(1, CurrentLine, " parries ")
560       Case 7
570           MyPos = InStr(1, CurrentLine, " blocks ")
580       Case 8
590           MyPos = InStr(1, CurrentLine, " shadows absorbs ")
600       Case 9
610           MyPos = InStr(1, CurrentLine, " anticipates ")
620       Case 10
630           MyPos = InStr(1, CurrentLine, " ranged attack misses.")
640       Case 11
650           MyPos = InStr(1, CurrentLine, " evades.")
660       End Select
    
670       If InStr(1, CurrentLine, " uses ") Then
680           MyPos = InStr(1, CurrentLine, " uses ")
690       End If
700       If MyPos = 0 Then Exit Sub
710       Attacker = Mid$(CurrentLine, 3, MyPos - 3)
720     If ActiveLineType = 8 Then
730         Attacker = Mid(Attacker, 6)
740     End If
750       If InStr(1, CurrentLine, " uses ") Then AttackerUses = Attacker
    
760       Select Case ActiveLineType
          Case 4
770           MyPos = MyPos + 6
780       Case 5
790           MyPos = MyPos + 8
800       Case 6
810           MyPos = MyPos + 9
820       Case 7
830           MyPos = MyPos + 8
840       Case 8
850           MyPos = MyPos + 16
860       Case 9
870           MyPos = MyPos + 13
880       Case 10
890           MyPos = MyPos + 22
900       Case 11
910           MyPos = MyPos + 9
920       End Select
    
930       Attacker = Replace(Attacker, "'s", "")
940       If InStr(1, CurrentLine, " misses.") = 0 And InStr(1, CurrentLine, ", but") = 0 And InStr(1, CurrentLine, " parries ") = 0 And InStr(1, CurrentLine, " blocks ") = 0 And InStr(1, CurrentLine, " absorbs the damage and ") = 0 And InStr(1, CurrentLine, " evades.") = 0 And InStr(1, CurrentLine, " anticipates the attack.") = 0 Then
950           MyPos2 = InStr(1, CurrentLine, ".")
960           Defender = Mid$(CurrentLine, MyPos, MyPos2 - MyPos)
970       ElseIf InStr(1, CurrentLine, " misses.") = 0 And InStr(1, CurrentLine, " parries ") = 0 And InStr(1, CurrentLine, " blocks ") = 0 And InStr(1, CurrentLine, " absorbs the damage and ") = 0 And InStr(1, CurrentLine, " evades.") = 0 And InStr(1, CurrentLine, " anticipates the attack.") = 0 Then
980           MyPos = InStr(1, CurrentLine, ", but misses ")
990           MyPos2 = InStr(1, CurrentLine, ".")
1000          Defender = Replace(Mid$(CurrentLine, MyPos + 13, MyPos2 - (MyPos + 13)), "the ", "The ")
1010      ElseIf InStr(1, CurrentLine, " parries ") Or InStr(1, CurrentLine, " blocks ") Then
1020          MyPos2 = InStr(1, CurrentLine, "attack ")
1030          Defender = Mid$(CurrentLine, MyPos, MyPos2 - MyPos)
1040      ElseIf Mid$(CurrentLine, 3, 4) <> "The " Then
1050          Defender = CurrentFight
1060      Else
1070          Defender = ""
1080      End If
1090      If Defender = "" Then Defender = CurrentFight
1100      If InStr(1, Attacker, "'s") Then Attacker = Trim(Replace(Attacker, "'s", ""))
1110      If InStr(1, Defender, "'s") Then Defender = Trim(Replace(Defender, "'s", ""))
1120      If InStr(1, CurrentLine, " absorbs the damage and ") Or InStr(1, CurrentLine, " evades.") Or InStr(1, CurrentLine, " blocks ") Or InStr(1, CurrentLine, " parries ") Or InStr(1, CurrentLine, " anticipates the attack.") Then
1130          PreP1 = Attacker
1140          Attacker = Defender
1150          Defender = PreP1
1160      End If
1170  ElseIf ActiveLineType = 12 Then
1180      MyPos = InStr(3, CurrentLine, " uses ")
1190      If MyPos = 0 Then
1200          MyPos = InStr(3, CurrentLine, "'s use ")
1210      End If
1220      If MyPos = 0 Then
1230          MyPos = InStr(3, CurrentLine, "s use ")
1240      End If
1250      AttackerUses = Mid$(CurrentLine, 3, MyPos - 3)
1260      MyPos2 = InStr(1, CurrentLine, ".")
1270      If MyPos2 = 0 Then
1280          MyPos2 = InStr(1, CurrentLine, "!")
1290      End If
1300      If MyPos2 = 0 Then
1310          MyPos2 = InStrRev(CurrentLine, " ")
1320      End If
1330      MyPos = InStr(3, CurrentLine, " uses ")
1340      If MyPos <> 0 Then
1350          AttackerSpecial = Mid(CurrentLine, MyPos + 6, MyPos2 - (MyPos + 6))
1360      Else
1370          AttackerSpecial = Mid(CurrentLine, MyPos + 7, MyPos2 - (MyPos + 7))
1380      End If
1390  ElseIf ActiveLineType = 14 Then
1400      MyPos = InStr(3, CurrentLine, ".")
1410      AttackerUses = "SC: " & Mid$(CurrentLine, 15, MyPos - 15)
1420  ElseIf ActiveLineType = 13 Then
1430      If InStr(1, CurrentLine, "ranged") = 0 Then
1440          MyPos = InStr(3, CurrentLine, " score")
1450          AttackerUses = Mid$(CurrentLine, 3, MyPos - 3)
1460      Else
1470          MyPos = InStr(3, CurrentLine, "'s")
1480          AttackerUses = Mid$(CurrentLine, 3, MyPos - 3)
1490      End If
1500  ElseIf ActiveLineType = 15 Then
1510      MyPos = InStr(3, CurrentLine, " casts ")
1520      AttackerUses = Mid$(CurrentLine, 3, MyPos - 3)
1530      MyPos2 = InStr(1, CurrentLine, ".")
1540      If MyPos2 = 0 Then
1550          MyPos2 = InStrRev(CurrentLine, " ")
1560      End If
1570      AttackerSpecial = Mid(CurrentLine, MyPos + 7, MyPos2 - (MyPos + 7))
1580  ElseIf ActiveLineType = 16 Then
1590      Attacker = AttackerUses
1600      MyPos = InStr(3, CurrentLine, " take")
1610      Defender = Mid$(CurrentLine, 3, MyPos - 3)
1620  ElseIf ActiveLineType = 17 Then
1630      Attacker = AttackerUses
1640      MyPos = InStr(3, CurrentLine, " from ")
1650      Defender = Mid$(CurrentLine, MyPos + 6, InStr(1, CurrentLine, ".") - (MyPos + 6))
1660  ElseIf ActiveLineType = 18 Then
1670      MyPos = InStr(3, CurrentLine, " uses ")
1680      If MyPos = 0 Then
1690        MyPos = InStr(3, CurrentLine, " use ")
1700      End If
1710      Attacker = Mid$(CurrentLine, 3, MyPos - 3)
1720      MyPos = InStr(1, CurrentLine, " misses ")
1730      If MyPos = 0 Then
1740          MyPos = InStr(3, CurrentLine, " miss ")
1750          MyPos2 = InStr(1, CurrentLine, ".")
1760          If MyPos2 = 0 Then
1770              MyPos2 = InStrRev(CurrentLine, " ")
1780          End If
1790          Defender = Mid(CurrentLine, MyPos + 6, MyPos2 - (MyPos + 6))
1800      Else
1810          MyPos2 = InStr(1, CurrentLine, ".")
1820          If MyPos2 = 0 Then
1830              MyPos2 = InStrRev(CurrentLine, " ")
1840          End If
1850          Defender = Mid(CurrentLine, MyPos + 8, MyPos2 - (MyPos + 8))
1860      End If
1870      MyPos = InStr(3, CurrentLine, " uses ")
1880      MyPos2 = InStr(3, CurrentLine, ", ")
1890      AttackerSpecial = Mid(CurrentLine, MyPos + 6, MyPos2 - (MyPos + 6))
1900  ElseIf ActiveLineType = 19 Then
1910      Attacker = AttackerUses
1920      Defender = CurrentFight
1930  ElseIf ActiveLineType = 90 Then
1940      If InStr(1, LCase(CurrentLine), "defeats") Then
1950          MyPos = InStr(1, CurrentLine, "defeats ")
1960          MyPos2 = InStr(1, CurrentLine, ".")
1970          Defender = Mid$(CurrentLine, MyPos + 8, MyPos2 - (MyPos + 8))
1980      ElseIf InStr(1, LCase(CurrentLine), "defeated by") Then
1990          MyPos = InStr(1, CurrentLine, "defeated by")
2000          MyPos2 = InStr(1, CurrentLine, ".")
2010          Defender = Mid$(CurrentLine, MyPos + 12, MyPos2 - (MyPos + 12))
2020      Else
2030          MyPos = InStr(1, CurrentLine, "fall")
2040          Defender = Mid$(CurrentLine, 3, MyPos - 4)
2050      End If
2060  End If
    
      'Touch up names
2070  Defender = Replace(Replace(Defender, "the ", "The "), "Cover!", "")
2080  Attacker = Replace(Replace(Attacker, "the ", "The "), "Cover!", "")

      '      if instr(1, defender," of "
2090  Defender = Replace(Defender, "Magic Burst! ", "")
2100  If InStr(1, Defender, "'s") Then Defender = Trim(Replace(Defender, "'s", ""))
2110  If InStr(1, Attacker, "'s") Then Attacker = Trim(Replace(Attacker, "'s", ""))
2120  Attacker = Trim(Attacker)
2130  Defender = Trim(Defender)

2140  If Attacker <> "" And LineType <> "1c" And LineType <> "20" And LineType <> "1d" And LineType <> "28" And LineType <> "29" Then 'Used for Additional Effects.. Excluding enemy hits.. Bleh this is gay
2150      PreviousAttacker = Attacker
2160  End If

      'Add players to player combobox and listbox
      Dim FoundPlayer As Boolean, i As Integer
2170  If (LineType = "14" Or LineType = "19" Or LineType = "28") And Attacker <> "" And InStr(1, Attacker, " ") = 0 Then
2180      For i = 0 To listPlayers.ListCount - 1
2190          If listPlayers.List(i) = Attacker Then
2200              FoundPlayer = True
2210              Exit For
2220          End If
2230      Next
2240      If FoundPlayer = False Then
2250          listPlayers.AddItem Attacker
2260          comboUser.AddItem Attacker
2270          For i = 0 To listPlayers.ListCount - 1
2280              listPlayers.Selected(i) = True
2290          Next
2300      End If
2310  End If
      'Add SkillChains to user report.. Interesting data =p
2320  FoundPlayer = False
2330  If InStr(1, Attacker, "SC:") Then
2340      For i = 0 To comboUser.ListCount - 1
2350          If comboUser.List(i) = Attacker Then
2360              FoundPlayer = True
2370              Exit For
2380          End If
2390      Next
2400      If FoundPlayer = False Then
2410          comboUser.AddItem Attacker
2420      End If
2430  End If
2440  Exit Sub
Err_Handler:
2450  HasErrors = True
2460  ErrorCount = ErrorCount + 1
2470  ReportError = "Error: " & Err.Number & vbNewLine & "Source: RetrieveUsers" & vbNewLine & "Description: " & Err.Description & vbNewLine & "Line: " & Erl & vbNewLine & "FFXI Log Line: " & CurrentLine & vbNewLine & "Previous FFXI Log Line: " & PrevLineA & vbNewLine & vbNewLine
2480  Print #ErrorFile, ReportError
2490  If ErrorCount >= 25 Then
2500      lblStatus.Caption = "Errors: " & HasErrors & " - Too many errors - Parsing stopped for this log."
2510      Exit Sub
2520  Else
2530      Resume Next
2540  End If
End Sub

Private Sub SendData()
inet.OpenURL ""
End Sub

Private Sub SetActiveLineType()
10    On Error GoTo Err_Handler
      'ActiveLineTypes:
      '0 = Basic
      '1 = Additional Effect
      '2 = Ranged
      '3 = Counter
      '4 = Miss
      '5 = Misses
      '6 = Parry
      '7 = Block
      '8 = Absorb
      '9 = Anticipate
      '10 = Ranged miss
      '11 = Evade

      '12 = Ability/Skill Use
      '13 = Crit
      '14 = SkillChain
      '15 = Spell
      '16 = Ability/Skill Hit
      '17 = HP Drain
      '18 = Ability/Skill Miss
      '19 = Heals

      '20 = Fishing
      '30 = Loot Found
      '31 = Loot by Player
      '32 = Gil Obtained
      '40 = Exp Lines
      '50 = /Clock
      '51 = Crafting
      '52 = /Clock part 2!
      '53 = /Clock part 3!
      '54 = Crafting Part 2!
      '60 = Monster Check
      '70 = Chat Text
      '80 = Parser Command
      '90 = Enemy Defeated
      '99 = Other
20    PrevActiveLineType = ActiveLineType
30    If InStr(1, "09,0a,01,02,0c,04,0d,05,0e,06,0f,07,9d,98,a1,79,83,6a,94,8c,7f,bf,cd,cc,7a,92", LineType) = 0 Then 'Ignore lines that we don't need
40        If InStr(1, CurrentLine, "Additional effect: ") <> 0 And InStr(1, LCase(CurrentLine), " damage") <> 0 Then
50            ActiveLineType = 1
60        ElseIf (InStr(1, CurrentLine, " uses") Or InStr(1, LCase(CurrentLine), "s use ")) And InStr(1, CurrentLine, ", but miss") = 0 Then
70            ActiveLineType = 12
80        ElseIf (InStr(1, CurrentLine, " uses") Or InStr(1, LCase(CurrentLine), "s use ")) And InStr(1, CurrentLine, ", but miss") <> 0 Then
90            ActiveLineType = 18
100       ElseIf InStr(1, CurrentLine, " take") Then
110           ActiveLineType = 16
120       ElseIf InStr(1, CurrentLine, "critical hit!") Then
130           ActiveLineType = 13
140       ElseIf InStr(1, LCase(CurrentLine), " counter") <> 0 And InStr(1, LCase(CurrentLine), " counterstance") = 0 Then
150           ActiveLineType = 3
160       ElseIf InStr(1, CurrentLine, " parries ") Then
170           ActiveLineType = 6
180       ElseIf InStr(1, CurrentLine, " blocks ") Then
190           ActiveLineType = 7
200       ElseIf InStr(1, CurrentLine, " absorbs the damage and ") Then
210           ActiveLineType = 8
220       ElseIf InStr(1, CurrentLine, " anticipates ") Then
230           ActiveLineType = 9
240       ElseIf InStr(1, CurrentLine, " ranged attack misses") Or InStr(1, CurrentLine, " Ranged Attack, but misses") Then
250           ActiveLineType = 10
260       ElseIf InStr(1, LCase(CurrentLine), "ranged attack") And LineType <> "65" Then
270           ActiveLineType = 2
280       ElseIf InStr(1, CurrentLine, " miss ") Then
290           ActiveLineType = 4
300       ElseIf InStr(1, CurrentLine, " misses ") Then
310           ActiveLineType = 5
320       ElseIf InStr(1, CurrentLine, " evades") Then
330           ActiveLineType = 11
340       ElseIf InStr(1, CurrentLine, " hit ") Or InStr(1, CurrentLine, " hits ") Then
350           ActiveLineType = 0
360       ElseIf InStr(1, LCase(CurrentLine), "skillchain: ") Then
370           ActiveLineType = 14
380       ElseIf InStr(1, CurrentLine, " casts ") Then
390           ActiveLineType = 15
400       ElseIf InStr(1, CurrentLine, " HP drained from") Then
410           ActiveLineType = 17
420       ElseIf CurrentFight <> "" And InStr(1, LCase(CurrentLine), " recovers hp") = 0 And InStr(1, LCase(CurrentLine), " recovers ") <> 0 And InStr(1, CurrentLine, " MP.") = 0 Then
430           ActiveLineType = 19
440       ElseIf LineType = "ce" And InStr(1, CurrentLine, "parser") Then
450           ActiveLineType = 80
460       ElseIf InStr(1, "24,25,a6", LineType) Then
470           ActiveLineType = 90
480       Else
490           ActiveLineType = 99
500       End If
510   ElseIf LineType = "79" And (InStr(1, LCase(CurrentLine), "you obtained") Or InStr(1, LCase(CurrentLine), "was lost")) Then
520       ActiveLineType = 54
530   ElseIf LineType = "79" And (InStr(1, LCase(CurrentLine), "you synthesized") Or InStr(1, LCase(CurrentLine), "synthesis failed")) Then
540       ActiveLineType = 51
550   ElseIf (InStr(1, LCase(CurrentLine), "obtained: ") <> 0 Or InStr(1, LCase(CurrentLine), "you lost your catch.") <> 0 Or InStr(1, LCase(CurrentLine), "you didn't catch anything.") <> 0) And Mid$(CurrentLine, 3, 1) <> "<" And Mid$(CurrentLine, 3, 1) <> ">" And Mid$(CurrentLine, 3, 1) <> "(" And InStr(1, CurrentLine, " : ") = 0 And LineType <> "0f" Then
560       ActiveLineType = 20
570   ElseIf LineType = "7f" And InStr(1, CurrentLine, " obtains ") And InStr(1, CurrentLine, " gil.") = 0 Then
580       ActiveLineType = 31
590   ElseIf InStr(1, LCase(CurrentLine), " obtains ") <> 0 And InStr(1, LCase(CurrentLine), " gil.") <> 0 And Mid$(CurrentLine, 3, 1) <> "<" And Mid$(CurrentLine, 3, 1) <> ">" And Mid$(CurrentLine, 3, 1) <> "(" And InStr(1, CurrentLine, " : ") = 0 And LineType <> "0f" Then
600       ActiveLineType = 32
610   ElseIf (Left$(CurrentLine, 12) = "yYou find") Then
620       ActiveLineType = 30
630   ElseIf (LineType = "79" Or LineType = "83") And (InStr(1, LCase(CurrentLine), "exp") Or InStr(1, LCase(CurrentLine), "limit")) Then
640       ActiveLineType = 40
650   ElseIf InStr(1, CurrentLine, "Earth:") And LineType = "8c" Then
660       ActiveLineType = 50
670   ElseIf InStr(1, CurrentLine, "Vana'diel:") And LineType = "8c" Then
680       ActiveLineType = 52
690   ElseIf LineType = "8c" Then
700       ActiveLineType = 53
710   ElseIf LineType = "bf" Then
720       ActiveLineType = 60
730   ElseIf InStr(1, "09,0a,01,02,0c,04,0d,05,0e,06,0f", LineType) Then
740       ActiveLineType = 70
750   Else
760       ActiveLineType = 99
770   End If


780   Exit Sub
Err_Handler:
790   HasErrors = True
800   ErrorCount = ErrorCount + 1
810   ReportError = "Error: " & Err.Number & vbNewLine & "Source: SetActiveLineType" & vbNewLine & "Description: " & Err.Description & vbNewLine & "Line: " & Erl & vbNewLine & "FFXI Log Line: " & CurrentLine & vbNewLine & "Previous FFXI Log Line: " & PrevLineA & vbNewLine & vbNewLine
820   Print #ErrorFile, ReportError
830   If ErrorCount >= 25 Then
840       lblStatus.Caption = "Errors: " & HasErrors & " - Too many errors - Parsing stopped for this log."
850       Exit Sub
860   Else
870       Resume Next
880   End If
End Sub

Private Sub SetTimeA()
10    On Error GoTo Err_Handler
      Dim MyPos, MyPos2
20    MyPos = InStr(1, CurrentLine, ",")
30    MyPos2 = InStr(MyPos + 1, CurrentLine, ",")
40    CurrentDay = Mid(CurrentLine, MyPos + 2, MyPos2 - (MyPos + 2))
50    MyPos = MyPos2
60    MyPos2 = InStr(MyPos + 2, CurrentLine, " ")
70    CurrentTime = Mid(CurrentLine, MyPos + 2, MyPos2 - (MyPos + 2))
80    Exit Sub
Err_Handler:
90    HasErrors = True
100   ErrorCount = ErrorCount + 1
110   ReportError = "Error: " & Err.Number & vbNewLine & "Source: SetTimeA" & vbNewLine & "Description: " & Err.Description & vbNewLine & "Line: " & Erl & vbNewLine & "FFXI Log Line: " & CurrentLine & vbNewLine & "Previous FFXI Log Line: " & PrevLineA & vbNewLine & vbNewLine
120   Print #ErrorFile, ReportError
130   If ErrorCount >= 25 Then
140       lblStatus.Caption = "Errors: " & HasErrors & " - Too many errors - Parsing stopped for this log."
150       Exit Sub
160   Else
170       Resume Next
180   End If
End Sub
Private Sub SetTimeB()
10    On Error GoTo Err_Handler
      Dim MyPos, MyPos2
20    MyPos = InStr(1, CurrentLine, "(")
30    MyPos2 = InStr(MyPos + 2, CurrentLine, ")")
40    CurrentMoon = Trim(Mid(CurrentLine, 3, MyPos - 4))
50    CurrentPerc = Mid(CurrentLine, MyPos + 1, MyPos2 - (MyPos + 1))
60    Exit Sub
Err_Handler:
70    HasErrors = True
80    ErrorCount = ErrorCount + 1
90    ReportError = "Error: " & Err.Number & vbNewLine & "Source: SetTimeB" & vbNewLine & "Description: " & Err.Description & vbNewLine & "Line: " & Erl & vbNewLine & "FFXI Log Line: " & CurrentLine & vbNewLine & "Previous FFXI Log Line: " & PrevLineA & vbNewLine & vbNewLine
100   Print #ErrorFile, ReportError
110   If ErrorCount >= 25 Then
120       lblStatus.Caption = "Errors: " & HasErrors & " - Too many errors - Parsing stopped for this log."
130       Exit Sub
140   Else
150       Resume Next
160   End If
End Sub


Public Sub StartNew()
      Dim i As Integer, MyPos As Integer, f
10    On Error GoTo Err_Handler
20    ClearEdit = True
30    mnuClear_Click
40    ClearEdit = False

50    RTB_Log.Text = Date & " - " & Time & vbNewLine & vbNewLine & RTB_Log.Text
60    RTB_Log.Text = "Log Location: " & GetSetting(App.Title, "Settings", "LogPath", "C:\Program Files\PlayOnline\SquareEnix\FINAL FANTASY XI\TEMP") & vbNewLine & RTB_Log.Text
70    RTB_Log.Text = "Read New Logs Only: " & mnuOnly.Checked & vbNewLine & RTB_Log.Text
80    RTB_Log.Text = "Parser Commands: " & mnuParserCommands.Checked & vbNewLine & RTB_Log.Text
90    RTB_Log.Text = "Key Activated Timers: " & mnuKeyEnable.Checked & vbNewLine & RTB_Log.Text

100   If fileList.ListCount <> 0 Then
          Dim FSO As FileSystemObject
          Dim fo As Integer
          Dim MyDate As Date
          Dim FullFile() As String, CurrentFile As String
          Dim Index As Long

110       Set FSO = New FileSystemObject
120       If FSO.FolderExists(App.Path & "\FFXI_Logs") = False Then
130           FSO.CreateFolder App.Path & "\FFXI_Logs"
140           RTB_Log.Text = "Creating folder: " & App.Path & "\FFXI_Logs" & vbNewLine & RTB_Log.Text
150       End If
160       If FSO.FolderExists(App.Path & "\FFXI_Gather") = False Then
170           FSO.CreateFolder App.Path & "\FFXI_Gather"
180           RTB_Log.Text = "Creating folder: " & App.Path & "\FFXI_Gather" & vbNewLine & RTB_Log.Text
190       End If
200       If FSO.FileExists(App.Path & "\EditFile.log") = True Then
210           FSO.DeleteFile (App.Path & "\EditFile.log")
220           RTB_Log.Text = "Deleting File: " & App.Path & "\EditFile.log" & vbNewLine & RTB_Log.Text
230       End If
240       lblStatus.Caption = "Errors: " & HasErrors & " - " & "Parsing Data..."
250       DoEvents
260       fileListBox.Clear

270       If Gather = True Or ParseGather = True Then
280           If GatherDate Then
290             CreateSingleFile
300           End If
310           If FSO.FileExists(SingleFile) = True Then
320               FSO.DeleteFile (SingleFile)
330               RTB_Log.Text = "Deleting File: " & SingleFile & vbNewLine & RTB_Log.Text
340           End If
              Dim EditFile
350           EditFile = FreeFile
360           Open SingleFile For Append As #EditFile
370       End If
380       If OpenSingle = False And Gather = False And ParseGather = False Then
              Dim TmpFile
390           TmpFile = FreeFile
400           Open App.Path & "\ffxip_tmp_.tmp" For Output As #TmpFile
410       End If
    
420       If OpenSingle = False Then
430           For i = 0 To fileList.ListCount - 1
440               fileList.ListIndex = i
450               fileListBox.AddItem Format(FileDateTime(dirList.Path & "\" & fileList.FileName), "MM/DD HhNnSs") & " - " & fileList.Path & "\" & fileList.FileName
460           Next
470       End If
    
480       If OpenSingle Then
490           Erase FullFile
500           Index = 0
510           f = FreeFile
520           Open SingleFile For Input As f
530             Do Until EOF(f)
540               Line Input #f, CurrentLine
550               ReDim Preserve FullFile(Index)
560               FullFile(Index) = CurrentLine
570               Index = Index + 1
580             Loop
590           Close #f
600           If Index <> 0 Then
610             ParseLog FullFile
620           End If
630           mnuRecalculate.Enabled = True
640           mnuExport.Enabled = True
650           mnuExportXML.Enabled = True
660           If optionResults(1).Value = True Then
670               comboUser_Click
680           Else
690               comboDisplay_Click
700           End If
710           lblStatus.Caption = "Errors: " & HasErrors & " - " & "Finished Parsing Data."
720           Exit Sub
730       Else
740           Erase FullFile
750           Index = 0
760           If Me.mnuOnly.Checked = False Then
770               For fo = 0 To fileListBox.ListCount - 1
780                 fileListBox.ListIndex = fo
790                 f = FreeFile
  
800                 RTB.LoadFile Mid$(fileListBox.Text, 16)
810                 RTB_Log.Text = "Loading File: " & Mid$(fileListBox.Text, 16) & vbNewLine & RTB_Log.Text
820                 RTB.Text = Mid(RTB.Text, 101)
830                 RTB.Text = Replace(RTB.Text, Chr(0), vbNewLine)
840                 MyPos = InStrRev(fileListBox.Text, "\")
850                 If Gather = False Then
860                   CurrentFile = App.Path & "\FFXI_Logs" & Mid(fileListBox.Text, MyPos)
870                 Else
880                   CurrentFile = App.Path & "\FFXI_Gather" & Mid(fileListBox.Text, MyPos)
890                 End If
900                 RTB.SaveFile CurrentFile, rtfText
910                 RTB_Log.Text = "Saving File: " & CurrentFile & vbNewLine & RTB_Log.Text
  
920                 MyDate = Left$(fileListBox.Text, 5) & Format(Date, "/YYYY") & " " & Format(Format(Mid$(fileListBox.Text, 7, 6), "00:00:00"), "Hh:Nn:Ss AM/PM")
930                 ResetTimeFile CurrentFile, MyDate
940                 Open CurrentFile For Input As f
950                   Do Until EOF(f)
960                       Line Input #f, CurrentLine
970                       LineType = Left(CurrentLine, 2)
980                       If LineType = "ce" And ParserCommands = True Then
990                           If InStr(1, LCase(CurrentLine), "parser stop logging") Then
1000                              StopLogging = True
1010                          ElseIf InStr(1, LCase(CurrentLine), "parser start logging") Then
1020                              StopLogging = False
1030                          End If
1040                      End If
1050                      If StopLogging = False Then
1060                          If Mid(CurrentLine, 51, 2) = "01" And Index <> 0 Then
1070                              FullFile(Index - 1) = Left(FullFile(Index - 1), Len(FullFile(Index - 1)) - 3) & Mid(CurrentLine, 56) & " " & LineType
1080                          Else
1090                              ReDim Preserve FullFile(Index)
1100                              FullFile(Index) = Mid(CurrentLine, 54) & " " & LineType
1110                              Index = Index + 1
1120                          End If
1130                      End If
1140                  Loop
1150                Close #f
1160              Next
1170          End If
1180          If Gather = False Or ParseGather = True Then
1190            If Index <> 0 Then
1200              ParseLog FullFile
1210            End If
1220            mnuRecalculate.Enabled = True
1230            mnuExport.Enabled = True
1240            mnuExportXML.Enabled = True
1250            If ParseGather Then
1260              If Index <> 0 Then
1270                  For fo = 0 To UBound(FullFile)
1280                      Print #EditFile, FullFile(fo)
1290                  Next
1300              End If
1310            End If
1320          Else
1330              If Index <> 0 Then
1340                  For fo = 0 To UBound(FullFile)
1350                      Print #EditFile, FullFile(fo)
1360                  Next
1370              End If
1380          End If
1390          If Gather = False And ParseGather = False Then
1400              If Index <> 0 Then
1410                  For fo = 0 To UBound(FullFile)
1420                      Print #TmpFile, FullFile(fo)
1430                  Next
1440              End If
1450              Close #TmpFile
1460          End If
1470      End If
    
1480      fileListBox.ListIndex = fileListBox.ListCount - 1
1490      LastItem = fileListBox.Text
1500      timerRead.Enabled = True
1510      mnuParse.Caption = "&Stop Parsing"
1520      mnuGather.Enabled = True
1530      mnuGatherDate.Enabled = True
1540      mnuOpen.Enabled = True
1550      If optionResults(1).Value = True Then
1560         comboUser_Click
1570      Else
1580         comboDisplay_Click
1590      End If
1600      lblStatus.Caption = "Errors: " & HasErrors & " - " & "Waiting for new log...."
1610      If Gather = True Or ParseGather = True Then
1620          Close #EditFile
1630      End If
1640  Else
1650      MsgBox "No log files found in this folder. Please select another folder." & vbNewLine & vbNewLine & "Usually: ""C:\Program Files\PlayOnline\SquareEnix\FINAL FANTASY XI\TEMP""", vbInformation, "Error"
1660      mnuLocation_Click
1670      mnuParse.Caption = "&Start Parsing"
1680      lblStatus.Caption = "Errors: " & HasErrors & " - " & "Waiting."
1690      timerRead.Enabled = False
1700      mnuGather.Enabled = True
1710      mnuGatherDate.Enabled = True
1720      mnuOpen.Enabled = True
1730  End If

1740  Set FSO = Nothing


1750  Exit Sub
Err_Handler:
1760  If Err.Number = 52 And ErrorCount < 25 Then
1770      Err.Clear
1780      ErrorCount = ErrorCount + 1
1790      f = FreeFile
1800      EditFile = FreeFile
1810      TmpFile = FreeFile
1820      Resume
1830  End If
1840  HasErrors = True
1850  ErrorCount = ErrorCount + 1
1860  ReportError = "Error: " & Err.Number & vbNewLine & "Source: StartNew" & vbNewLine & "Description: " & Err.Description & vbNewLine & "Line: " & Erl & vbNewLine & "FFXI Log Line: " & CurrentLine & vbNewLine & "Previous FFXI Log Line: " & PrevLineA & vbNewLine & vbNewLine
1870  Print #ErrorFile, ReportError
1880  If ErrorCount >= 25 Then
1890      lblStatus.Caption = "Errors: " & HasErrors & " - Too many errors - Parsing stopped for this log."
1900      Exit Sub
1910  Else
1920      Resume Next
1930  End If
End Sub

Private Sub ExportHTML(SummaryOnly As Boolean, FileName As String)
10    On Error GoTo Err_Handler
20    Screen.MousePointer = vbHourglass
      Dim HTMLFile, TotalFights As Long, HTMLCode As String
      Dim HTMLSummaryStats() As udtStatistics, FoundSummary As Boolean, IncludePlayer As Boolean
      Dim SummaryTotals As udtStatistics, EmptyStats As udtStatistics
      Dim i As Integer, p As Integer, c As Integer
30    ReDim HTMLSummaryStats(0)


40    HTMLFile = FreeFile
50    Open FileName For Output As HTMLFile
    
60    For i = 0 To listResults.ListCount - 1
70        If listResults.Selected(i) Then
80            TotalFights = TotalFights + 1
90        End If
100   Next


110   For i = 0 To UBound(FullStats) - 1
120       IncludePlayer = False
130       For c = 0 To listPlayers.ListCount - 1
140           If listPlayers.List(c) = FullStats(i).Attacker Or InStr(1, FullStats(i).Attacker, "SC:") <> 0 Then
150               If listPlayers.Selected(c) = True Then
160                   IncludePlayer = True
170               End If
180               Exit For
190           End If
200       Next
210       If listResults.Selected(FullStats(i).BattleID) = True And IncludePlayer Then
220           FoundSummary = False
230           For p = 0 To UBound(HTMLSummaryStats)
240               If FullStats(i).Attacker = HTMLSummaryStats(p).Attacker Then
250                   CombineStats HTMLSummaryStats(p), FullStats(i)
260                   FoundSummary = True
270               End If
280           Next
290           If FoundSummary = False Then
300               CombineStats HTMLSummaryStats(UBound(HTMLSummaryStats)), FullStats(i)
310               ReDim Preserve HTMLSummaryStats(UBound(HTMLSummaryStats) + 1)
320           End If
330       End If
340   Next

350   For i = 0 To UBound(HTMLSummaryStats) - 1
360       CombineStats SummaryTotals, HTMLSummaryStats(i)
370   Next

380   SummaryTotals.Battles = TotalFights
390   HTMLCode = HTMLCode & "<style type=""text/css"">"
400   HTMLCode = HTMLCode & "TD {BORDER-RIGHT: #7CB1CB 1px solid; BORDER-TOP: #7CB1CB 1px solid; BORDER-LEFT: #7CB1CB 1px solid; BORDER-BOTTOM: #7CB1CB 1px solid}"
410   HTMLCode = HTMLCode & "</style>"
420   HTMLCode = GenerateCode(HTMLSummaryStats(), SummaryTotals, True)

      Dim f As Long
430   ReDim HTMLSummaryStats(0)
440   If SummaryOnly = False Then
450       For i = 0 To listResults.ListCount - 1
460           If listResults.Selected(i) = True Then
470               For f = 0 To UBound(FullStats) - 1
480                   IncludePlayer = False
490                   For c = 0 To listPlayers.ListCount - 1
500                       If listPlayers.List(c) = FullStats(f).Attacker Or InStr(1, FullStats(f).Attacker, "SC:") <> 0 Then
510                           If listPlayers.Selected(c) = True Then
520                               IncludePlayer = True
530                           End If
540                           Exit For
550                       End If
560                   Next
570                   If FullStats(f).BattleID = i And IncludePlayer Then
580                       CombineStats HTMLSummaryStats(UBound(HTMLSummaryStats)), FullStats(f)
590                       ReDim Preserve HTMLSummaryStats(UBound(HTMLSummaryStats) + 1)
600                   End If
610               Next
620               SummaryTotals = EmptyStats
630               For f = 0 To UBound(HTMLSummaryStats) - 1
640                   CombineStats SummaryTotals, HTMLSummaryStats(f)
650               Next
660               HTMLCode = HTMLCode & GenerateCode(HTMLSummaryStats(), SummaryTotals, False)
670               ReDim HTMLSummaryStats(0)
680           End If
690       Next
700   End If

710   Print #HTMLFile, HTMLCode
720   Close #HTMLFile
730   Screen.MousePointer = vbDefault

740   Exit Sub
Err_Handler:
750   HasErrors = True
760   ErrorCount = ErrorCount + 1
770   ReportError = "Error: " & Err.Number & vbNewLine & "Source: ExportHTML" & vbNewLine & "Description: " & Err.Description & vbNewLine & "Line: " & Erl & vbNewLine & "FFXI Log Line: " & CurrentLine & vbNewLine & "Previous FFXI Log Line: " & PrevLineA & vbNewLine & vbNewLine
780   Print #ErrorFile, ReportError
790   If ErrorCount >= 25 Then
800       lblStatus.Caption = "Errors: " & HasErrors & " - Too many errors - Parsing stopped for this log."
810       Exit Sub
820   Else
830       Resume Next
840   End If
End Sub







Private Sub displaySummary()
Screen.MousePointer = vbHourglass
Dim TotalFights As Long
Dim HTMLSummaryStats() As udtStatistics, FoundSummary As Boolean, IncludePlayer As Boolean
Dim SummaryTotals As udtStatistics, EmptyStats As udtStatistics
Dim i As Integer, p As Integer, c As Integer
ReDim HTMLSummaryStats(0)

    
For i = 0 To listResults.ListCount - 1
    If listResults.Selected(i) Then
        TotalFights = TotalFights + 1
    End If
Next

For i = 0 To UBound(FullStats) - 1
    IncludePlayer = False
    For c = 0 To listPlayers.ListCount - 1
        If listPlayers.List(c) = FullStats(i).Attacker Or InStr(1, FullStats(i).Attacker, "SC:") <> 0 Then
            If listPlayers.Selected(c) = True Then
                IncludePlayer = True
            End If
            Exit For
        End If
    Next
    If listResults.Selected(FullStats(i).BattleID) = True And IncludePlayer Then
        FoundSummary = False
        For p = 0 To UBound(HTMLSummaryStats)
            If FullStats(i).Attacker = HTMLSummaryStats(p).Attacker Then
                CombineStats HTMLSummaryStats(p), FullStats(i)
                FoundSummary = True
            End If
        Next
        If FoundSummary = False Then
            CombineStats HTMLSummaryStats(UBound(HTMLSummaryStats)), FullStats(i)
            ReDim Preserve HTMLSummaryStats(UBound(HTMLSummaryStats) + 1)
        End If
    End If
Next

For i = 0 To UBound(HTMLSummaryStats) - 1
    CombineStats SummaryTotals, HTMLSummaryStats(i)
Next

SummaryTotals.Battles = TotalFights
RTB_Averages.Text = ""
RTB_Averages.SelBold = True
RTB_Averages.SelUnderline = True
RTB_Averages.SelText = ColumnText & vbNewLine
For i = 0 To UBound(HTMLSummaryStats) - 1
    If i = UBound(HTMLSummaryStats) - 1 Then
        RTB_Averages.SelUnderline = True
    End If
    RTB_Averages.SelBold = False
    RTB_Averages.SelText = PlayerText(HTMLSummaryStats, i, SummaryTotals.TotalDMG) & vbNewLine
Next
RTB_Averages.SelBold = True
RTB_Averages.SelText = FooterText(SummaryTotals, False)
Screen.MousePointer = vbDefault
End Sub








Private Sub cmdRefresh_Click()
Dim i, p, o
Dim MyStartTime As String, MyCurrentTime As String, HeaderString As String, PrevHeaderString As String, MyResult As String
Dim Craft() As udtCrafting
Dim FoundIt As Boolean, AddIt As Boolean

ReDim Preserve Craft(0)

RTB_Crafting.Text = ""
p = -1

For i = 0 To UBound(CraftingCSV)
    FoundIt = False
    With CraftingCSV(i)
        For o = 0 To UBound(Craft)
            AddIt = False
            MyResult = .Result
            If .CriticalFailure Then
                MyResult = Replace(MyResult, "Failure-", "Critical-")
            End If
            If Craft(o).Result = MyResult Then
                AddIt = True
            End If
            If checkCraft(3).Value = 1 And AddIt Then 'Time
                If comboTime.Text = "15 Mins" And .CurrentTime <> "" Then
                    If Right(.CurrentTime, 2) <= 15 Then
                        MyCurrentTime = Left(.CurrentTime, Len(.CurrentTime) - 2) & "00-15"
                    ElseIf Right(.CurrentTime, 2) <= 30 Then
                        MyCurrentTime = Left(.CurrentTime, Len(.CurrentTime) - 2) & "16-30"
                    ElseIf Right(.CurrentTime, 2) <= 45 Then
                        MyCurrentTime = Left(.CurrentTime, Len(.CurrentTime) - 2) & "31-45"
                    ElseIf Right(.CurrentTime, 2) <= 59 Then
                        MyCurrentTime = Left(.CurrentTime, Len(.CurrentTime) - 2) & "46-59"
                    End If
                ElseIf comboTime.Text = "30 Mins" Then
                    If Right(.CurrentTime, 2) <= 30 Then
                        MyCurrentTime = Left(.CurrentTime, Len(.CurrentTime) - 2) & "00-30"
                    ElseIf Right(.CurrentTime, 2) <= 59 Then
                        MyCurrentTime = Left(.CurrentTime, Len(.CurrentTime) - 2) & "31-59"
                    End If
                ElseIf comboTime.Text = "60 Mins" Then
                    MyCurrentTime = Left(.CurrentTime, Len(.CurrentTime) - 2) & "00"
                End If
                If Craft(o).CurrentTime = MyCurrentTime Then
                    AddIt = True
                Else
                    AddIt = False
                End If
            End If
            If checkCraft(0).Value = 1 And AddIt Then 'Day
                If Craft(o).DayType = .DayType Then
                    AddIt = True
                Else
                    AddIt = False
                End If
            End If
            If checkCraft(1).Value = 1 And AddIt Then 'Moon
                If Craft(o).MoonPhase = .MoonPhase Then
                    AddIt = True
                Else
                    AddIt = False
                End If
            End If
            If checkCraft(2).Value = 1 And AddIt Then 'Moon Perc
                If Craft(o).MoonPerc = .MoonPerc Then
                    AddIt = True
                Else
                    AddIt = False
                End If
            End If
            If checkCraft(4).Value = 1 And AddIt Then 'Direction
                If Craft(o).Direction = .Direction Then
                    AddIt = True
                Else
                    AddIt = False
                End If
            End If
            If AddIt Then
                Craft(o).Count = Craft(o).Count + .Count
                FoundIt = True
                Exit For
            Else
                FoundIt = False
            End If
        Next
        If Not FoundIt Then
            ReDim Preserve Craft(UBound(Craft) + 1)
            If checkCraft(3).Value = 1 And .CurrentTime <> "" Then  'Time
                If comboTime.Text = "15 Mins" Then
                    If Right(.CurrentTime, 2) <= 15 Then
                        MyCurrentTime = Left(.CurrentTime, Len(.CurrentTime) - 2) & "00-15"
                    ElseIf Right(.CurrentTime, 2) <= 30 Then
                        MyCurrentTime = Left(.CurrentTime, Len(.CurrentTime) - 2) & "16-30"
                    ElseIf Right(.CurrentTime, 2) <= 45 Then
                        MyCurrentTime = Left(.CurrentTime, Len(.CurrentTime) - 2) & "31-45"
                    ElseIf Right(.CurrentTime, 2) <= 59 Then
                        MyCurrentTime = Left(.CurrentTime, Len(.CurrentTime) - 2) & "46-59"
                    End If
                ElseIf comboTime.Text = "30 Mins" Then
                    If Right(.CurrentTime, 2) <= 30 Then
                        MyCurrentTime = Left(.CurrentTime, Len(.CurrentTime) - 2) & "00-30"
                    ElseIf Right(.CurrentTime, 2) <= 59 Then
                        MyCurrentTime = Left(.CurrentTime, Len(.CurrentTime) - 2) & "31-59"
                    End If
                ElseIf comboTime.Text = "60 Mins" Then
                    MyCurrentTime = Left(.CurrentTime, Len(.CurrentTime) - 2) & "00"
                End If
                Craft(UBound(Craft)).CurrentTime = MyCurrentTime
            End If
            If checkCraft(0).Value = 1 Then  'Day
                Craft(UBound(Craft)).DayType = .DayType
            End If
            If checkCraft(1).Value = 1 Then 'Moon
                Craft(UBound(Craft)).MoonPhase = .MoonPhase
            End If
            If checkCraft(2).Value = 1 Then 'Moon Perc
                Craft(UBound(Craft)).MoonPerc = .MoonPerc
            End If
            If checkCraft(4).Value = 1 Then 'Direction
                Craft(UBound(Craft)).Direction = .Direction
            End If
            Craft(UBound(Craft)).Result = MyResult
            Craft(UBound(Craft)).Count = .Count
        End If
    End With
Next

  
For i = 1 To UBound(Craft)
    With Craft(i)
        HeaderString = ""
        If checkCraft(0).Value = 1 And checkCraft(3).Value = 1 Then  'Day/Time
            HeaderString = HeaderString & .DayType & "(" & .CurrentTime & ")"
        ElseIf checkCraft(0).Value = 1 Then  'Day
            HeaderString = HeaderString & .DayType
        ElseIf checkCraft(3).Value = 1 Then  'Time
            HeaderString = HeaderString & .CurrentTime
        End If
            
        If checkCraft(1).Value = 1 And checkCraft(2).Value = 1 Then 'Moon/Perc
            If HeaderString <> "" Then
                HeaderString = HeaderString & " - "
            End If
            HeaderString = HeaderString & .MoonPhase & "(" & .MoonPerc & ")"
        ElseIf checkCraft(1).Value = 1 Then 'Moon
            If HeaderString <> "" Then
                HeaderString = HeaderString & " - "
            End If
            HeaderString = HeaderString & .MoonPhase
        ElseIf checkCraft(2).Value = 1 Then 'Moon Perc
            If HeaderString <> "" Then
                HeaderString = HeaderString & " - "
            End If
            HeaderString = HeaderString & .MoonPerc
        End If
        If checkCraft(4).Value = 1 Then 'Direction
            If HeaderString <> "" Then
                HeaderString = HeaderString & " - "
            End If
            HeaderString = HeaderString & .Direction
        End If
        If PrevHeaderString <> HeaderString Then
            RTB_Crafting.SelBold = True
            RTB_Crafting.SelText = vbNewLine & HeaderString & vbNewLine
            PrevHeaderString = HeaderString
            RTB_Crafting.SelBold = False
            RTB_Crafting.SelText = .Count & " - " & .Result & vbNewLine
        Else
            RTB_Crafting.SelBold = False
            RTB_Crafting.SelText = .Count & " - " & .Result & vbNewLine
        End If
    End With
Next
    
End Sub

Private Sub cmdSelect_Click()
Dim i
For i = 0 To listResults.ListCount - 1
    listResults.Selected(i) = True
Next
End Sub

Private Sub cmdSelectPlayers_Click()
Dim i
For i = 0 To listPlayers.ListCount - 1
    listPlayers.Selected(i) = True
Next
End Sub


Private Sub cmdUnselect_Click()
Dim i
For i = 0 To listResults.ListCount - 1
    listResults.Selected(i) = False
Next
End Sub





Private Sub cmdUnselectPlayers_Click()
Dim i
For i = 0 To listPlayers.ListCount - 1
    listPlayers.Selected(i) = False
Next
End Sub

Private Sub comboDisplay_Click()
Dim i As Integer, lf, AddLoot As String, Players() As String, PlayerName As String, FoundPlayer As Boolean, PlayerCount As Long, MyPos As Integer, pl As Integer
ReDim Players(0)

For i = 0 To mnuView.UBound
    mnuView(i).Checked = False
Next
For i = 0 To mnuViewPlayer.UBound
    mnuViewPlayer(i).Checked = False
Next
mnuView(comboDisplay.ListIndex).Checked = True
optionResults(0).Value = True
If comboDisplay.Text = "Report" Then
    frameEdit.Visible = False
    frameCraft.Visible = False
    RTB_Report.Visible = True
    RTB_User.Visible = False
    frameSummary.Visible = False
    RTB_Details.Visible = False
    frameChat.Visible = False
    RTB_Fish.Visible = False
    RTB_Log.Visible = False
ElseIf comboDisplay.Text = "Fishing" Then
    If RTB_Fish.Text = "" Then
        FishRPT
    End If
    frameCraft.Visible = False
    frameEdit.Visible = False
    RTB_Report.Visible = False
    RTB_User.Visible = False
    frameSummary.Visible = False
    RTB_Details.Visible = False
    RTB_Fish.Visible = True
    frameChat.Visible = False
    RTB_Log.Visible = False
ElseIf comboDisplay.Text = "Chat" Then
    RTB_Chat.Text = ""
    For i = 0 To optionChat.UBound
        optionChat(i).Value = False
    Next
    frameCraft.Visible = False
    RTB_Fish.Visible = False
    frameEdit.Visible = False
    RTB_Report.Visible = False
    RTB_User.Visible = False
    frameSummary.Visible = False
    RTB_Details.Visible = False
    frameChat.Visible = True
    RTB_Log.Visible = False
ElseIf comboDisplay.Text = "Summary" Then
    frameCraft.Visible = False
    RTB_Fish.Visible = False
    frameEdit.Visible = False
    RTB_Report.Visible = False
    RTB_User.Visible = False
    frameSummary.Visible = True
    RTB_Details.Visible = False
    frameChat.Visible = False
    RTB_Log.Visible = False
ElseIf comboDisplay.Text = "Details" Then
    frameCraft.Visible = False
    RTB_Fish.Visible = False
    frameEdit.Visible = False
    RTB_Report.Visible = False
    RTB_User.Visible = False
    frameSummary.Visible = False
    RTB_Details.Visible = True
    frameChat.Visible = False
    RTB_Log.Visible = False
ElseIf comboDisplay.Text = "Edit" Then
    frameCraft.Visible = False
    RTB_Fish.Visible = False
    frameEdit.Visible = True
    RTB_Report.Visible = False
    RTB_User.Visible = False
    frameSummary.Visible = False
    RTB_Details.Visible = False
    frameChat.Visible = False
    RTB_Log.Visible = False
ElseIf comboDisplay.Text = "FFXIP Log" Then
    frameCraft.Visible = False
    RTB_Fish.Visible = False
    frameEdit.Visible = False
    RTB_Report.Visible = False
    RTB_User.Visible = False
    frameSummary.Visible = False
    RTB_Details.Visible = False
    frameChat.Visible = False
    RTB_Log.Visible = True
ElseIf comboDisplay.Text = "Crafting" Then
    frameCraft.Visible = True
    RTB_Fish.Visible = False
    frameEdit.Visible = False
    RTB_Report.Visible = False
    RTB_User.Visible = False
    frameSummary.Visible = False
    RTB_Details.Visible = False
    frameChat.Visible = False
    RTB_Log.Visible = False
ElseIf comboDisplay.Text = "Loot!" Then
    frameCraft.Visible = False
    RTB_Log.Visible = False
    RTB_Fish.Visible = False
    frameEdit.Visible = False
    RTB_Report.Visible = False
    RTB_User.Visible = True
    frameSummary.Visible = False
    RTB_Details.Visible = False
    frameChat.Visible = False
    RTB_User.Text = ""
    RTB_User.TextRTF = ""
    RTB_User.SelStart = 0
    RTB_User.SelLength = Len(RTB_User.TextRTF)
    RTB_User.SelFontName = "Arial"
    RTB_User.SelBold = False
    RTB_User.SelColor = vbBlack
    
    RTB_User.SelBold = True
    RTB_User.SelColor = vbRed
    RTB_User.SelText = "Items/Gil:" & vbNewLine
    RTB_User.SelColor = vbBlack
    For lf = 0 To UBound(LootFound)
        RTB_User.SelBold = False
        If LootFound(lf) <> "" Then
            AddLoot = LootFound(lf)
            MyPos = InStr(1, AddLoot, " - ")
            AddLoot = Left$(AddLoot, MyPos - 1) & " - " & UCase(Mid$(AddLoot, MyPos + 3, 1)) & Mid$(AddLoot, MyPos + 4)
            RTB_User.SelBold = False
            RTB_User.SelColor = vbBlack
            RTB_User.SelText = vbTab & AddLoot & vbNewLine
        End If
    Next
          
    RTB_User.SelBold = True
    RTB_User.SelColor = vbRed
    RTB_User.SelText = vbNewLine & "Loot by Player:" & vbNewLine
    RTB_User.SelColor = vbBlack
    RTB_User.SelBold = False
    For lf = 0 To UBound(PlayerLoot)
        FoundPlayer = False
        AddLoot = PlayerLoot(lf)
        MyPos = InStr(1, AddLoot, ";")
        PlayerName = Mid(AddLoot, MyPos + 1)
        For pl = 0 To UBound(Players)
            If Players(pl) = PlayerName Then
                FoundPlayer = True
                Exit For
            End If
        Next
        If FoundPlayer = False Then
            ReDim Preserve Players(PlayerCount)
            Players(PlayerCount) = PlayerName
            PlayerCount = PlayerCount + 1
        End If
    Next
    
    For pl = 0 To UBound(Players)
        RTB_User.SelBold = True
        RTB_User.SelColor = vbBlack
        RTB_User.SelText = vbNewLine & vbTab & Players(pl) & vbNewLine
        RTB_User.SelBold = False
        
        For lf = 0 To UBound(PlayerLoot)
            AddLoot = PlayerLoot(lf)
            If AddLoot <> "" Then
                MyPos = InStr(1, AddLoot, ";")
                PlayerName = Mid(AddLoot, MyPos + 1)
                If PlayerName = Players(pl) Then
                    MyPos = InStr(1, AddLoot, ";")
                    RTB_User.SelBold = False
                    RTB_User.SelColor = vbBlack
                    RTB_User.SelText = vbTab & vbTab & Left(AddLoot, MyPos - 1) & vbNewLine
                End If
            End If
        Next
    Next
End If
End Sub


Private Sub comboMOB_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Dim i
    For i = 0 To listResults.ListCount - 1
        If InStr(1, LCase(listResults.List(i)), LCase(comboMOB.Text)) Then
            listResults.Selected(i) = True
        Else
            listResults.Selected(i) = False
        End If
    Next
End If
End Sub


Public Sub comboUser_Click()
PlayerReport comboUser.Text
RTB_User.SelStart = 0
frameEdit.Visible = False
RTB_Report.Visible = False
frameSummary.Visible = False
frameChat.Visible = False
frameCraft.Visible = False
RTB_User.Visible = True
RTB_Fish.Visible = False
RTB_Details.Visible = False
optionResults(1).Value = True
End Sub








Private Sub dirList_Change()
fileList.Path = dirList.Path
End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer
If KeyCode = 112 Then 'F1
    i = 0
ElseIf KeyCode = 113 Then
    i = 1
ElseIf KeyCode = 114 Then
    i = 2
ElseIf KeyCode = 115 Then
    i = 3
ElseIf KeyCode = 116 Then
    i = 4
ElseIf KeyCode = 117 Then
    i = 5
ElseIf KeyCode = 118 Then
    i = 6
Else
    i = 7
End If
If i <> 7 Then
    comboDisplay.ListIndex = i
    comboDisplay_Click
    optionResults(0).SetFocus
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim i
ReDim FullStats(0)
ReDim SummaryStats(0)
ReDim LootFound(0)
ReDim FishFound(0)
ReDim PlayerLoot(0)
ReDim ChatText(0)
ReDim CraftingCSV(0)
LastGather = True
Dim Updates As String
Dim f, ReportError As String
ErrorFile = FreeFile
Open App.Path & "\error_log.txt" For Output As #ErrorFile
comboTime.ListIndex = 0
For i = 0 To UBound(BeepSounds)
    BeepSounds(i) = GetSetting(App.Title, "Settings", "Sound-" & i, Default:="")
Next
For i = 0 To UBound(KeyLogs)
    KeyLogs(i, 0) = GetSetting(App.Title, "Keylogs", "Type/Key-" & i, Default:="")
    KeyLogs(i, 1) = GetSetting(App.Title, "Keylogs", "Duration-" & i, Default:="")
    KeyLogs(i, 2) = GetSetting(App.Title, "Keylogs", "Sound-" & i, Default:="")
    KeyLogs(i, 3) = GetSetting(App.Title, "Keylogs", "Message-" & i, Default:="")
Next

mnuRecalculate.Enabled = False
mnuExport.Enabled = False
mnuExportXML.Enabled = False
        
Me.Caption = "FFXIP - Online - " & App.Major & "." & App.Minor & "." & App.Revision
comboUser.ListIndex = 0
If GetSetting(App.Title, "Settings", "AutoCheck", Default:="") = "" Then
    If MsgBox("OK to always check for updates?", vbYesNo + vbQuestion, "Version Check") = vbYes Then
        SaveSetting App.Title, "Settings", "AutoCheck", "1"
    Else
        SaveSetting App.Title, "Settings", "AutoCheck", "0"
    End If
End If
BeepNotWave = GetSetting(App.Title, "Settings", "Beeps", Default:="1")
Dim ColumnText As String
For i = 0 To UBound(ReportOptions)
    If i = 0 Or i = 1 Or i = 2 Or i = 7 Or i = 18 Then
        ReportOptions(i) = GetSetting(App.Title, "Settings", "Report" & i, Default:=1)
    Else
        ReportOptions(i) = GetSetting(App.Title, "Settings", "Report" & i, Default:=0)
    End If
Next

HiddenAds = GetSetting(App.Title, "Settings", "Monkey", Default:="0")
If HiddenAds Then
    frameSupport.Visible = False
    frameEdit.Top = 100
    frameChat.Top = 100
    frameCraft.Top = 100
    frameSummary.Top = 100
    RTB_Chat.Top = 285
    RTB_Averages.Top = 285
    RTB_Report.Top = 100
    RTB_User.Top = 100
    RTB_Fish.Top = 100
    RTB_Log.Top = 100
    RTB_Details.Top = 100
End If
mnuEnableSounds.Checked = GetSetting(App.Title, "Settings", "EnableSounds", Default:="1")
mnuAltHome.Checked = GetSetting(App.Title, "Settings", "AltHome", Default:="1")
mnuUpdate.Checked = GetSetting(App.Title, "Settings", "AutoCheck", Default:="0")
mnuTray.Checked = GetSetting(App.Title, "Settings", "TrayIcon", Default:="1")
mnuOnly.Checked = GetSetting(App.Title, "Settings", "NewOnly", Default:="0")
mnuKeyEnable.Checked = GetSetting(App.Title, "Settings", "EnableKeyLogging", Default:="0")
mnuParserCommands.Checked = GetSetting(App.Title, "Settings", "ParserCommands", Default:="1")
ParserCommands = mnuParserCommands.Checked
If mnuKeyEnable.Checked Then
    timerKeyLogger.Enabled = True
End If
If mnuTray.Checked Then
    AddTray
End If
If GetSetting(App.Title, "Settings", "AutoCheck", Default:="") = "1" Or GetSetting(App.Title, "Settings", "AutoCheck", Default:="") = "True" Then
    Dim MyUpdate As String
    MyUpdate = inet.OpenURL("http://www.frontiernet.net/~Spyle/FFXI/update.txt")
    
    If Left$(MyUpdate, 7) = "Version" Then
        Dim MyPosA, MyPosB, MyVersion
        MyPosA = InStr(1, MyUpdate, "|")
        MyVersion = App.Major & "." & App.Minor & "." & App.Revision
        Clipboard.Clear
        Clipboard.SetText Mid$(MyUpdate, 10, MyPosA - 10)
        If MyVersion <> Mid$(MyUpdate, 10, MyPosA - 10) Then
            Do Until MyUpdate = ""
                MyPosA = InStr(1, MyUpdate, "|")
                If MyPosA <> 0 Then
                    Updates = Updates & Mid$(MyUpdate, 1, MyPosA - 1) & vbNewLine
                    MyUpdate = Mid$(MyUpdate, MyPosA + 1)
                Else
                    Updates = Updates & MyUpdate
                    MyUpdate = ""
                End If
            Loop
            If MsgBox(Updates & vbNewLine & vbNewLine & "Visit website?" & vbNewLine & "(www.frontiernet.net/~Spyle/FFXI/ffxi.html)", vbYesNo, "Update Info") = vbYes Then
                ShellExecute Me.hwnd, vbNullString, "http://www.frontiernet.net/~Spyle/FFXI/ffxi.html", vbNullString, "C:\", SW_SHOWNORMAL
            End If
        End If
    End If
End If


comboDisplay.ListIndex = 0
Dim FSO As FileSystemObject, MyX As Long, MyY As Long, MyWidth As Long, MyHeight As Long, MainSt As Integer
Set FSO = New FileSystemObject
If FSO.FolderExists("C:\Program Files\PlayOnline\SquareEnix\FINAL FANTASY XI\TEMP") = True Then
    dirList.Path = GetSetting(App.Title, "Settings", "LogPath", Default:="C:\Program Files\PlayOnline\SquareEnix\FINAL FANTASY XI\TEMP")
Else
    dirList.Path = GetSetting(App.Title, "Settings", "LogPath", Default:="C:\")
End If

MyX = GetSetting(App.Title, "Window Locations", "MainX", Default)
MyY = GetSetting(App.Title, "Window Locations", "MainY", Default)
MyWidth = GetSetting(App.Title, "Window Locations", "MainWidth", Default)
MyHeight = GetSetting(App.Title, "Window Locations", "MainHeight", Default)
MainSt = GetSetting(App.Title, "Window Locations", "MainSTATE", Default)
If MyWidth <> 0 And MyWidth < Screen.Width Then
   Me.Width = MyWidth
End If
If MyHeight <> 0 And MyHeight < Screen.Height Then
   Me.Height = MyHeight
End If
If MyX < 0 Then MyX = 0
If MyY < 0 Then MyY = 0
Me.Left = MyX
Me.Top = MyY
If MainSt = "Normal" Then
   Me.WindowState = 0
ElseIf MainSt = "Minimized" Then
   Me.WindowState = 1
ElseIf MainSt = "Maximized" Then
   Me.WindowState = 2
Else
   Me.WindowState = 0
End If
If Me.Left > Screen.Width Then
    Me.Left = 0
End If
If Me.Top > Screen.Height Then
    Me.Top = 0
End If
If Command$ = "" Then
    If MsgBox("If your PC explodes while using this program, it's not my fault." & vbNewLine & vbNewLine & "Please read the Notes/Known issues on the website before emailing me!" & vbNewLine & vbNewLine & "********TaruTaru heal Galka - Galka eat TaruTaru********", vbOKCancel + vbInformation, "Notice!") = vbCancel Then
        Unload Me
        Exit Sub
    End If
End If

Dim OpenLine As String, MyPos As Integer
If FSO.FileExists(App.Path & "\weaponskills.txt") Then
    f = FreeFile
    Open App.Path & "\weaponskills.txt" For Input As #f
    Do Until EOF(f)
        Line Input #f, OpenLine
        SkillList = SkillList & OpenLine
    Loop
    Close #f
End If
If FSO.FileExists(App.Path & "\spells.txt") Then
    f = FreeFile
    i = 0
    Open App.Path & "\spells.txt" For Input As #f
    Do Until EOF(f)
        Line Input #f, OpenLine
        MyPos = InStr(1, OpenLine, ",")
        If MyPos <> 0 Then
            ReDim Preserve SpellList(i)
            SpellList(i).Name = Left(OpenLine, MyPos - 1)
            SpellList(i).MPCost = Mid(OpenLine, MyPos + 1)
            i = i + 1
        End If
    Loop
    Close #f
Else
    ReDim Preserve SpellList(0)
    SpellList(0).Name = "Cure"
    SpellList(0).MPCost = 8
End If

If StartWithOpen <> "" Then
    OpenSingle = True
    Gather = False
    SingleFile = StartWithOpen
    frmRead.StartNew
End If

Dim SaveFile
If Command$ = "-g" Then 'Start with Gather to date
    OpenSingle = False
    GatherDate = True
    Gather = True
    CreateSingleFile
    Set SaveFile = FSO.CreateTextFile(SingleFile, True)
    SaveFile.Close
    Set SaveFile = Nothing
    frmRead.StartNew
    mnuStopGathering.Enabled = True
    mnuGather.Enabled = False
    mnuGatherDate.Enabled = False
    mnuParse.Enabled = False
    mnuOpen.Enabled = False
    mnuBoth.Enabled = False
    mnuBothFile.Enabled = False
ElseIf Command$ = "-p" Then 'Start with Parse
    OpenSingle = False
    Gather = False
    frmRead.StartNew
    mnuGather.Enabled = False
    mnuGatherDate.Enabled = False
    mnuOpen.Enabled = False
    mnuBoth.Enabled = False
    mnuBothFile.Enabled = False
ElseIf Command$ = "-gp" Then 'Start with Gather/Parse to date
    OpenSingle = False
    GatherDate = True
    ParseGather = True
    CreateSingleFile
    Set SaveFile = FSO.CreateTextFile(SingleFile, True)
    SaveFile.Close
    Set SaveFile = Nothing
    frmRead.StartNew
    mnuStopGathering.Enabled = True
    mnuGather.Enabled = False
    mnuGatherDate.Enabled = False
    mnuParse.Enabled = False
    mnuBoth.Enabled = False
    mnuBothFile.Enabled = False
    mnuOpen.Enabled = False
End If
Set FSO = Nothing
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If ReportError <> "" Then
    Dim Result
    If MsgBox("Send last error to website for debugging?", vbQuestion + vbYesNo, "Errors Found") = vbYes Then
        Screen.MousePointer = vbHourglass
        Result = inet.OpenURL("http://ffxi.mmorpgparsers.com/update_error.php?error_text=" & ReportError & "&version=" & App.Major & "." & App.Minor & "." & App.Revision)
        If Result = "error" Then
            MsgBox "Upload failed. Please email 'error_log.txt' to Spyle@Frontiernet.net", vbInformation, "Error"
        End If
        Screen.MousePointer = vbDefault
    End If
End If

On Error Resume Next
timerRead.Enabled = False
tIcon.Visible = False
Set tIcon = Nothing
End Sub


Private Sub Form_Resize()
On Error Resume Next
If mnuTray.Checked Then
    If Me.WindowState = 1 Then
        mnuRestore.Visible = True
        Me.Visible = False
        Exit Sub
    End If
End If
frameChat.Width = Me.Width - 225
frameSummary.Width = Me.Width - 225
frameEdit.Width = Me.Width - 225
frameCraft.Width = Me.Width - 225
frameSupport.Width = Me.Width - 225
imgAd.Left = (frameSupport.Width / 2) - (imgAd.Width / 2) - 750
imgPaypal.Left = imgAd.Left + imgAd.Width + 500
lblHide.Left = (frameSupport.Width - lblHide.Width) - 100

RTB_Report.Width = Me.Width - 225
RTB_Log.Width = Me.Width - 225
RTB_User.Width = Me.Width - 225
RTB_Chat.Width = Me.Width - 225
RTB_Crafting.Width = Me.Width - 225
RTB_Fish.Width = Me.Width - 225
RTB_Details.Width = Me.Width - 225
RTB_Averages.Width = Me.Width - 225
'lblInfo(0).Left = listResults.Width + 250
'lblInfo(1).Left = listResults.Width + 250
'lblInfo(0).Top = (frameEdit.Height - lblInfo(0).Height) - 200

If HiddenAds Then
    frameEdit.Height = Me.Height - 1300
Else
    frameEdit.Height = Me.Height - 2500
End If
listResults.Height = frameEdit.Height - 1000
cmdSelect.Top = listResults.Top + listResults.Height
cmdUnselect.Top = listResults.Top + listResults.Height
listPlayers.Height = frameEdit.Height - 1000
cmdSelectPlayers.Top = listPlayers.Top + listPlayers.Height
cmdUnselectPlayers.Top = listPlayers.Top + listPlayers.Height
comboMOB.Top = cmdSelectPlayers.Top + cmdSelectPlayers.Height


listResults.Width = (frameEdit.Width / 2) - 100
cmdSelect.Width = (listResults.Width / 2) - 50
cmdUnselect.Width = (listResults.Width / 2) - 50
cmdSelect.Left = listResults.Left + 50
cmdUnselect.Left = listResults.Left + cmdSelect.Width + 75

comboMOB.Left = listResults.Left + 50
comboMOB.Width = listResults.Width - 50

listPlayers.Left = listResults.Width + listResults.Left
listPlayers.Width = (frameEdit.Width / 2)
lbl(1).Left = listPlayers.Left
lbl(2).Left = comboMOB.Left + comboMOB.Width + 75
lbl(2).Top = comboMOB.Top + 75
cmdSelectPlayers.Width = (listPlayers.Width / 2) - 50
cmdUnselectPlayers.Width = (listPlayers.Width / 2) - 50
cmdSelectPlayers.Left = listPlayers.Left + 50
cmdUnselectPlayers.Left = listPlayers.Left + cmdSelectPlayers.Width + 75

If HiddenAds Then
    frameChat.Height = Me.Height - 1300
    frameCraft.Height = Me.Height - 1300
    frameSummary.Height = Me.Height - 1300
    RTB_Crafting.Height = frameCraft.Height - 320
    RTB_Chat.Height = frameChat.Height - 320
    RTB_Averages.Height = frameSummary.Height - 320
    RTB_Report.Height = Me.Height - 1300
    RTB_User.Height = Me.Height - 1300
    RTB_Fish.Height = Me.Height - 1300
    RTB_Log.Height = Me.Height - 1300
    RTB_Details.Height = Me.Height - 1300
Else
    frameChat.Height = Me.Height - 2500
    frameCraft.Height = Me.Height - 2500
    frameSummary.Height = Me.Height - 2500
    RTB_Chat.Height = frameChat.Height - 320
    RTB_Crafting.Height = frameCraft.Height - 320
    RTB_Averages.Height = frameSummary.Height - 320
    RTB_Report.Height = Me.Height - 2500
    RTB_User.Height = Me.Height - 2500
    RTB_Fish.Height = Me.Height - 2500
    RTB_Log.Height = Me.Height - 2500
    RTB_Details.Height = Me.Height - 2500
End If

Shape1.Width = Me.Width - 3100

lblStatus.Width = Me.Width - 3700

optionResults(0).Top = RTB_Report.Top + RTB_Report.Height + 90
optionResults(1).Top = RTB_Report.Top + RTB_Report.Height + 90
comboDisplay.Top = RTB_Report.Top + RTB_Report.Height + 40
comboUser.Top = RTB_Report.Top + RTB_Report.Height + 40
Shape1.Top = RTB_Report.Top + RTB_Report.Height + 40
Label1.Top = RTB_Report.Top + RTB_Report.Height + 75
lblStatus.Top = RTB_Report.Top + RTB_Report.Height + 75
End Sub


Private Sub AddTray()
On Error Resume Next
tIcon.Visible = False

Set tIcon = New TrayIcon
Set tIcon.OwnerForm = frmRead
tIcon.Icon = frmRead.Icon
tIcon.Tooltip = "FFXI Log Parser"
tIcon.Visible = True
End Sub
Private Sub Form_Unload(Cancel As Integer)
Close #ErrorFile
timerRead.Enabled = False
If Me.WindowState = 2 Then
    SaveSetting App.Title, "Window Locations", "MainSTATE", "Maximized"
ElseIf Me.WindowState = 0 Or Me.WindowState = 1 Then
    SaveSetting App.Title, "Window Locations", "MainSTATE", "Normal"
End If

Me.WindowState = vbNormal
SaveSetting App.Title, "Window Locations", "MainWidth", Me.Width
SaveSetting App.Title, "Window Locations", "MainHeight", Me.Height
SaveSetting App.Title, "Window Locations", "MainX", Me.Left
SaveSetting App.Title, "Window Locations", "MainY", Me.Top
End Sub







Private Sub imgAd_Click()
ShellExecute Me.hwnd, vbNullString, "http://tracking.ige.com/a/2105/b/1/e/38", vbNullString, "C:\", SW_SHOWNORMAL
End Sub

Private Sub imgPaypal_Click()
ShellExecute Me.hwnd, vbNullString, "https://www.paypal.com/cgi-bin/webscr?cmd=_xclick&business=Spyle%40Frontiernet%2enet&item_name=Parser Support&no_shipping=0&no_note=1&tax=0&currency_code=USD", vbNullString, "C:\", SW_SHOWNORMAL
End Sub



Private Sub lblHide_Click()
'Dim KeyCode, Result
'KeyCode = InputBox("Code:", "Hide")
'If KeyCode <> "" Then
'    Screen.MousePointer = vbHourglass
'    Result = inet.OpenURL("http://ffxi.mmorpgparsers.com/acti.php?code=" & KeyCode)
'    If Result <> "1" Then
'        MsgBox "Invalid Code.", vbInformation, "Failed"
'    Else
'        SaveSetting App.Title, "Settings", "Monkey", "1"
'        HiddenAds = True
'        frameSupport.Visible = False
'        frameEdit.Top = 100
'        frameChat.Top = 100
'        RTB_Chat.Top = 285
'        RTB_Report.Top = 100
'        RTB_User.Top = 100
'        RTB_Fish.Top = 100
'        RTB_Averages.Top = 100
'        RTB_Log.Top = 100
'        RTB_Details.Top = 100
'        Form_Resize
'    End If
'    Screen.MousePointer = vbDefault
'End If
End Sub


Private Sub lblStatus_Change()
If StopLogging Then
    lblStatus.Caption = "Waiting for start command..."
End If
If HasErrors Then
    lblStatus.ForeColor = vbRed
Else
    lblStatus.ForeColor = vbBlack
End If
End Sub



Private Sub mnuAbout_Click()
frmAbout.Show
End Sub





Private Sub mnuAltHome_Click()
If mnuAltHome.Checked = False Then
    mnuAltHome.Checked = True
Else
    mnuAltHome.Checked = False
End If

SaveSetting App.Title, "Settings", "AltHome", mnuAltHome.Checked
End Sub

Private Sub mnuBeep_Click()
frmBeep.Show
End Sub

Private Sub mnuBoth_Click()

Erase FullStats
ReDim FullStats(0)
Erase SummaryStats
ReDim SummaryStats(0)
If GetSetting(App.Title, "Settings", "LogPath", "") = "" Then
    mnuLocation_Click
End If

Dim FSO As FileSystemObject
Set FSO = New FileSystemObject
Dim SaveFile
OpenSingle = False
GatherDate = True
ParseGather = True
CreateSingleFile

Set SaveFile = FSO.CreateTextFile(SingleFile, True)
SaveFile.Close
Set SaveFile = Nothing
LastGather = True
frmRead.StartNew
mnuStopGathering.Enabled = True
mnuGather.Enabled = False
mnuGatherDate.Enabled = False
mnuParse.Enabled = False
mnuBoth.Enabled = False
mnuBothFile.Enabled = False
mnuOpen.Enabled = False
End Sub

Private Sub mnuBothFile_Click()
On Error GoTo Err_Handler
Erase FullStats
ReDim FullStats(0)
Erase SummaryStats
ReDim SummaryStats(0)
If GetSetting(App.Title, "Settings", "LogPath", "") = "" Then
    mnuLocation_Click
End If

Dim FSO As FileSystemObject
Set FSO = New FileSystemObject
Dim SaveFile
OpenSingle = False
GatherDate = False
ParseGather = True

CD1.Flags = cdlOFNHideReadOnly
CD1.DialogTitle = "Save Logs As"
CD1.Filter = "Gathered Logs (*.prs)|*.prs"
CD1.CancelError = True
CD1.InitDir = App.Path
CD1.ShowSave
If CD1.FileName <> "" Then
    Set SaveFile = FSO.CreateTextFile(CD1.FileName, True)
    SaveFile.Close
    Set SaveFile = Nothing
    SingleFile = CD1.FileName
    LastGather = True
    frmRead.StartNew
    mnuStopGathering.Enabled = True
    mnuGather.Enabled = False
    mnuGatherDate.Enabled = False
    mnuParse.Enabled = False
    mnuBoth.Enabled = False
    mnuBothFile.Enabled = False
    mnuOpen.Enabled = False
End If
Exit Sub
Err_Handler:
If Err.Number = 32755 Then
    Exit Sub
End If
End Sub


Public Sub mnuClear_Click()
Dim ClearTotals As udtStatistics
Dim ClearCraft As udtCrafting
Erase CraftingCSV
ReDim CraftingCSV(0)
'Crafting = ClearCraft
'CraftingHours = ClearCraft
ReadDPS_Start = False
ReadEXP_Start = False
Read_Start = False
Read_Stop = False
StopEXP = False
BeginDPS = False
ReadDPS_Stop = False
ReadEXP_Stop = False
StopDPS = False
ReadFISH_Start = False
StopLogging = False
FishHeader = ""
FishComment = ""
ExpType = 0

FightStartTime = Empty
FightStopTime = Empty
StartTime = Empty
StopTime = Empty
StartTimeDPS = Empty
StopTimeDPS = Empty
If ClearEdit Then
    listResults.Clear
    listPlayers.Clear
    comboUser.Clear
End If
Erase ChainExp
Erase DPS
Erase ChatText
ReDim ChatText(0)
Erase LootFound
ReDim LootFound(0)
Erase FishFound
ReDim FishFound(0)
Erase PlayerLoot
ReDim PlayerLoot(0)
Erase BattleStats
Erase SummaryStats
ReDim SummaryStats(0)
BattleTotals = ClearTotals

BattleID = 0
HasErrors = False
CurrentFight = ""
FightComment = ""
MonsterCheck = ""
TotalExp = 0
ErrorCount = 0
TotalExp = 0

RTB_Report.Text = ""
RTB_Fish.Text = ""
RTB_Chat.Text = ""
RTB_User.Text = ""
RTB_Averages.Text = ""
RTB_Details.Text = ""
End Sub

Private Sub mnuCommands_Click()
frmHelp.Show
End Sub



Private Sub mnuCrafting_Click()
ShellExecute Me.hwnd, vbNullString, "http://ffxi.mmorpgparsers.com/craft", vbNullString, "C:\", SW_SHOWNORMAL
End Sub

Private Sub mnuCSV_Click()
On Error GoTo Err_Handler
CD1.CancelError = True
CD1.Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly
CD1.Filter = "CSV (*.csv)|*.csv"
CD1.FileName = Format(Date, "MM-DD-YYYY") & ".csv"
CD1.DialogTitle = "Save Crafting As"
CD1.InitDir = GetSetting(App.Title, "Settings", "ExportPath", App.Path)
CD1.ShowSave
ExportCSV CD1.FileName
MsgBox "Done!", vbInformation, "Export"
Exit Sub
Err_Handler:
If Err.Number = 32755 Then
    Exit Sub
End If
End Sub

Private Sub mnuDamage_Click()
ShellExecute Me.hwnd, vbNullString, "http://ffxi.mmorpgparsers.com", vbNullString, "C:\", SW_SHOWNORMAL
End Sub

Private Sub mnuEnableSounds_Click()
If mnuEnableSounds.Checked = False Then
    mnuEnableSounds.Checked = True
Else
    mnuEnableSounds.Checked = False
End If

SaveSetting App.Title, "Settings", "EnableSounds", mnuEnableSounds.Checked
End Sub

Private Sub mnuExit_Click()
If Me.WindowState <> 0 Then Me.WindowState = 0
Me.Visible = True
Me.SetFocus
Unload Me
End Sub

Private Sub mnuExport_Click()
On Error GoTo Err_Handler
CD1.CancelError = True
CD1.Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly
CD1.Filter = "HTML (*.html)|*.html"
CD1.DialogTitle = "Save Export As"
CD1.FileName = Format(Date, "MM-DD-YYYY") & ".html"
CD1.InitDir = GetSetting(App.Title, "Settings", "ExportPath", App.Path)
CD1.ShowSave
Screen.MousePointer = vbHourglass
Dim MyAns
MyAns = MsgBox("Export summary only?", vbQuestion + vbYesNoCancel, "Export HTML")
If MyAns = vbYes Then
    ExportHTML True, CD1.FileName
ElseIf MyAns = vbNo Then
    ExportHTML False, CD1.FileName
Else
    Screen.MousePointer = vbDefault
    Exit Sub
End If
MsgBox "Done!", vbInformation, "Export"
Screen.MousePointer = vbDefault
Exit Sub
Err_Handler:
If Err.Number = 32755 Then
    Exit Sub
End If
End Sub

Private Sub mnuExportXML_Click()
On Error GoTo Err_Handler
CD1.CancelError = True
CD1.Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly
CD1.Filter = "XML (*.xml)|*.xml"
CD1.FileName = Format(Date, "MM-DD-YYYY") & ".xml"
CD1.DialogTitle = "Save Export As"
CD1.InitDir = GetSetting(App.Title, "Settings", "ExportPath", App.Path)
CD1.ShowSave
ExportXML CD1.FileName
MsgBox "Done!", vbInformation, "Export"
Exit Sub
Err_Handler:
If Err.Number = 32755 Then
    Exit Sub
End If
End Sub

Private Sub mnuFile_Click()
If LastGather = True Then
    mnuSaveLog.Enabled = False
Else
    mnuSaveLog.Enabled = True
End If
End Sub

Private Sub mnuGather_Click()
If GetSetting(App.Title, "Settings", "LogPath", "") = "" Then
    mnuLocation_Click
End If
On Error GoTo Err_Handler
Dim FSO As FileSystemObject
Set FSO = New FileSystemObject
Dim SaveFile
OpenSingle = False
Gather = True
GatherDate = False

CD1.Flags = cdlOFNHideReadOnly
CD1.DialogTitle = "Save Logs As"
CD1.Filter = "Gathered Logs (*.prs)|*.prs"
CD1.CancelError = True
CD1.InitDir = App.Path
CD1.ShowSave
If CD1.FileName <> "" Then
    Set SaveFile = FSO.CreateTextFile(CD1.FileName, True)
    SaveFile.Close
    Set SaveFile = Nothing
    SingleFile = CD1.FileName
    LastGather = True
    StartNew
    mnuStopGathering.Enabled = True
    mnuGather.Enabled = False
    mnuGatherDate.Enabled = False
    mnuParse.Enabled = False
    mnuOpen.Enabled = False
    mnuBoth.Enabled = False
    mnuBothFile.Enabled = False
End If
Exit Sub
Err_Handler:
If Err.Number = 32755 Then
    Exit Sub
End If
End Sub

Private Sub mnuGatherDate_Click()
If GetSetting(App.Title, "Settings", "LogPath", "") = "" Then
    mnuLocation_Click
End If

Dim FSO As FileSystemObject
Set FSO = New FileSystemObject
Dim SaveFile
OpenSingle = False
GatherDate = True
Gather = True
CreateSingleFile
Set SaveFile = FSO.CreateTextFile(SingleFile, True)
SaveFile.Close
Set SaveFile = Nothing
LastGather = True
frmRead.StartNew
mnuStopGathering.Enabled = True
mnuGather.Enabled = False
mnuGatherDate.Enabled = False
mnuParse.Enabled = False
mnuOpen.Enabled = False
mnuBoth.Enabled = False
mnuBothFile.Enabled = False
End Sub


Private Sub mnuKey_Click()
frmKeylogger.Show
End Sub

Private Sub mnuKeyEnable_Click()
If mnuKeyEnable.Checked = False Then
    mnuKeyEnable.Checked = True
    timerKeyLogger.Enabled = True
Else
    mnuKeyEnable.Checked = False
    timerKeyLogger.Enabled = False
End If

SaveSetting App.Title, "Settings", "EnableKeyLogging", mnuKeyEnable.Checked
End Sub

Private Sub mnuLocation_Click()
On Error GoTo Err_Handler
If GetSetting(App.Title, "Settings", "NewUserA", Default:=True) Then
    MsgBox "Select the folder where the FFXI log files are located by selecting one of the log files." & vbNewLine & vbNewLine & "Usually: ""C:\Program Files\PlayOnline\SquareEnix\FINAL FANTASY XI\TEMP""", vbInformation, "Folder Select"
    SaveSetting App.Title, "Settings", "NewUserA", False
End If
CD1.CancelError = True
CD1.Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly
CD1.Filter = "Logs (*.log)|*.log"
CD1.DialogTitle = "Select ANY FFXI Log File"
CD1.InitDir = GetSetting(App.Title, "Settings", "LogPath", "C:\Program Files\PlayOnline\SquareEnix\FINAL FANTASY XI\TEMP")
CD1.ShowOpen

SaveSetting App.Title, "Settings", "LogPath", Replace(CD1.FileName, CD1.FileTitle, "")
Exit Sub
Err_Handler:
If Err.Number = 32755 Then
    Exit Sub
End If
End Sub



Private Sub mnuMainExit_Click()
Unload Me
End Sub

Private Sub mnuOnlineSetup_Click()
frmSetup.Show
End Sub

Private Sub mnuOnlineTransmit_Click()
frmTransmit.Show
End Sub

Private Sub mnuOnly_Click()
If mnuOnly.Checked = False Then
    mnuOnly.Checked = True
Else
    mnuOnly.Checked = False
End If

SaveSetting App.Title, "Settings", "NewOnly", mnuOnly.Checked
End Sub

Private Sub mnuOpen_Click()
On Error GoTo Err_Handler
Erase FullStats
ReDim FullStats(0)
Erase SummaryStats
ReDim SummaryStats(0)
OpenSingle = True
Gather = False
CD1.CancelError = True
CD1.Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly
CD1.Filter = "Gathered Logs (*.log;*.prs)|*.log;*.prs"
CD1.DialogTitle = "Select Gathered Log File"
CD1.InitDir = App.Path
CD1.ShowOpen

If CD1.FileName <> "" Then
    Screen.MousePointer = vbHourglass
    DoEvents
    SingleFile = CD1.FileName
    If InStr(1, CD1.FileName, "EditFile.log") = 0 Then
        frmRead.StartNew
    Else
        MsgBox "Unable to open '" & SingleFile & "'", vbCritical, "Error"
    End If
    Screen.MousePointer = vbDefault
End If
Exit Sub
Err_Handler:
If Err.Number = 32755 Then
    Exit Sub
End If
End Sub

Private Sub mnuParse_Click()
If GetSetting(App.Title, "Settings", "LogPath", "") = "" Then
    mnuLocation_Click
End If
dirList.Path = GetSetting(App.Title, "Settings", "LogPath", "C:\")
fileList.Path = GetSetting(App.Title, "Settings", "LogPath", "C:\")
If mnuParse.Caption = "&Start Parsing" Then
    Erase FullStats
    ReDim FullStats(0)
    Erase SummaryStats
    ReDim SummaryStats(0)
    OpenSingle = False
    Gather = False
    LastGather = False
    frmRead.StartNew
    mnuGather.Enabled = False
    mnuGatherDate.Enabled = False
    mnuOpen.Enabled = False
    mnuBoth.Enabled = False
    mnuBothFile.Enabled = False
ElseIf mnuParse.Caption = "&Stop Parsing" Then
    FishRPT
    Erase FishFound
    ReDim FishFound(0)
    timerRead.Enabled = False
    mnuParse.Caption = "&Start Parsing"
    If Gather Then
        lblStatus.Caption = "Stopped - File saved as '" & SingleFile & "'"
    Else
        lblStatus.Caption = "Stopped - Waiting."
    End If
    mnuGather.Enabled = True
    mnuGatherDate.Enabled = True
    mnuOpen.Enabled = True
    mnuBoth.Enabled = True
    mnuBothFile.Enabled = True
End If


End Sub








Private Sub mnuParserCommands_Click()
If mnuParserCommands.Checked = False Then
    mnuParserCommands.Checked = True
    ParserCommands = True
Else
    mnuParserCommands.Checked = False
    ParserCommands = True
End If

SaveSetting App.Title, "Settings", "ParserCommands", mnuParserCommands.Checked
End Sub



Private Sub mnuRecalculate_Click()
RecalculateData
End Sub

Private Sub mnuReport_Click()
frmReport.Show
End Sub

Private Sub mnuRestore_Click()
If Me.WindowState <> 0 Then Me.WindowState = 0
Me.Visible = True
Me.SetFocus
mnuRestore.Visible = False
End Sub



Private Sub mnuSave_Click()
If optionResults(0).Value = True Then
    If comboDisplay.Text = "Report" Then
        RTB_Report_DblClick
    ElseIf comboDisplay.Text = "Fishing" Then
        RTB_Fish_DblClick
    ElseIf comboDisplay.Text = "Chat" Then
        RTB_Chat_DblClick
    ElseIf comboDisplay.Text = "Summary" Then
        RTB_Averages_DblClick
    ElseIf comboDisplay.Text = "Details" Then
        RTB_Details_DblClick
    ElseIf comboDisplay.Text = "Loot!" Then
        RTB_User_DblClick
    End If
Else
    RTB_User_DblClick
End If
End Sub

Private Sub mnuSaveLog_Click()
Dim FSO As FileSystemObject
Set FSO = New FileSystemObject

CD1.Flags = cdlOFNHideReadOnly Or &H2 Or &H4 Or &H8 Or &H2000
CD1.DialogTitle = "Save Data As..."
CD1.Filter = "Gathered Logs (*.prs)|*.prs"
CD1.CancelError = True
CD1.InitDir = App.Path
CD1.ShowSave

If CD1.FileName <> "" Then
    FSO.CopyFile App.Path & "\ffxip_tmp_.tmp", CD1.FileName, True
    MsgBox "File Saved As:" & vbNewLine & vbNewLine & CD1.FileName, vbInformation, "Save"
End If
Set FSO = Nothing
End Sub

Private Sub mnuStopGathering_Click()
FishRPT
Erase FishFound
ReDim FishFound(0)
timerRead.Enabled = False
If Gather = True Or ParseGather = True Then
    lblStatus.Caption = "Stopped - File saved as '" & SingleFile & "'"
Else
    lblStatus.Caption = "Stopped - Waiting."
End If
mnuParse.Caption = "&Start Parsing"
mnuStopGathering.Enabled = False
mnuGather.Enabled = True
mnuGatherDate.Enabled = True
mnuParse.Enabled = True
mnuBoth.Enabled = True
mnuBothFile.Enabled = True
mnuOpen.Enabled = True
End Sub

Private Sub mnuTray_Click()
If mnuTray.Checked = False Then
    mnuTray.Checked = True
    AddTray
Else
    mnuTray.Checked = False
    tIcon.Visible = False
    Set tIcon = Nothing
End If
SaveSetting App.Title, "Settings", "TrayIcon", mnuTray.Checked
End Sub

Private Sub mnuUpdate_Click()
If mnuUpdate.Checked = False Then
    mnuUpdate.Checked = True
Else
    mnuUpdate.Checked = False
End If

SaveSetting App.Title, "Settings", "AutoCheck", mnuUpdate.Checked
End Sub



Private Sub mnuView_Click(Index As Integer)
Dim i As Integer
For i = 0 To mnuView.UBound
    mnuView(i).Checked = False
Next
For i = 0 To mnuViewPlayer.UBound
    mnuViewPlayer(i).Checked = False
Next
mnuView(Index).Checked = True
If comboDisplay.ListIndex <> Index Then
    comboDisplay.ListIndex = Index
Else
    comboDisplay_Click
End If
optionResults(0).SetFocus
End Sub


Private Sub mnuViewPlayer_Click(Index As Integer)
Dim i As Integer
For i = 0 To mnuViewPlayer.UBound
    mnuViewPlayer(i).Checked = False
Next
For i = 0 To mnuView.UBound
    mnuView(i).Checked = False
Next
mnuViewPlayer(Index).Checked = True
If comboUser.ListIndex <> Index Then
    comboUser.ListIndex = Index
Else
    comboUser_Click
End If
optionResults(0).SetFocus
End Sub


Private Sub optionChat_Click(Index As Integer)
Dim i, OldLabel As String

OldLabel = lblStatus.Caption
RTB_Chat.Text = ""
RTB_Chat.Font.Name = "Arial"
Screen.MousePointer = vbHourglass
RTB_Chat.Visible = False
frameChat.Enabled = False
RTB_Chat.SelStart = 0
For i = 0 To UBound(ChatText)
    lblStatus.Caption = "Reading " & i & " of " & UBound(ChatText) & " chat messages."
    DoEvents
    
    If Trim$(ChatText(i)) <> "" Then
        LineType = Right(ChatText(i), 2)
        If LineType = "09" Or LineType = "01" Then 'say
            RTB_Chat.SelColor = &H404040
        ElseIf LineType = "0a" Or LineType = "02" Then 'shout
            RTB_Chat.SelColor = &HC0&
        ElseIf LineType = "0c" Or LineType = "04" Then 'tell
            RTB_Chat.SelColor = &HC000C0
        ElseIf LineType = "0d" Or LineType = "05" Then 'party
            RTB_Chat.SelColor = &HC00000
        ElseIf LineType = "0e" Or LineType = "06" Then 'ls
            RTB_Chat.SelColor = &H8000&
        ElseIf LineType = "0f" Then 'ls
            RTB_Chat.SelColor = &H0&
        End If
        If InStr(1, "02,01,05,06,04", LineType) Then
           RTB_Chat.SelBold = True
        Else
           RTB_Chat.SelBold = False
        End If

        If Index = 0 Then 'say
            If LineType = "09" Or LineType = "01" Then
                RTB_Chat.SelText = Replace(Replace(Mid$(ChatText(i), 3, Len(ChatText(i)) - 4), "ï'", "["), "ï(", "]")
                RTB_Chat.SelText = vbNewLine
                RTB_Chat.SelStart = Len(RTB_Chat.Text)
            End If
        ElseIf Index = 1 Then 'shout
            If LineType = "0a" Or LineType = "02" Then
                RTB_Chat.SelText = Replace(Replace(Mid$(ChatText(i), 3, Len(ChatText(i)) - 4), "ï'", "["), "ï(", "]")
                RTB_Chat.SelText = vbNewLine
                RTB_Chat.SelStart = Len(RTB_Chat.Text)
            End If
        ElseIf Index = 2 Then 'tell
            If LineType = "0c" Or LineType = "04" Then
                RTB_Chat.SelText = Replace(Replace(Mid$(ChatText(i), 3, Len(ChatText(i)) - 4), "ï'", "["), "ï(", "]")
                RTB_Chat.SelText = vbNewLine
                RTB_Chat.SelStart = Len(RTB_Chat.Text)
            End If
        ElseIf Index = 3 Then 'party
            If LineType = "0d" Or LineType = "05" Then
                RTB_Chat.SelText = Replace(Replace(Mid$(ChatText(i), 3, Len(ChatText(i)) - 4), "ï'", "["), "ï(", "]")
                RTB_Chat.SelText = vbNewLine
                RTB_Chat.SelStart = Len(RTB_Chat.Text)
            End If
        ElseIf Index = 4 Then 'ls
            If LineType = "0e" Or LineType = "06" Then
                RTB_Chat.SelText = Replace(Replace(Mid$(ChatText(i), 3, Len(ChatText(i)) - 4), "ï'", "["), "ï(", "]")
                RTB_Chat.SelText = vbNewLine
                RTB_Chat.SelStart = Len(RTB_Chat.Text)
            End If
        ElseIf Index = 5 Then 'all
            If LineType = "0f" Then
                If Right(ChatText(i), 5) = "1 0f" Then ChatText(i) = Replace(ChatText(i), "1 ", "")
            End If
            RTB_Chat.SelText = Replace(Replace(Mid$(ChatText(i), 3, Len(ChatText(i)) - 4), "ï'", "["), "ï(", "]")
            RTB_Chat.SelText = vbNewLine
            RTB_Chat.SelStart = Len(RTB_Chat.Text)
        ElseIf Index = 6 Then 'all
            If LineType = "0f" Then
                If Right(ChatText(i), 5) = "1 0f" Then ChatText(i) = Replace(ChatText(i), "1 ", "")
                RTB_Chat.SelText = Replace(Replace(Mid$(ChatText(i), 3, Len(ChatText(i)) - 4), "ï'", "["), "ï(", "]")
                RTB_Chat.SelText = vbNewLine
                RTB_Chat.SelStart = Len(RTB_Chat.Text)
            End If
        End If
    End If
Next
frameChat.Enabled = True
RTB_Chat.Visible = True
lblStatus = OldLabel
Screen.MousePointer = vbDefault
End Sub



Private Sub optionResults_Click(Index As Integer)
If Index = 0 Then
    comboDisplay_Click
ElseIf Index = 1 Then
    comboUser_Click
End If
End Sub

Private Sub optionSummary_Click(Index As Integer)
If optionSummary(0).Value = True Then
    Dim EstDPS As String, dp As Integer, p, i
    With RTB_Averages
        .TextRTF = ""
        .SelBold = True
        .SelText = "Experience" & vbNewLine
        .SelBold = False
        
        If TotalExp <> 0 And StartTime <> Empty And StopTime <> Empty Then
          .SelText = "Start: " & StartTime & " / Stop: " & StopTime & vbNewLine & "Total Exp: " & TotalExp & vbNewLine & "Per Hour: " & Round(TotalExp / DateDiff("n", StartTime, StopTime), 2) * 60 & vbNewLine & "Per Minute: " & Round(TotalExp / DateDiff("n", StartTime, StopTime), 2) & vbNewLine & vbNewLine
        ElseIf TotalExp <> 0 And StartTime <> Empty Then
          .SelText = "Start: " & StartTime & " / Stop: " & Now & vbNewLine & "Total Exp: " & TotalExp & vbNewLine & "Per Hour: " & Round(TotalExp / DateDiff("n", StartTime, Now), 2) * 60 & vbNewLine & "Per Minute: " & Round(TotalExp / DateDiff("n", StartTime, Now), 2) & vbNewLine & vbNewLine
        Else
          .SelText = "Start: " & StartTime & vbNewLine & "Total Exp: " & TotalExp & vbNewLine & "Per Hour: 0" & vbNewLine & "Per Minute: 0" & vbNewLine & vbNewLine
        End If
        
        .SelBold = True
        .SelText = "Experience Chains" & vbNewLine
        .SelBold = False
        For p = 0 To UBound(ChainExp)
            If ChainExp(p, 0) <> 0 Then
                .SelBold = False
                .SelText = "EXP Chain #" & p & ": " & CStr(ChainExp(p, 0)) & " - " & CStr(ChainExp(p, 1)) & " times, " & Round(ChainExp(p, 0) / ChainExp(p, 1), 2) & " average." & vbNewLine
            End If
        Next
        .SelText = vbNewLine
    End With
    For i = 0 To UBound(SummaryStats) - 1
        With SummaryStats(i)
            RTB_Averages.SelBold = True
            RTB_Averages.SelText = .Attacker & vbNewLine
            RTB_Averages.SelBold = False
            RTB_Averages.SelText = ResizePart("Total Fights: ", 1500) & vbTab & .Battles & vbNewLine
            If .Battles <> 0 Then
                RTB_Averages.SelBold = False
                RTB_Averages.SelText = ResizePart("Average Damage: ", 1500) & vbTab & Round(.TotalDMG / .Battles, 2) & vbNewLine
                RTB_Averages.SelBold = False
                RTB_Averages.SelText = ResizePart("Average Percent: ", 1500) & vbTab & Round((.Percent / .Battles), 2) & vbNewLine
                RTB_Averages.SelBold = False
                RTB_Averages.SelText = ResizePart("Average Accuracy: ", 1500) & vbTab & Round((.Accuracy / .Battles), 2) & vbNewLine
            Else
                RTB_Averages.SelBold = False
                RTB_Averages.SelText = ResizePart("Average Damage: ", 1500) & vbTab & "0.00" & vbNewLine
                RTB_Averages.SelBold = False
                RTB_Averages.SelText = ResizePart("Average Percent: ", 1500) & vbTab & "0.00" & vbNewLine
                RTB_Averages.SelBold = False
                RTB_Averages.SelText = ResizePart("Average Accuracy: ", 1500) & vbTab & "0.00" & vbNewLine
            End If
      
            EstDPS = ""
            For dp = 0 To UBound(DPS)
              If Trim(DPS(dp, 0)) = .Attacker Then
                  If DPS(dp, 0) <> "" Then
                        If DPS(dp, 1) <> "0" And DPS(dp, 2) <> "0" And DPS(dp, 2) <> "" And DPS(dp, 1) <> "" Then
                            EstDPS = Round(CDbl(DPS(dp, 1)) / CDbl(DPS(dp, 2)), 2) & " (" & DPS(dp, 2) & " seconds / " & DPS(dp, 1) & " dmg)"
                        Else
                            EstDPS = "0.00"
                        End If
                      Exit For
                  End If
              End If
            Next
            RTB_Averages.SelBold = False
            RTB_Averages.SelText = ResizePart("Estimated DPS: ", 1500) & vbTab & EstDPS & vbNewLine & vbNewLine
        End With
    Next
    RTB_Averages.SelStart = 0
Else
    displaySummary
End If
End Sub

Private Sub RTB_Averages_DblClick()
On Error GoTo Err_Handler
CD_Save.CancelError = True
CD_Save.Flags = cdlOFNOverwritePrompt
CD_Save.Filter = "Rich Text File (*.rtf)|*.rtf|Plain Text File (*.txt)|*.txt"
CD_Save.FileName = Format(Date, "MMDDYY") & "-Averages.rtf"

CD_Save.ShowSave
If CD_Save.FileName <> "0" And CD_Save.FileName <> "" Then
    If Right$(CD_Save.FileName, 3) = "rtf" Then
        RTB_Averages.SaveFile CD_Save.FileName
    Else
        RTB_Averages.SaveFile CD_Save.FileName, rtfText
    End If
End If
Exit Sub

Err_Handler:
If Err.Number = 32755 Then
    Exit Sub
End If
End Sub


Private Sub RTB_Chat_DblClick()
On Error GoTo Err_Handler
CD_Save.CancelError = True
CD_Save.Flags = cdlOFNOverwritePrompt
CD_Save.Filter = "Rich Text File (*.rtf)|*.rtf|Plain Text File (*.txt)|*.txt"
CD_Save.FileName = Format(Date, "MMDDYY") & "-Chat.rtf"

CD_Save.ShowSave
If CD_Save.FileName <> "0" And CD_Save.FileName <> "" Then
    If Right$(CD_Save.FileName, 3) = "rtf" Then
        RTB_Chat.SaveFile CD_Save.FileName
    Else
        RTB_Chat.SaveFile CD_Save.FileName, rtfText
    End If
End If
Exit Sub

Err_Handler:
If Err.Number = 32755 Then
    Exit Sub
End If
End Sub


Private Sub RTB_Details_DblClick()
On Error GoTo Err_Handler
CD_Save.CancelError = True
CD_Save.Flags = cdlOFNOverwritePrompt
CD_Save.Filter = "Rich Text File (*.rtf)|*.rtf|Plain Text File (*.txt)|*.txt"
CD_Save.FileName = Format(Date, "MMDDYY") & "-Details.rtf"

CD_Save.ShowSave
If CD_Save.FileName <> "0" And CD_Save.FileName <> "" Then
    If Right$(CD_Save.FileName, 3) = "rtf" Then
        RTB_Details.SaveFile CD_Save.FileName
    Else
        RTB_Details.SaveFile CD_Save.FileName, rtfText
    End If
End If
Exit Sub

Err_Handler:
If Err.Number = 32755 Then
    Exit Sub
End If
End Sub








Private Sub RTB_Fish_DblClick()
On Error GoTo Err_Handler
CD_Save.CancelError = True
CD_Save.Flags = cdlOFNOverwritePrompt
CD_Save.Filter = "Rich Text File (*.rtf)|*.rtf|Plain Text File (*.txt)|*.txt"
CD_Save.FileName = Format(Date, "MMDDYY") & "-Fish.rtf"

CD_Save.ShowSave
If CD_Save.FileName <> "0" And CD_Save.FileName <> "" Then
    If Right$(CD_Save.FileName, 3) = "rtf" Then
        RTB_Fish.SaveFile CD_Save.FileName
    Else
        RTB_Fish.SaveFile CD_Save.FileName, rtfText
    End If
End If
Exit Sub

Err_Handler:
If Err.Number = 32755 Then
    Exit Sub
End If
End Sub


Private Sub RTB_Report_DblClick()
On Error GoTo Err_Handler
CD_Save.CancelError = True
CD_Save.Flags = cdlOFNOverwritePrompt
CD_Save.Filter = "Rich Text File (*.rtf)|*.rtf|Plain Text File (*.txt)|*.txt"
CD_Save.FileName = Format(Date, "MMDDYY") & "-Report.rtf"

CD_Save.ShowSave
If CD_Save.FileName <> "0" And CD_Save.FileName <> "" Then
    If Right$(CD_Save.FileName, 3) = "rtf" Then
        RTB_Report.SaveFile CD_Save.FileName
    Else
        RTB_Report.SaveFile CD_Save.FileName, rtfText
    End If
End If
Exit Sub

Err_Handler:
If Err.Number = 32755 Then
    Exit Sub
End If
End Sub




Private Sub RTB_User_DblClick()
On Error GoTo Err_Handler
CD_Save.CancelError = True
CD_Save.Flags = cdlOFNOverwritePrompt
CD_Save.Filter = "Rich Text File (*.rtf)|*.rtf|Plain Text File (*.txt)|*.txt"
CD_Save.FileName = Format(Date, "MMDDYY") & "-" & comboUser.Text & ".rtf"

CD_Save.ShowSave
If CD_Save.FileName <> "0" And CD_Save.FileName <> "" Then
    If Right$(CD_Save.FileName, 3) = "rtf" Then
        RTB_User.SaveFile CD_Save.FileName
    Else
        RTB_User.SaveFile CD_Save.FileName, rtfText
    End If
End If
Exit Sub

Err_Handler:
If Err.Number = 32755 Then
    Exit Sub
End If
End Sub




Private Sub tIcon_DblClick()
If Me.WindowState = 1 Then
    If Me.WindowState <> 0 Then Me.WindowState = 0
    Me.Visible = True
    Me.SetFocus
    mnuRestore.Visible = False
Else
    Me.WindowState = 1
    Me.Visible = False
    mnuRestore.Visible = True
End If
End Sub

Private Sub tIcon_MouseUp(ByVal button As Integer)
If button = 2 Then
    PopupMenu mnuIcon
End If
End Sub


Private Sub timerAd_Timer()
If IGE Then
    imgAd.Picture = imgA.Picture
    IGE = False
Else
    imgAd.Picture = imgB.Picture
    IGE = True
End If
End Sub

Private Sub timerBeeps_Timer()
On Error Resume Next
If mnuEnableSounds.Checked = True Then
    Dim i As Integer, b As Integer, StartCounterTime As Date
    For i = 0 To UBound(TimerStart)
        If IsDate(TimerStart(i, 0)) Then
            StartCounterTime = DateValue(TimerStart(i, 0)) & " " & (TimeValue(TimerStart(i, 0)) + TimeValue(TimerStart(i, 1)))
            If StartCounterTime <= Now Then
                For b = 1 To TimerStart(i, 2)
                    If BeepNotWave Then
                        Call Beep(100, 100)
                    Else
                        If BeepSounds(CDbl(TimerStart(i, 2)) - 1) <> "" Then
                            PlaySound BeepSounds(CDbl(TimerStart(i, 2)) - 1), 0&, SND_FILENAME Or SND_ASYNC
                        Else
                            Call Beep(100, 100)
                        End If
                        Exit For
                    End If
                Next
                
                TimerStart(i, 0) = ""
                TimerStart(i, 1) = ""
                TimerStart(i, 2) = ""
                TimerStart(i, 3) = ""
            End If
        End If
    Next
End If
End Sub

Private Sub timerKeyLogger_Timer()

Dim VS(1 To 255) As Long
Dim d(255) As Byte
Dim X As Long, k As Integer, t As Integer
Dim i As Integer, o As String, results As Integer

For i = 1 To 126
    results = 0
    results = GetAsyncKeyState(i)
    o = ""
    
    If results = -32767 Then
    'If i <> 116 Then Stop
        'This is the case for only a-z
        If (i > 64 And i < 91) Then
            If GetKeyState(i) <> VS(i) Then
                VS(i) = GetKeyState(i)
                o = LCase(Chr(i))
                If GetAsyncKeyState(&H10) Then
                    o = "SHIFT-" & o
                End If
                If GetAsyncKeyState(&H11) Then
                    o = "CTRL-" & o
                End If
                If GetAsyncKeyState(&H12) Then
                    o = "ALT-" & o
                End If
            End If
        End If
        
        'case for the numbers
        If (i < 58 And i > 47) Then
            If GetKeyState(i) <> VS(i) Then
                VS(i) = GetKeyState(i)
                d(&H10) = IIf(GetAsyncKeyState(&H10) >= 0, 0, 255)
                ToAscii i, 0, d(0), X, 0
                o = Chr(X)
                If o = "!" Then o = "1"
                If o = "@" Then o = "2"
                If o = "#" Then o = "3"
                If o = "$" Then o = "4"
                If o = "%" Then o = "5"
                If o = "^" Then o = "6"
                If o = "&" Then o = "7"
                If o = "*" Then o = "8"
                If o = "(" Then o = "9"
                If o = ")" Then o = "0"
                If GetAsyncKeyState(&H10) Then
                    o = "SHIFT-" & o
                End If
                If GetAsyncKeyState(&H11) Then
                    o = "CTRL-" & o
                End If
                If GetAsyncKeyState(&H12) Then
                    o = "ALT-" & o
                End If
            End If
        End If
        
        If i = 32 Then o = " "
        If i = 8 Then o = Chr(8)
        For k = 0 To UBound(KeyLogs)
            If UCase(o) = UCase(KeyLogs(k, 0)) And KeyLogs(k, 0) <> "" Then
                For t = 0 To UBound(TimerStart)
                    If TimerStart(t, 0) = "" Then
                        TimerStart(t, 0) = Now
                        TimerStart(t, 1) = KeyLogs(k, 1)
                        TimerStart(t, 2) = KeyLogs(k, 2)
                        TimerStart(t, 3) = KeyLogs(k, 3)
                        Exit For
                    End If
                Next
            End If
        Next
    End If

Next i
End Sub
Private Sub timerAltHome_Timer()
If mnuAltHome.Checked = True Then
    Dim VS(1 To 255) As Long
    Dim d(255) As Byte
    Dim results As Integer
    timerAltHome.Interval = 1
    results = GetAsyncKeyState(36)
    If results = -32767 Then
        If GetAsyncKeyState(&H12) And GetAsyncKeyState(&H24) Then
            SendKeys AltHome
            timerAltHome.Interval = 1500
        End If
    End If
End If
End Sub


Private Sub timerRead_Timer()
      Dim z As Integer, o As Integer, i As Integer, f
      Dim MyDate As Date
      Dim FullFile() As String, CurrentLine As String, PrevLine As String, MyPosAdd As Integer, MyPos As Long, MyPos2 As Integer, CurrentFile As String
      Dim Index As Long

10    If fileList.ListCount <> 0 Then

20        If lblStatus.Caption = "Too many errors - Parsing stopped for this log." Then
30        Else
40            lblStatus.Caption = "Errors: " & HasErrors & " - " & "Waiting for new log...."
50        End If
60        DoEvents
    
70        fileListBox.Clear

80        fileList.Refresh
90        For i = 0 To fileList.ListCount - 1
100           fileList.ListIndex = i
110           fileListBox.AddItem Format(FileDateTime(dirList.Path & "\" & fileList.FileName), "MM/DD HhNnSs") & " - " & fileList.Path & "\" & fileList.FileName
120       Next
    
130       fileListBox.ListIndex = fileListBox.ListCount - 1
140       If LastItem <> fileListBox.Text Then
    
150           If Gather = True Or ParseGather = True Then
160               If GatherDate Then
170                 CreateSingleFile
180               End If
                  Dim EditFile
190               EditFile = FreeFile
200               Open SingleFile For Append As #EditFile
210           End If
    
220           If OpenSingle = False And Gather = False And ParseGather = False Then
                  Dim TmpFile
230               TmpFile = FreeFile
240               Open App.Path & "\ffxip_tmp_.tmp" For Output As #TmpFile
250           End If

    
260           RTB_Report.SelStart = Len(RTB_Report.Text)
270           lblStatus.Caption = "Errors: " & HasErrors & " - " & "Parsing Data...."
280           DoEvents
290           f = FreeFile

300           RTB.LoadFile Mid$(fileListBox.Text, 16)
310           RTB_Log.Text = "Loading File: " & Mid$(fileListBox.Text, 16) & vbNewLine & RTB_Log.Text
320           RTB.Text = Mid(RTB.Text, 101)
330           RTB.Text = Replace(RTB.Text, Chr(0), vbNewLine)
340           MyPos = InStrRev(fileListBox.Text, "\")
350           If Gather = False Then
360             CurrentFile = App.Path & "\FFXI_Logs" & Mid(fileListBox.Text, MyPos)
370           Else
380             CurrentFile = App.Path & "\FFXI_Gather" & Mid(fileListBox.Text, MyPos)
390           End If
400           RTB.SaveFile CurrentFile, rtfText
410           RTB_Log.Text = "Saving File: " & CurrentFile & vbNewLine & RTB_Log.Text

420           MyDate = Left$(fileListBox.Text, 5) & Format(Date, "/YYYY") & " " & Format(Format(Mid$(fileListBox.Text, 7, 6), "00:00:00"), "Hh:Nn:Ss AM/PM")
430           ResetTimeFile CurrentFile, MyDate
  
440           Erase FullFile
450           Index = 0
460           Open CurrentFile For Input As f
470             Do Until EOF(f)
480               Line Input #f, CurrentLine
      
490               LineType = Left(CurrentLine, 2)
500               If LineType = "ce" And ParserCommands = True Then
510                   If InStr(1, LCase(CurrentLine), "parser stop logging") Then
520                       StopLogging = True
530                   ElseIf InStr(1, LCase(CurrentLine), "parser start logging") Then
540                       StopLogging = False
550                   ElseIf InStr(1, LCase(CurrentLine), "parser gather ") Then
560                       MyPos = InStr(1, CurrentLine, "'")
570                       If MyPos <> 0 Then
580                           MyPos2 = InStr(MyPos + 1, CurrentLine, "'")
590                           If MyPos2 <> 0 Then
600                               SingleFile = Mid(CurrentLine, MyPos + 1, MyPos2 - (MyPos + 1))
                                  Dim FSO As FileSystemObject
610                               Set FSO = New FileSystemObject
                                  Dim SaveFile
620                               Set SaveFile = FSO.CreateTextFile(SingleFile, True)
630                               SaveFile.Close
640                               Set SaveFile = Nothing
650                               Set FSO = Nothing
660                           End If
670                       End If
680                   End If
690               End If
700               If StopLogging = False Then
710                   If Mid(CurrentLine, 51, 2) = "01" And Index <> 0 Then
720                       FullFile(Index - 1) = Left(FullFile(Index - 1), Len(FullFile(Index - 1)) - 3) & Mid(CurrentLine, 56) & " " & LineType
730                   Else
740                       ReDim Preserve FullFile(Index)
750                       FullFile(Index) = Mid(CurrentLine, 54) & " " & LineType
760                       Index = Index + 1
770                   End If
780               End If
790             Loop
800           Close #f
810           If Gather = False Or ParseGather = True Then
820             If Index <> 0 Then
830               ParseLog FullFile
840             End If
850             If ParseGather Then
860               If Index <> 0 Then
870                   For i = 0 To UBound(FullFile)
880                       Print #EditFile, FullFile(i)
890                   Next
900               End If
910             End If
920           Else
930               If Index <> 0 Then
940                   For i = 0 To UBound(FullFile)
950                       Print #EditFile, FullFile(i)
960                   Next
970               End If
980           End If
990           If Gather = False And ParseGather = False Then
1000              If Index <> 0 Then
1010                  For i = 0 To UBound(FullFile)
1020                      Print #TmpFile, FullFile(i)
1030                  Next
1040              End If
1050              Close #TmpFile
1060          End If
1070          lblStatus.Caption = "Errors: " & HasErrors & " - " & "Waiting for new log...."
1080          fileListBox.ListIndex = fileListBox.ListCount - 1
1090          LastItem = fileListBox.Text
1100          If optionResults(1).Value = True Then
1110              comboUser_Click
1120          End If
1130          If Gather = True Or ParseGather = True Then
1140              Close #EditFile
1150          End If
1160      End If
1170  End If
End Sub




