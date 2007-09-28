VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmRead 
   Caption         =   "FFXI Parser 3.6.0"
   ClientHeight    =   5760
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7350
   Icon            =   "frmReadnew.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5760
   ScaleWidth      =   7350
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox RTB_Fish 
      Height          =   4650
      Left            =   90
      TabIndex        =   36
      TabStop         =   0   'False
      ToolTipText     =   "Double Click to Save"
      Top             =   405
      Visible         =   0   'False
      Width           =   7170
      _ExtentX        =   12647
      _ExtentY        =   8202
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      RightMargin     =   1.00000e5
      TextRTF         =   $"frmReadnew.frx":1E72
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
   Begin VB.Frame frameChat 
      BorderStyle     =   0  'None
      Height          =   4650
      Left            =   90
      TabIndex        =   28
      Top             =   405
      Visible         =   0   'False
      Width           =   7215
      Begin VB.OptionButton optionChat 
         Caption         =   "All"
         Height          =   240
         Index           =   5
         Left            =   3825
         TabIndex        =   35
         Top             =   45
         Width           =   1005
      End
      Begin VB.OptionButton optionChat 
         Caption         =   "LinkShell"
         Height          =   240
         Index           =   4
         Left            =   2835
         TabIndex        =   34
         Top             =   45
         Width           =   1005
      End
      Begin VB.OptionButton optionChat 
         Caption         =   "Party"
         Height          =   240
         Index           =   3
         Left            =   2115
         TabIndex        =   33
         Top             =   45
         Width           =   690
      End
      Begin VB.OptionButton optionChat 
         Caption         =   "Tell"
         Height          =   240
         Index           =   2
         Left            =   1440
         TabIndex        =   32
         Top             =   45
         Width           =   690
      End
      Begin VB.OptionButton optionChat 
         Caption         =   "Shout"
         Height          =   240
         Index           =   1
         Left            =   630
         TabIndex        =   31
         Top             =   45
         Width           =   825
      End
      Begin VB.OptionButton optionChat 
         Caption         =   "Say"
         Height          =   240
         Index           =   0
         Left            =   0
         TabIndex        =   30
         Top             =   45
         Width           =   690
      End
      Begin RichTextLib.RichTextBox RTB_Chat 
         Height          =   4290
         Left            =   0
         TabIndex        =   29
         TabStop         =   0   'False
         ToolTipText     =   "Double Click to Save"
         Top             =   315
         Width           =   7170
         _ExtentX        =   12647
         _ExtentY        =   7567
         _Version        =   393217
         BorderStyle     =   0
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         RightMargin     =   1.00000e5
         TextRTF         =   $"frmReadnew.frx":1EE9
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
   End
   Begin VB.ComboBox comboUser 
      Height          =   315
      ItemData        =   "frmReadnew.frx":1F69
      Left            =   1800
      List            =   "frmReadnew.frx":1F7F
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   5085
      Width           =   1095
   End
   Begin VB.OptionButton optionResults 
      Height          =   195
      Index           =   0
      Left            =   90
      TabIndex        =   12
      ToolTipText     =   "Double click / Right click"
      Top             =   5130
      Value           =   -1  'True
      Width           =   240
   End
   Begin VB.ComboBox comboDisplay 
      Height          =   315
      ItemData        =   "frmReadnew.frx":1FBF
      Left            =   360
      List            =   "frmReadnew.frx":1FD8
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   5085
      Width           =   1095
   End
   Begin VB.OptionButton optionResults 
      Height          =   195
      Index           =   1
      Left            =   1530
      TabIndex        =   9
      Top             =   5145
      Width           =   240
   End
   Begin MSComDlg.CommonDialog CD_Save 
      Left            =   45
      Top             =   135
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   4455
      Top             =   405
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer timerRead 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   4995
      Top             =   450
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Start"
      Height          =   330
      Left            =   90
      TabIndex        =   0
      Top             =   45
      Width           =   1005
   End
   Begin VB.FileListBox fileList 
      Height          =   1455
      Left            =   90
      Pattern         =   "*.log"
      TabIndex        =   3
      Top             =   2385
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.ListBox fileListBox 
      Height          =   2595
      Left            =   90
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   5760
      Visible         =   0   'False
      Width           =   7170
   End
   Begin VB.DirListBox dirList 
      Height          =   1890
      Left            =   90
      TabIndex        =   1
      Top             =   405
      Visible         =   0   'False
      Width           =   2805
   End
   Begin VB.CommandButton cmdOptions 
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6390
      TabIndex        =   6
      Top             =   5085
      Width           =   870
   End
   Begin InetCtlsObjects.Inet inet 
      Left            =   6615
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   4
      URL             =   "http://"
   End
   Begin RichTextLib.RichTextBox RTB_Details 
      Height          =   4650
      Left            =   90
      TabIndex        =   14
      TabStop         =   0   'False
      ToolTipText     =   "Double Click to Save"
      Top             =   405
      Visible         =   0   'False
      Width           =   7170
      _ExtentX        =   12647
      _ExtentY        =   8202
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      RightMargin     =   1.00000e5
      TextRTF         =   $"frmReadnew.frx":2012
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
   Begin RichTextLib.RichTextBox RTB_Averages 
      Height          =   4650
      Left            =   90
      TabIndex        =   11
      TabStop         =   0   'False
      ToolTipText     =   "Double Click to Save"
      Top             =   405
      Visible         =   0   'False
      Width           =   7170
      _ExtentX        =   12647
      _ExtentY        =   8202
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      RightMargin     =   1.00000e5
      TextRTF         =   $"frmReadnew.frx":2089
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
   Begin RichTextLib.RichTextBox RTB_User 
      Height          =   4650
      Left            =   90
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Double Click to Save"
      Top             =   405
      Visible         =   0   'False
      Width           =   7170
      _ExtentX        =   12647
      _ExtentY        =   8202
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      RightMargin     =   1.00000e5
      TextRTF         =   $"frmReadnew.frx":2100
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
   Begin RichTextLib.RichTextBox RTB 
      Height          =   2400
      Left            =   405
      TabIndex        =   7
      Top             =   2025
      Visible         =   0   'False
      Width           =   6315
      _ExtentX        =   11139
      _ExtentY        =   4233
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmReadnew.frx":2180
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
      Left            =   90
      TabIndex        =   17
      Top             =   405
      Width           =   7215
      Begin VB.CommandButton cmdExport 
         Caption         =   "Export/Recalc"
         Enabled         =   0   'False
         Height          =   330
         Left            =   5130
         TabIndex        =   27
         Top             =   540
         Width           =   1995
      End
      Begin VB.CommandButton cmdCustom 
         Caption         =   "Select"
         Height          =   330
         Left            =   5130
         TabIndex        =   24
         Top             =   2160
         Width           =   1995
      End
      Begin VB.ComboBox comboMOB 
         Height          =   315
         ItemData        =   "frmReadnew.frx":21F7
         Left            =   5130
         List            =   "frmReadnew.frx":21F9
         TabIndex        =   23
         Text            =   "Type or select monster"
         Top             =   1800
         Width           =   1995
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "Select All"
         Height          =   330
         Left            =   5130
         TabIndex        =   21
         Top             =   1260
         Width           =   1995
      End
      Begin VB.CommandButton cmdUnselect 
         Caption         =   "Unselect All"
         Height          =   330
         Left            =   5130
         TabIndex        =   20
         Top             =   900
         Width           =   1995
      End
      Begin VB.CommandButton cmdRecalc 
         Caption         =   "Recalculate"
         Enabled         =   0   'False
         Height          =   330
         Left            =   5130
         TabIndex        =   19
         Top             =   180
         Width           =   1995
      End
      Begin VB.ListBox listResults 
         Height          =   3960
         Left            =   45
         MultiSelect     =   2  'Extended
         TabIndex        =   18
         Top             =   135
         Width           =   5010
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         Caption         =   "Example: Type ""goblin"" without quotes for all goblins."
         ForeColor       =   &H00C00000&
         Height          =   600
         Index           =   1
         Left            =   5175
         TabIndex        =   25
         Top             =   2520
         Width           =   1905
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         Caption         =   "All log information except battles will be lost if you recalculate."
         ForeColor       =   &H000000C0&
         Height          =   600
         Index           =   0
         Left            =   5175
         TabIndex        =   22
         Top             =   3510
         Width           =   1905
      End
   End
   Begin RichTextLib.RichTextBox RTB_Report 
      Height          =   4560
      Left            =   90
      TabIndex        =   37
      TabStop         =   0   'False
      ToolTipText     =   "Double Click to Save"
      Top             =   405
      Width           =   7170
      _ExtentX        =   12647
      _ExtentY        =   8043
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      RightMargin     =   1.00000e5
      TextRTF         =   $"frmReadnew.frx":21FB
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
   Begin VB.Label lblAbout 
      Alignment       =   1  'Right Justify
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   5805
      MouseIcon       =   "frmReadnew.frx":2272
      MousePointer    =   99  'Custom
      TabIndex        =   26
      Top             =   5145
      Width           =   510
   End
   Begin VB.Label lblChange 
      Caption         =   "Change"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   2925
      MouseIcon       =   "frmReadnew.frx":23C4
      MousePointer    =   99  'Custom
      TabIndex        =   16
      Top             =   5145
      Width           =   510
   End
   Begin VB.Label lblUpdate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " No update available."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   90
      TabIndex        =   13
      Top             =   5490
      Visible         =   0   'False
      Width           =   7170
   End
   Begin VB.Label lblStatus 
      Alignment       =   1  'Right Justify
      Caption         =   "Waiting."
      Height          =   195
      Left            =   1755
      TabIndex        =   5
      Top             =   105
      Width           =   5460
   End
   Begin VB.Label Label1 
      Caption         =   "Status:"
      Height          =   195
      Left            =   1170
      TabIndex        =   4
      Top             =   100
      Width           =   690
   End
   Begin VB.Shape Shape1 
      Height          =   330
      Left            =   1125
      Top             =   45
      Width           =   6180
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "mnuOptions"
      Visible         =   0   'False
      Begin VB.Menu mnuPlayer 
         Caption         =   "Player Damage Only"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMonster 
         Caption         =   "Monster Damage Only"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSpacer1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuReport 
         Caption         =   "Report Fields"
      End
      Begin VB.Menu mnuUser 
         Caption         =   "Change User"
      End
      Begin VB.Menu mnuClear 
         Caption         =   "Clear"
      End
      Begin VB.Menu mnuSpacer2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUpdate 
         Caption         =   "Auto Update Check"
      End
      Begin VB.Menu mnuTray 
         Caption         =   "Show as Tray Icon"
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
Dim DPS(17, 2) As String, Stats(50, 30) As String, Players(50, 4) As String, P1Uses As String, HasErrors As Boolean, FishFound() As String, LootFound() As String, PlayerLoot() As String, CurrentFight As String
Dim SingleUser As String, SingleAcc As Long, SingleLines As Long, SingleHit As Long, SingleMiss As Long, SingleDmg As String, SingleCrit As String, LastLoc As Long, PrevTotalDMG As Long, PrevSwings As Long, PrevTotalTaken As Long, PrevTakenSwings As Long, TotalSwingTaken As Long
Dim SingleUserB As String, SingleAccB As Long, SingleLinesB As Long, SingleHitB As Long, SingleMissB As Long, SingleDmgB As String, SingleCritB As String, LastLocB As Long
Dim ChatText() As String, EffList As String

Dim ReadDPS_Start As Boolean, ReadEXP_Start As Boolean, StopEXP As Boolean, BeginDPS As Boolean
Dim ReadDPS_Stop As Boolean, ReadEXP_Stop As Boolean, StopDPS As Boolean
Dim Read_Start As Boolean, Read_Stop As Boolean

Dim ReadFISH_Start As Boolean, FishHeader As String, FishComment As String

Dim GrandPList(17, 28) As String, MB As Boolean
Dim PList(17, 28) As String, TotalExp As Long, TotalDPS As Long, StartTime As Date, StopTime As Date, StartTimeDPS As Date, StopTimeDPS As Date, FightStartTime As Date, FightStopTime As Date
Dim Prev0 As String, Prev1 As String, Prev2 As String, Prev3 As String, Prev4 As String, Prev5 As String, Prev6 As String, Prev7 As String, Prev8 As String
Dim FoundP1 As Boolean, SkipIt As Boolean
Dim dHigh As Long, dLow As Long, SelStart As Long, ErrorCount As Long, UniqueMOB As Long
Dim ff As Integer, i As Integer, X As Integer, p As Integer, pl As Integer, u As Integer, f As Integer
Dim MyPos As Integer, MyPos2 As Integer, MyPos3 As Integer
Dim CurrentLine As String, NextLine As String, P1Special As String, P1 As String, P1Opp As String, P1Stat As String, P1Takes As String, PartA As String, FightComment As String, SaveFileName As String, NewPlayer As String, MonsterCheck As String
Dim PrevlineA As String, PrevlineB As String, PrevlineC As String, PrevlineD As String, PrevlineE As String, PrevlineF As String
Dim HTMLCode As String, SummaryCode As String


Dim Critical As Boolean, Casts As Boolean
Dim LastItem As String, Updates As String


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

Const SW_SHOWNORMAL = 1

Option Explicit


           
Private Function ColumnText() As String
ColumnText = ResizePart("Player", 1000)
ColumnText = ColumnText & vbTab & ResizePart("Total", 525)
If ReportOptions(18, 0) = 1 Then
    ColumnText = ColumnText & vbTab & ResizePart("DMG %", 525)
End If
If ReportOptions(0, 0) = 1 Then
    ColumnText = ColumnText & vbTab & ResizePart("Basic", 525)
End If
If ReportOptions(1, 0) = 1 Then
    ColumnText = ColumnText & vbTab & ResizePart("Skill", 525)
End If
If ReportOptions(2, 0) = 1 Then
    ColumnText = ColumnText & vbTab & ResizePart("Spell", 525)
End If
If ReportOptions(3, 0) = 1 Then
    ColumnText = ColumnText & vbTab & ResizePart("Hi/Lo", 525)
End If
If ReportOptions(4, 0) = 1 Then
    ColumnText = ColumnText & vbTab & ResizePart("Avg", 525)
End If
If ReportOptions(5, 0) = 1 Then
    ColumnText = ColumnText & vbTab & ResizePart("Crit %", 525)
End If
If ReportOptions(6, 0) = 1 Then
    ColumnText = ColumnText & vbTab & ResizePart("Crit #", 525)
End If
If ReportOptions(21, 0) = 1 Then
    ColumnText = ColumnText & vbTab & ResizePart("Effect", 525)
End If
If ReportOptions(22, 0) = 1 Then
    ColumnText = ColumnText & vbTab & ResizePart("WS #", 525)
End If
If ReportOptions(7, 0) = 1 Then
    ColumnText = ColumnText & vbTab & ResizePart("Hit %", 525)
End If
If ReportOptions(8, 0) = 1 Then
    ColumnText = ColumnText & vbTab & ResizePart("Hit/Miss", 525)
End If
If ReportOptions(9, 0) = 1 Then
    ColumnText = ColumnText & vbTab & ResizePart("Avd %", 525)
End If
If ReportOptions(10, 0) = 1 Then
    ColumnText = ColumnText & vbTab & ResizePart("Tk/Av", 525)
End If
If ReportOptions(11, 0) = 1 Then
    ColumnText = ColumnText & vbTab & ResizePart("Evade", 525)
End If
If ReportOptions(12, 0) = 1 Then
    ColumnText = ColumnText & vbTab & ResizePart("Parry", 525)
End If
If ReportOptions(13, 0) = 1 Then
    ColumnText = ColumnText & vbTab & ResizePart("Block", 525)
End If
If ReportOptions(14, 0) = 1 Then
    ColumnText = ColumnText & vbTab & ResizePart("Absorb", 525)
End If
If ReportOptions(15, 0) = 1 Then
    ColumnText = ColumnText & vbTab & ResizePart("Avoid", 525)
End If
If ReportOptions(20, 0) = 1 Then
    ColumnText = ColumnText & vbTab & ResizePart("Anti", 525)
End If
If ReportOptions(23, 0) = 1 Then
    ColumnText = ColumnText & vbTab & ResizePart("Cnter", 525)
End If
If ReportOptions(16, 0) = 1 Then
    ColumnText = ColumnText & vbTab & ResizePart("Taken", 525)
End If
If ReportOptions(17, 0) = 1 Then
    ColumnText = ColumnText & vbTab & ResizePart("Rcver", 525)
End If
If ReportOptions(19, 0) = 1 Then
    ColumnText = ColumnText & vbTab & ResizePart("Heal", 525)
End If
ColumnText = ColumnText & vbTab
End Function

Private Sub FishRPT()
Dim lf, MyPos, AddLoot As String
If FishFound(0) <> "" Then
    If FishHeader = "" Then FishHeader = "Unknown Time"
    RTB_Fish.SelBold = True
    RTB_Fish.SelText = FishHeader & vbNewLine
    RTB_Fish.SelBold = False
    For lf = 0 To UBound(FishFound)
        RTB_Fish.SelBold = False
        If FishFound(lf) <> "" Then
'            If InStr(1, LCase(FishFound(lf)), " - catches lost") = 0 And InStr(1, LCase(FishFound(lf)), " - didn't catch") = 0 Then
'                AddLoot = Replace(Replace(FishFound(lf), "a ", ""), "an ", "") & "s"
'                MyPos = InStr(1, AddLoot, " - ")
'                AddLoot = Left$(AddLoot, MyPos - 1) & " - " & UCase(Mid$(AddLoot, MyPos + 3, 1)) & Mid$(AddLoot, MyPos + 4)
'                RTB_Fish.SelBold = False
'                RTB_Fish.SelColor = vbBlack
'                RTB_Fish.SelText = vbTab & AddLoot & vbNewLine
'            Else
                AddLoot = FishFound(lf)
                MyPos = InStr(1, AddLoot, " - ")
                AddLoot = Left$(AddLoot, MyPos - 1) & " - " & UCase(Mid$(AddLoot, MyPos + 3, 1)) & Mid$(AddLoot, MyPos + 4)
                RTB_Fish.SelBold = False
                RTB_Fish.SelColor = vbBlack
                RTB_Fish.SelText = vbTab & AddLoot & vbNewLine
 '           End If
        End If
    Next
    If FishComment <> "" Then
        RTB_Fish.SelBold = False
        RTB_Fish.SelColor = vbBlack
        RTB_Fish.SelText = vbTab & "Comment: " & FishComment
        FishComment = ""
    End If
    RTB_Fish.SelText = vbNewLine & vbNewLine
End If
End Sub


Private Function HTMLCodeHeader() As String
HTMLCodeHeader = ""
HTMLCodeHeader = HTMLCodeHeader & "<TR style=""FONT-WEIGHT:bold;BACKGROUND-COLOR:#b8ced9"">"
HTMLCodeHeader = HTMLCodeHeader & "<TD WIDTH=75></TD>"
If ExportOptions(0, 0) = 1 Then HTMLCodeHeader = HTMLCodeHeader & "<TD WIDTH=50>Basic</TD>"
If ExportOptions(1, 0) = 1 Then HTMLCodeHeader = HTMLCodeHeader & "<TD WIDTH=50>Skill</TD>"
If ExportOptions(2, 0) = 1 Then HTMLCodeHeader = HTMLCodeHeader & "<TD WIDTH=50>Spell</TD>"
If ExportOptions(21, 0) = 1 Then HTMLCodeHeader = HTMLCodeHeader & "<TD WIDTH=50>Effects</TD>"
If ExportOptions(22, 0) = 1 Then HTMLCodeHeader = HTMLCodeHeader & "<TD WIDTH=50>WS #</TD>"
If ExportOptions(3, 0) = 1 Then HTMLCodeHeader = HTMLCodeHeader & "<TD WIDTH=50>High/Low</TD>"
If ExportOptions(4, 0) = 1 Then HTMLCodeHeader = HTMLCodeHeader & "<TD WIDTH=50>Average</TD>"
If ExportOptions(5, 0) = 1 Then HTMLCodeHeader = HTMLCodeHeader & "<TD WIDTH=50>Crit %</TD>"
If ExportOptions(6, 0) = 1 Then HTMLCodeHeader = HTMLCodeHeader & "<TD WIDTH=50>Crits</TD>"
If ExportOptions(7, 0) = 1 Then HTMLCodeHeader = HTMLCodeHeader & "<TD WIDTH=50>Hit %</TD>"
If ExportOptions(8, 0) = 1 Then HTMLCodeHeader = HTMLCodeHeader & "<TD WIDTH=50>Hit/Miss</TD>"
If ExportOptions(9, 0) = 1 Then HTMLCodeHeader = HTMLCodeHeader & "<TD WIDTH=50>Avoid %</TD>"
If ExportOptions(10, 0) = 1 Then HTMLCodeHeader = HTMLCodeHeader & "<TD WIDTH=50>Take/Avoid</TD>"
If ExportOptions(11, 0) = 1 Then HTMLCodeHeader = HTMLCodeHeader & "<TD WIDTH=50>Evades</TD>"
If ExportOptions(12, 0) = 1 Then HTMLCodeHeader = HTMLCodeHeader & "<TD WIDTH=50>Parries</TD>"
If ExportOptions(13, 0) = 1 Then HTMLCodeHeader = HTMLCodeHeader & "<TD WIDTH=50>Blocks</TD>"
If ExportOptions(23, 0) = 1 Then HTMLCodeHeader = HTMLCodeHeader & "<TD WIDTH=50>Counters</TD>"
If ExportOptions(20, 0) = 1 Then HTMLCodeHeader = HTMLCodeHeader & "<TD WIDTH=50>Anticipates</TD>"
If ExportOptions(14, 0) = 1 Then HTMLCodeHeader = HTMLCodeHeader & "<TD WIDTH=50>Absorbs</TD>"
If ExportOptions(15, 0) = 1 Then HTMLCodeHeader = HTMLCodeHeader & "<TD WIDTH=50>Avoids</TD>"
If ExportOptions(16, 0) = 1 Then HTMLCodeHeader = HTMLCodeHeader & "<TD WIDTH=75>DMG Taken</TD>"
If ExportOptions(17, 0) = 1 Then HTMLCodeHeader = HTMLCodeHeader & "<TD WIDTH=75>HP Rec'd</TD>"
If ExportOptions(19, 0) = 1 Then HTMLCodeHeader = HTMLCodeHeader & "<TD WIDTH=75>HP Healed</TD>"
HTMLCodeHeader = HTMLCodeHeader & "<TD WIDTH=75>TTL DMG</TD></TR>" & vbNewLine
End Function

Private Sub NewStatsArray(Location As Integer, Amt As Long)
Dim i
For i = 0 To Amt
    If i <> 10 And i <> 11 And i <> 12 And i <> 13 Then
        Stats(Location, i) = "0"
    End If
Next
End Sub

Public Sub ParseLog(FullFile() As String, CustomMode As Boolean, GenerateHTML As Boolean)
On Error GoTo err_handler
Dim ExpLine As Long, ExpGain As String, LineType As String, ExpChecks As Long, AvgAcc As String, CustomAdd As Boolean, PrevLoot As Boolean, PrevFish As Boolean, StatClear As Integer, TotalSeconds As String
Dim AddMOB As Boolean, lf, LootItem As String, FishItem As String, GilAmt As Long, TotalBase As Long, TotalSkill As Long, TotalSpell As Long, TotalTaken As Long, TotalHP As Long, TotalHPH As Long, TotalDMG As Long, TotalHit As Long, TotalSwing As Long, PrevUseType As String, PreP1 As String, TotalEffect As Long, HTMLCodeNew As String, LootPlayer As String, dp, EstDPS As String
Dim Part(25) As String, Parts As Integer, CurrentSel As String, LastPlayer As Boolean

If CustomMode = False Then
    Dim EditFile
    EditFile = FreeFile
    Open App.Path & "\EditFile.log" For Append As #EditFile
End If

    
For ff = 0 To UBound(FullFile)
    If Len(CurrentLine) > 10 Then
        PrevlineF = PrevlineE
        PrevlineE = PrevlineD
        PrevlineD = PrevlineC
        PrevlineC = PrevlineB
        PrevlineB = PrevlineA
        PrevlineA = CurrentLine
    End If
    CurrentLine = FullFile(ff)
    If ff + 1 <= UBound(FullFile) Then
        NextLine = FullFile(ff + 1)
    Else
        NextLine = ""
    End If
    
    LineType = Trim(Right(CurrentLine, 3))
    If Left$(NextLine, 2) <> "" And IsNumeric(Left$(NextLine, 1)) = False And Len(NextLine) <> 2 And Trim(NextLine) <> "" Then
        CurrentLine = CurrentLine & ", " & NextLine
        If ff + 1 <= UBound(FullFile) Then
            ff = ff + 1
        End If
    End If
    'MISS/BLOCK/PARRY/ABSORB/EVADE/ANTI
    If (InStr(1, CurrentLine, " anticipates the attack.") <> 0 Or InStr(1, CurrentLine, " evades.") <> 0 Or InStr(1, CurrentLine, " absorbs the damage and ") <> 0 Or InStr(1, CurrentLine, " blocks ") <> 0 Or InStr(1, CurrentLine, " parries ") <> 0 Or InStr(1, CurrentLine, " misses ") <> 0 Or InStr(1, CurrentLine, " miss ") <> 0 Or InStr(1, CurrentLine, " misses.") <> 0) And InStr(1, "09,0a,01,02,0c,04,0d,05,0e,06,0f", LineType) = 0 Then
        P1Special = ""
        FoundP1 = False
        MyPos = InStr(1, CurrentLine, " misses ")
        If InStr(1, CurrentLine, " misses ") Then
            MyPos = InStr(1, CurrentLine, " misses ")
        ElseIf InStr(1, CurrentLine, " parries ") Then
            MyPos = InStr(1, CurrentLine, " parries ")
        ElseIf InStr(1, CurrentLine, " blocks ") Then
            MyPos = InStr(1, CurrentLine, " blocks ")
        ElseIf InStr(1, CurrentLine, " absorbs ") Then
            MyPos = InStr(1, CurrentLine, " shadow absorbs ")
        ElseIf InStr(1, CurrentLine, " anticipates ") Then
            MyPos = InStr(1, CurrentLine, " anticipates ")
        ElseIf InStr(1, CurrentLine, " ranged attack misses.") Then
            MyPos = InStr(1, CurrentLine, " ranged attack misses.")
        ElseIf InStr(1, CurrentLine, " evades.") Then
            MyPos = InStr(1, CurrentLine, " evades.")
        ElseIf MyPos = 0 Then
            MyPos = InStr(1, CurrentLine, " miss ")
        End If
        
        If InStr(1, CurrentLine, " uses ") Then
            MyPos = InStr(1, CurrentLine, " uses ")
        End If
        P1 = Mid$(CurrentLine, 3, MyPos - 3)
        If InStr(1, CurrentLine, " uses ") Then P1Uses = P1
        If InStr(1, CurrentLine, " misses ") Then
            MyPos = MyPos + 8
        ElseIf InStr(1, CurrentLine, " miss ") Then
            MyPos = MyPos + 6
        ElseIf InStr(1, CurrentLine, " parries ") Then
            MyPos = MyPos + 9
        ElseIf InStr(1, CurrentLine, " blocks ") Then
            MyPos = MyPos + 8
        ElseIf InStr(1, CurrentLine, " shadow absorbs ") Then
            MyPos = MyPos + 16
        ElseIf InStr(1, CurrentLine, " anticipates ") Then
            MyPos = MyPos + 13
        ElseIf InStr(1, CurrentLine, " ranged attack misses.") Then
            MyPos = MyPos + 22
        ElseIf InStr(1, CurrentLine, " evades.") Then
            MyPos = MyPos + 9
        End If
        P1 = Replace(P1, "'s", "")
        If InStr(1, CurrentLine, " misses.") = 0 And InStr(1, CurrentLine, ", but") = 0 And InStr(1, CurrentLine, " parries ") = 0 And InStr(1, CurrentLine, " blocks ") = 0 And InStr(1, CurrentLine, " absorbs the damage and ") = 0 And InStr(1, CurrentLine, " evades.") = 0 And InStr(1, CurrentLine, " anticipates the attack.") = 0 Then
            MyPos2 = InStr(1, CurrentLine, ".")
            P1Opp = Mid$(CurrentLine, MyPos, MyPos2 - MyPos)
        ElseIf InStr(1, CurrentLine, " misses.") = 0 And InStr(1, CurrentLine, " parries ") = 0 And InStr(1, CurrentLine, " blocks ") = 0 And InStr(1, CurrentLine, " absorbs the damage and ") = 0 And InStr(1, CurrentLine, " evades.") = 0 And InStr(1, CurrentLine, " anticipates the attack.") = 0 Then
            MyPos = InStr(1, CurrentLine, ", but misses ")
            MyPos2 = InStr(1, CurrentLine, ".")
            P1Opp = Replace(Mid$(CurrentLine, MyPos + 13, MyPos2 - (MyPos + 13)), "the ", "The ")
        ElseIf InStr(1, CurrentLine, " parries ") Or InStr(1, CurrentLine, " blocks ") Then
            MyPos2 = InStr(1, CurrentLine, "attack ")
            P1Opp = Mid$(CurrentLine, MyPos, MyPos2 - MyPos)
        ElseIf Mid$(CurrentLine, 3, 4) <> "The " Then
            P1Opp = CurrentFight
        Else
            P1Opp = ""
        End If
        P1Opp = Replace(P1Opp, "the ", "The ")
        If InStr(1, P1, "'s") Then P1 = Trim(Replace(P1, "'s", ""))
        If InStr(1, P1Opp, "'s") Then P1Opp = Trim(Replace(P1Opp, "'s", ""))
        If InStr(1, CurrentLine, " absorbs the damage and ") Or InStr(1, CurrentLine, " evades.") Or InStr(1, CurrentLine, " blocks ") Or InStr(1, CurrentLine, " parries ") Or InStr(1, CurrentLine, " anticipates the attack.") Then
            PreP1 = P1
            P1 = P1Opp
            P1Opp = PreP1
        End If
        If P1Opp <> "" And P1 <> "" Then
            For i = 0 To UBound(Stats)
                If LCase(Stats(i, 0)) = LCase(P1) And LCase(Stats(i, 1)) = LCase(P1Opp) Then
                    If Stats(i, 3) = "" Then Stats(i, 3) = "0"
                    Stats(i, 3) = CDbl(Stats(i, 3)) + 1
                    FoundP1 = True
                End If
            Next
            If FoundP1 = False Then
                For i = 0 To UBound(Stats)
                    If Stats(i, 0) = "" Then
                        NewStatsArray i, 30
                        Stats(i, 0) = P1
                        Stats(i, 1) = P1Opp
                        Stats(i, 3) = CDbl(Stats(i, 3)) + 1
                        Exit For
                    End If
                Next
            End If
            CurrentLine = CurrentLine
            FoundP1 = False
            For i = 0 To UBound(Stats)
                If LCase(Stats(i, 0)) = LCase(P1Opp) And LCase(Stats(i, 1)) = LCase(P1) Then
                    Stats(i, 15) = CDbl(Stats(i, 15)) + 1
                    If InStr(1, CurrentLine, " evades.") Then
                        Stats(i, 21) = CDbl(Stats(i, 21)) + 1
                    ElseIf InStr(1, CurrentLine, " parries ") Then
                        Stats(i, 22) = CDbl(Stats(i, 22)) + 1
                    ElseIf InStr(1, CurrentLine, " blocks ") Then
                        Stats(i, 23) = CDbl(Stats(i, 23)) + 1
                    ElseIf InStr(1, CurrentLine, " absorbs the damage and ") Then
                        Stats(i, 24) = CDbl(Stats(i, 24)) + 1
                    ElseIf InStr(1, CurrentLine, " miss") Then
                        Stats(i, 25) = CDbl(Stats(i, 25)) + 1
                    ElseIf InStr(1, CurrentLine, " anticipates the attack.") Then
                        Stats(i, 27) = CDbl(Stats(i, 27)) + 1
                    End If
                    FoundP1 = True
                    Exit For
                End If
            Next
            If FoundP1 = False Then
                For i = 0 To UBound(Stats)
                    If Stats(i, 0) = "" Then
                        NewStatsArray i, 30
                        Stats(i, 0) = P1Opp
                        Stats(i, 1) = P1
                        Stats(i, 15) = CDbl(Stats(i, 15)) + 1
                        If InStr(1, CurrentLine, " evades.") Then
                            Stats(i, 21) = CDbl(Stats(i, 21)) + 1
                        ElseIf InStr(1, CurrentLine, " parries ") Then
                            Stats(i, 22) = CDbl(Stats(i, 22)) + 1
                        ElseIf InStr(1, CurrentLine, " blocks ") Then
                            Stats(i, 23) = CDbl(Stats(i, 23)) + 1
                        ElseIf InStr(1, CurrentLine, " absorbs the damage and ") Then
                            Stats(i, 24) = CDbl(Stats(i, 24)) + 1
                        ElseIf InStr(1, CurrentLine, " miss") Then
                            Stats(i, 25) = CDbl(Stats(i, 25)) + 1
                        ElseIf InStr(1, CurrentLine, " anticipates the attack.") Then
                            Stats(i, 27) = CDbl(Stats(i, 27)) + 1
                        End If
                        Exit For
                    End If
                Next
            End If
            FoundP1 = False
            If CustomMode = False Then
                Print #EditFile, CurrentLine
            End If
        End If
    'BASIC/ADDITIONAL/COUTNER
    ElseIf (InStr(1, CurrentLine, " counter") <> 0 Or InStr(1, CurrentLine, "Additional effect: ") <> 0 Or InStr(1, CurrentLine, " hits ") <> 0 Or InStr(1, CurrentLine, " hit ") <> 0) And InStr(1, LCase(CurrentLine), " drained ") = 0 And InStr(1, "09,0a,01,02,0c,04,0d,05,0e,06,0f", LineType) = 0 Then
        P1Special = ""
        FoundP1 = False
        MyPos = InStr(1, CurrentLine, " hits ")
        If MyPos = 0 Then MyPos = InStr(1, CurrentLine, " hit ")
        If InStr(1, LCase(CurrentLine), "ranged attack") Then MyPos = MyPos - 16
        If InStr(1, LCase(CurrentLine), "counter") Then MyPos = InStr(1, CurrentLine, "'s attack")
        If InStr(1, CurrentLine, "Additional effect: ") = 0 Then P1 = Mid$(CurrentLine, 3, MyPos - 3)
        If InStr(1, LCase(CurrentLine), "ranged attack") Then MyPos = MyPos + 16
        P1 = Replace(P1, "Cover! ", "")
        If InStr(1, CurrentLine, " hit ") Then MyPos = MyPos - 1
        MyPos2 = InStr(MyPos + 7, CurrentLine, " for ")
        PrevlineA = PrevlineA
        If InStr(1, CurrentLine, "Additional effect: ") Then
            MyPos = InStr(1, CurrentLine, "Additional effect: ") + 19
            MyPos2 = InStr(MyPos, CurrentLine, " takes ")
            If MyPos2 = 0 Then
                P1Opp = CurrentFight
            Else
                P1Opp = Mid$(CurrentLine, MyPos, MyPos2 - (MyPos))
            End If
            MyPos3 = InStr(1, CurrentLine, " additional points of ")
            If MyPos3 = 0 Then MyPos3 = InStr(1, CurrentLine, " additional point of ")
            If MyPos2 <> 0 Then
                P1Stat = Mid$(CurrentLine, MyPos2 + 6, MyPos3 - (MyPos2 + 6))
            Else
                MyPos2 = InStr(MyPos + 1, CurrentLine, " ")
                P1Stat = Mid$(CurrentLine, MyPos, MyPos2 - (MyPos))
            End If
        ElseIf InStr(1, LCase(CurrentLine), "counter") Then
            P1Opp = P1
            MyPos = InStr(1, CurrentLine, " by ")
            MyPos2 = InStr(MyPos, CurrentLine, ".")
            P1 = Mid$(CurrentLine, MyPos + 4, MyPos2 - (MyPos + 4))
            MyPos2 = InStr(1, CurrentLine, " takes ")
            MyPos3 = InStr(1, CurrentLine, " points of ")
            P1Stat = Mid$(CurrentLine, MyPos2 + 7, MyPos3 - (MyPos2 + 7))
        Else
            P1Opp = Mid$(CurrentLine, MyPos + 6, MyPos2 - (MyPos + 6))
            MyPos3 = InStr(1, CurrentLine, " points of ")
            If MyPos3 = 0 Then MyPos3 = InStr(1, CurrentLine, " point of ")
            P1Stat = Mid$(CurrentLine, MyPos2 + 5, MyPos3 - (MyPos2 + 5))
        End If


        If InStr(1, P1, "'s") Then P1 = Replace(P1, "'s", "")
        P1Opp = Replace(P1Opp, "the ", "The ")
        If InStr(1, "14,19", Trim(Right$(CurrentLine, 3))) Then
            CurrentFight = P1Opp
        ElseIf InStr(1, "1c,20", Trim(Right$(CurrentLine, 3))) Then
            CurrentFight = P1
        End If
        If IsNumeric(Trim(P1Stat)) Then
            If StopDPS = False And BeginDPS = True Then
                If LineType <> "20" And LineType <> "1c" Then
                    For i = 0 To UBound(DPS)
                        If LCase(DPS(i, 0)) = LCase(P1) Then
                            DPS(i, 1) = CDbl(DPS(i, 1)) + CDbl(P1Stat)
                            FoundP1 = True
                        End If
                    Next
                    If FoundP1 = False Then
                        For i = 0 To UBound(DPS)
                            If DPS(i, 0) = "" Then
                                DPS(i, 0) = P1
                                DPS(i, 1) = CDbl(P1Stat)
                                Exit For
                            End If
                        Next
                    End If
                    FoundP1 = False
                End If
            End If
            
            For i = 0 To UBound(Stats)
                If LCase(Stats(i, 0)) = LCase(P1) And LCase(Stats(i, 1)) = LCase(P1Opp) Then
                    If InStr(1, CurrentLine, "Additional effect: ") = 0 Then
                        Stats(i, 2) = CDbl(Stats(i, 2)) + CDbl(P1Stat)
                    Else
                        Stats(i, 28) = CDbl(Stats(i, 28)) + CDbl(P1Stat)
                    End If
                    If InStr(1, CurrentLine, "Additional effect: ") = 0 Then
                        Stats(i, 4) = CDbl(Stats(i, 4)) + 1
                    End If
                    If InStr(1, LCase(CurrentLine), "counter") Then
                        Stats(i, 15) = CDbl(Stats(i, 15)) + 1
                    End If
                    Stats(i, 9) = CDbl(Stats(i, 9)) + CDbl(P1Stat)
                    If InStr(1, LCase(CurrentLine), "ranged attack") Then
                      If Stats(i, 13) = "" Then
                        Stats(i, 13) = P1Stat
                      Else
                        Stats(i, 13) = Stats(i, 13) & ", " & P1Stat
                      End If
                    Else
                      If Stats(i, 10) = "" Then
                        Stats(i, 10) = P1Stat
                      Else
                        Stats(i, 10) = Stats(i, 10) & ", " & P1Stat
                      End If
                    End If
                    If P1Stat > CDbl(Stats(i, 18)) Then
                        Stats(i, 18) = P1Stat
                    End If
                    If P1Stat < CDbl(Stats(i, 19)) Or CDbl(Stats(i, 19)) = 0 Then
                        Stats(i, 19) = P1Stat
                    End If
                    If InStr(1, CurrentLine, " counter") Then
                        Stats(i, 30) = CDbl(Stats(i, 30)) + 1
                    End If
                    FoundP1 = True
                    Exit For
                End If
            Next
            If FoundP1 = False Then
                For i = 0 To UBound(Stats)
                    If Stats(i, 0) = "" Then
                        NewStatsArray i, 30
                        Stats(i, 0) = P1
                        Stats(i, 1) = P1Opp
                        Stats(i, 18) = P1Stat
                        Stats(i, 19) = P1Stat
                        If InStr(1, CurrentLine, "Additional effect: ") = 0 Then
                            Stats(i, 2) = CDbl(Stats(i, 2)) + CDbl(P1Stat)
                        Else
                            Stats(i, 28) = CDbl(Stats(i, 28)) + CDbl(P1Stat)
                        End If
                        If InStr(1, CurrentLine, "Additional effect: ") = 0 Then
                            Stats(i, 4) = "1"
                        End If
                        If InStr(1, CurrentLine, " counter") Then
                            Stats(i, 30) = "1"
                            Stats(i, 15) = "1"
                        End If
                        Stats(i, 9) = CDbl(Stats(i, 9)) + CDbl(P1Stat)
                        If InStr(1, LCase(CurrentLine), "ranged attack") Then
                          Stats(i, 13) = P1Stat
                        Else
                          Stats(i, 10) = P1Stat
                        End If
                        Exit For
                    End If
                Next
            End If
            FoundP1 = False
            For i = 0 To UBound(Stats)
                If LCase(Stats(i, 0)) = LCase(P1Opp) And LCase(Stats(i, 1)) = LCase(P1) Then
                    Stats(i, 16) = CDbl(Stats(i, 16)) + 1
                    Stats(i, 17) = CDbl(Stats(i, 17)) + CDbl(P1Stat)
                    FoundP1 = True
                    Exit For
                End If
            Next
            If FoundP1 = False Then
                For i = 0 To UBound(Stats)
                    If Stats(i, 0) = "" Then
                        NewStatsArray i, 30
                        Stats(i, 0) = P1Opp
                        Stats(i, 1) = P1
                        Stats(i, 16) = "1"
                        Stats(i, 17) = P1Stat
                        Exit For
                    End If
                Next
            End If
            FoundP1 = False
            If CustomMode = False Then
                Print #EditFile, CurrentLine
            End If
        End If
    'SPECIAL/CRIT/SPELL USER
    ElseIf (InStr(1, LCase(CurrentLine), " uses ") <> 0 Or InStr(1, LCase(CurrentLine), "s use ") <> 0 Or InStr(1, LCase(CurrentLine), "critical hit!") <> 0 Or InStr(1, LCase(CurrentLine), "skillchain: ") <> 0 Or InStr(1, LCase(CurrentLine), " casts ") <> 0) And InStr(1, "09,0a,01,02,0c,04,0d,05,0e,06,0f", LineType) = 0 Then
        PrevUseType = LineType
        If InStr(1, LCase(CurrentLine), "critical hit!") Then
            Critical = True
        Else
            Critical = False
        End If

        If InStr(1, LCase(CurrentLine), " casts ") Then
            Casts = True
        Else
            Casts = False
        End If
        MyPos = InStr(3, CurrentLine, " uses ")
        If InStr(1, LCase(CurrentLine), "s use ") Then
            MyPos = InStr(3, CurrentLine, "s use ")
            P1Uses = Mid$(CurrentLine, 3, MyPos - 2)
        ElseIf InStr(1, LCase(CurrentLine), " uses ") Then
            MyPos = InStr(3, CurrentLine, " uses ")
            P1Uses = Mid$(CurrentLine, 3, MyPos - 3)
            MyPos = InStr(1, CurrentLine, ".")
            If MyPos = 0 Then
              MyPos = InStr(1, CurrentLine, "!")
            End If
            If MyPos = 0 Then
              MyPos = InStrRev(CurrentLine, " ")
            End If
            P1Special = Mid$(CurrentLine, InStr(1, CurrentLine, " uses ") + 6, MyPos - (InStr(1, CurrentLine, " uses ") + 6))
        ElseIf InStr(1, LCase(CurrentLine), " casts ") Then
            MyPos = InStr(3, CurrentLine, " casts ")
            P1Uses = Mid$(CurrentLine, 3, MyPos - 3)
            MyPos = InStr(1, CurrentLine, ".")
            If MyPos = 0 Then
              MyPos = InStr(1, CurrentLine, "!")
            End If
            P1Special = Mid$(CurrentLine, InStr(1, CurrentLine, " casts ") + 7, MyPos - (InStr(1, CurrentLine, " casts ") + 7))
        ElseIf InStr(1, LCase(CurrentLine), "skillchain: ") Then
            MyPos = InStr(3, CurrentLine, ".")
            P1Uses = "Skillchain: " & Mid$(CurrentLine, 15, MyPos - 15)
        Else
            MyPos = InStr(3, CurrentLine, " score")
            P1Uses = Mid$(CurrentLine, 3, MyPos - 3)
        End If
        If InStr(1, P1Uses, "ranged") Then P1Uses = Replace(P1Uses, "'s ranged attack", "")
        If InStr(1, P1Uses, "'s") Then P1Uses = Replace(P1Uses, "'s", "")
        P1Uses = Replace(P1Uses, "Cover! ", "")
        If CustomMode = False Then
            Print #EditFile, CurrentLine
        End If
    'SPECIAL/CRIT/SPELL
    ElseIf ((InStr(1, LCase(CurrentLine), " take") <> 0 And InStr(1, LCase(CurrentLine), "damage") <> 0) Or InStr(1, CurrentLine, "HP drained from") <> 0) And InStr(1, "09,0a,01,02,0c,04,0d,05,0e,06,0f", LineType) = 0 And PrevUseType <> "18" Then
        If InStr(1, PrevlineA, "cast") = 0 And InStr(1, PrevlineA, "use") = 0 Then
          P1Special = ""
        End If
        FoundP1 = False
        If InStr(1, CurrentLine, "drained") Then
          MyPos = InStr(3, CurrentLine, " from ")
          P1Takes = Mid$(CurrentLine, MyPos + 6, InStr(1, CurrentLine, ".") - (MyPos + 6))
        Else
          MyPos = InStr(3, CurrentLine, " take")
          P1Takes = Mid$(CurrentLine, 3, MyPos - 3)
          P1Takes = Replace(P1Takes, "Additional effect: ", "")
        End If
        If InStr(1, LCase(P1Takes), "magic burst! ") Then
          P1Takes = Mid$(P1Takes, 14)
          MB = True
        Else
          MB = False
        End If

        MyPos3 = InStr(1, CurrentLine, " points of ")
        If MyPos3 = 0 Then
            MyPos3 = InStr(1, CurrentLine, " point of ")
        End If
        If InStr(1, CurrentLine, "Additional effect: ") Then
            MyPos3 = MyPos3 - 11
            P1Uses = P1
        End If
        If InStr(1, CurrentLine, "drained") Then
            MyPos3 = InStr(1, CurrentLine, " HP ")
            P1Stat = Mid$(CurrentLine, 3, MyPos3 - 3)
            P1Stat = Replace(P1Stat, "Additional effect: ", "")
        Else
            P1Stat = Mid$(CurrentLine, MyPos + 6, MyPos3 - (MyPos + 6))
        End If
        
        If StopDPS = False And BeginDPS = True Then
            If LineType <> "20" And LineType <> "1c" Then
                For i = 0 To UBound(DPS)
                    If LCase(DPS(i, 0)) = LCase(P1Uses) Then
                        DPS(i, 1) = CDbl(DPS(i, 1)) + CDbl(P1Stat)
                        FoundP1 = True
                    End If
                Next
                If FoundP1 = False Then
                    For i = 0 To UBound(DPS)
                        If DPS(i, 0) = "" Then
                            DPS(i, 0) = P1Uses
                            DPS(i, 1) = CDbl(P1Stat)
                            Exit For
                        End If
                    Next
                End If
                FoundP1 = False
            End If
        End If
        
        For i = 0 To UBound(Stats)
            If LCase(Stats(i, 0)) = LCase(P1Uses) And LCase(Stats(i, 1)) = LCase(P1Takes) Then
                If Casts = False Then
                    Stats(i, 4) = CDbl(Stats(i, 4)) + 1
                End If
                If Casts = False And Critical = False Then
                    Stats(i, 5) = CDbl(Stats(i, 5)) + CDbl(P1Stat)
                    Stats(i, 29) = CDbl(Stats(i, 29)) + 1
                ElseIf Casts = True Then
                    Stats(i, 6) = CDbl(Stats(i, 6)) + CDbl(P1Stat)
                ElseIf Critical = True Then
                    Stats(i, 7) = CDbl(Stats(i, 7)) + CDbl(P1Stat)
                    Stats(i, 8) = CDbl(Stats(i, 8)) + 1
                End If
                Stats(i, 9) = CDbl(Stats(i, 9)) + CDbl(P1Stat)
                If Casts = False And Critical = False And P1Special <> "" Then
                  If Stats(i, 11) = "" Then
                      Stats(i, 11) = P1Stat & "(" & P1Special & ")"
                  Else
                      Stats(i, 11) = Stats(i, 11) & ", " & P1Stat & "(" & P1Special & ")"
                  End If
                ElseIf Critical = False And P1Special <> "" Then
                  If MB Then
                    P1Special = P1Special & "-MB"
                  End If
                  If Stats(i, 12) = "" Then
                      Stats(i, 12) = P1Stat & "(" & P1Special & ")"
                  Else
                      Stats(i, 12) = Stats(i, 12) & ", " & P1Stat & "(" & P1Special & ")"
                  End If
                Else
                  If Stats(i, 10) = "" Then
                      Stats(i, 10) = P1Stat
                  Else
                      Stats(i, 10) = Stats(i, 10) & ", " & P1Stat
                  End If
                End If
                P1Special = ""
                FoundP1 = True
            End If
        Next
        If FoundP1 = False Then
            For i = 0 To UBound(Stats)
                If Stats(i, 0) = "" Then
                    NewStatsArray i, 30
                    Stats(i, 0) = P1Uses
                    Stats(i, 1) = P1Takes
                    If Casts = False Then
                        Stats(i, 4) = "1"
                    End If
                    If Casts = False And Critical = False Then
                        Stats(i, 5) = CDbl(Stats(i, 5)) + CDbl(P1Stat)
                        Stats(i, 29) = "1"
                    ElseIf Casts = True Then
                        Stats(i, 6) = CDbl(Stats(i, 6)) + CDbl(P1Stat)
                    ElseIf Critical = True Then
                        Stats(i, 7) = CDbl(Stats(i, 7)) + CDbl(P1Stat)
                        Stats(i, 8) = CDbl(Stats(i, 8)) + 1
                    End If
                    Stats(i, 9) = CDbl(Stats(i, 9)) + CDbl(P1Stat)
                    If Casts = False And Critical = False And P1Special <> "" Then
                        Stats(i, 11) = P1Stat & "(" & P1Special & ")"
                    ElseIf Critical = False And P1Special <> "" Then
                        Stats(i, 12) = P1Stat & "(" & P1Special & ")"
                    Else
                        Stats(i, 10) = P1Stat
                    End If
                    P1Special = ""
                    Exit For
                End If
            Next
        End If
        FoundP1 = False
        For i = 0 To UBound(Stats)
            If LCase(Stats(i, 0)) = LCase(P1Takes) And LCase(Stats(i, 1)) = LCase(P1Uses) Then
                Stats(i, 16) = CDbl(Stats(i, 16)) + 1
                Stats(i, 17) = CDbl(Stats(i, 17)) + CDbl(P1Stat)
                FoundP1 = True
                Exit For
            End If
        Next
        If FoundP1 = False Then
            For i = 0 To UBound(Stats)
                If Stats(i, 0) = "" Then
                    NewStatsArray i, 30
                    Stats(i, 0) = P1Takes
                    Stats(i, 1) = P1Uses
                    Stats(i, 16) = "1"
                    Stats(i, 17) = CDbl(P1Stat)
                    Exit For
                End If
            Next
        End If
        FoundP1 = False
        If CustomMode = False Then
            Print #EditFile, CurrentLine
        End If
    ElseIf CurrentFight <> "" And InStr(1, LCase(CurrentLine), " recovers ") <> 0 And InStr(1, CurrentLine, " MP.") = 0 And Mid$(CurrentLine, 3, 1) <> "<" And Mid$(CurrentLine, 3, 1) <> ">" And Mid$(CurrentLine, 3, 1) <> "(" And InStr(1, CurrentLine, " : ") = 0 Then
    
        MyPos = InStr(1, CurrentLine, " recovers ")
        MyPos2 = InStr(1, CurrentLine, " HP")
        P1Stat = Mid$(CurrentLine, MyPos + 10, MyPos2 - (MyPos + 10))
        MyPos = InStr(3, CurrentLine, " recovers ")
        P1 = Mid$(CurrentLine, 3, MyPos - 3)
        For i = 0 To UBound(Stats)
            If LCase(Stats(i, 0)) = LCase(P1) And LCase(Stats(i, 1)) = LCase(CurrentFight) Then
                If Stats(i, 14) = "" Then Stats(i, 14) = "0"
                Stats(i, 14) = CDbl(Stats(i, 14)) + P1Stat
                FoundP1 = True
            End If
        Next

        If FoundP1 = False Then
            For i = 0 To UBound(Stats)
                If Stats(i, 0) = "" Then
                    NewStatsArray i, 30
                    Stats(i, 0) = P1
                    Stats(i, 1) = CurrentFight
                    Stats(i, 14) = P1Stat
                    Exit For
                End If
            Next
        End If
        FoundP1 = False
        For i = 0 To UBound(Stats)
            If LCase(Stats(i, 0)) = LCase(P1Uses) And LCase(Stats(i, 1)) = LCase(CurrentFight) Then
                Stats(i, 26) = CDbl(Stats(i, 26)) + P1Stat
                FoundP1 = True
            End If
        Next

        If FoundP1 = False Then
            For i = 0 To UBound(Stats)
                If Stats(i, 0) = "" Then
                    NewStatsArray i, 30
                    Stats(i, 0) = P1Uses
                    Stats(i, 1) = CurrentFight
                    Stats(i, 26) = P1Stat
                    Exit For
                End If
            Next
        End If
        If CustomMode = False Then
            Print #EditFile, CurrentLine
        End If
      'FINISHED
    ElseIf ((InStr(1, LCase(CurrentLine), "falls to the ground") <> 0 Or InStr(1, LCase(CurrentLine), "fall to the ground") <> 0) Or InStr(1, LCase(CurrentLine), "defeats") <> 0) And InStr(1, "09,0a,01,02,0c,04,0d,05,0e,06,0f,2c", LineType) = 0 Then
        If InStr(1, LCase(CurrentLine), "defeats") Then
            MyPos = InStr(1, CurrentLine, "defeats ")
            MyPos2 = InStr(1, CurrentLine, ".")
            P1Opp = Mid$(CurrentLine, MyPos + 8, MyPos2 - (MyPos + 8))
        Else
            MyPos = InStr(1, CurrentLine, "fall")
            P1Opp = Mid$(CurrentLine, 3, MyPos - 4)
        End If
        If CustomMode = True Then
            If listResults.Selected(UniqueMOB) = True Then
                CustomAdd = True
            Else
                CustomAdd = False
            End If
        Else
            CustomAdd = True
        End If
        If CustomAdd Then
            For i = 0 To UBound(Stats)
                If LCase(Stats(i, 1)) = LCase(P1Opp) Or LCase(Stats(i, 0)) = LCase(P1Opp) Then
                    If mnuPlayer.Checked = True Then
                        If InStr(1, Stats(i, 0), " ") = 0 Or InStr(1, Stats(i, 0), "Skillchain") <> 0 Then
                            SkipIt = False
                        Else
                            SkipIt = True
                        End If
                    ElseIf mnuMonster.Checked = True Then
                        If InStr(1, Stats(i, 0), " ") <> 0 And InStr(1, Stats(i, 0), "Skillchain") = 0 Then
                            SkipIt = False
                        Else
                            SkipIt = True
                        End If
                    Else
                        SkipIt = False
                    End If
                    If SkipIt = False Then
                        If LCase(Stats(i, 0)) <> LCase(P1Opp) Then
                          If Stats(i, 10) = "" And Stats(i, 11) = "" And Stats(i, 12) = "" And Stats(i, 13) = "" Then
                          Else
                            With RTB_Details
                              .SelBold = True
                              .SelColor = vbBlue
                              .SelText = Stats(i, 0) & " - " & Stats(i, 1) & vbNewLine
                              .SelColor = vbBlack
                              .SelBold = False

                              If Stats(i, 10) <> "" Then
                                  .SelBold = True
                                  .SelText = vbTab & "Basic Damage: "
                                  .SelBold = False
                                  .SelText = Stats(i, 10) & vbNewLine
                              End If
                              If Stats(i, 13) <> "" Then
                                  .SelBold = True
                                  .SelText = vbTab & "Ranged Damage: "
                                  .SelBold = False
                                  .SelText = Stats(i, 13) & vbNewLine
                              End If
                              If Stats(i, 11) <> "" Then
                                  .SelBold = True
                                  .SelText = vbTab & "WeaponSkills: "
                                  .SelBold = False
                                  .SelText = Stats(i, 11) & vbNewLine
                              End If
                              If Stats(i, 12) <> "" Then
                                  .SelBold = True
                                  .SelText = vbTab & "Spells: "
                                  .SelBold = False
                                  .SelText = Stats(i, 12) & vbNewLine
                              End If
                              .SelBold = True
                              .SelText = vbTab & "Total Damage: "
                              .SelBold = False
                              .SelColor = vbRed
                              .SelText = Stats(i, 9) & vbNewLine
                              .SelColor = vbBlack
                              .SelText = vbNewLine
                            End With
                          End If
                        End If

                    End If
                    If (InStr(1, Stats(i, 0), " ") = 0 Or InStr(1, Stats(i, 0), "Skillchain") <> 0) And LCase(Stats(i, 0)) <> LCase(P1Opp) Then
                        TotalDMG = CDbl(TotalDMG) + CDbl(Stats(i, 9))
                        TotalHit = CDbl(TotalHit) + CDbl(Stats(i, 4))
                        TotalSwing = CDbl(TotalSwing) + CDbl(Stats(i, 3)) + CDbl(Stats(i, 4))
                        TotalSwingTaken = CDbl(TotalSwingTaken) + CDbl(Stats(i, 15)) + CDbl(Stats(i, 16))
                        TotalBase = TotalBase + CDbl(Stats(i, 2)) + CDbl(Stats(i, 7))
                        TotalSpell = TotalSpell + CDbl(Stats(i, 6))
                        TotalSkill = TotalSkill + CDbl(Stats(i, 5))
                        TotalTaken = TotalTaken + CDbl(Stats(i, 17))
                        TotalEffect = TotalEffect + CDbl(Stats(i, 28))
                        TotalHP = TotalHP + CDbl(Stats(i, 14))
                        TotalHPH = TotalHPH + CDbl(Stats(i, 26))
                        
                        For p = 0 To UBound(PList)
                            If PList(p, 0) = "" Then
                                PList(p, 0) = Stats(i, 0)
                                Do Until Len(PList(p, 0)) >= 25
                                    PList(p, 0) = PList(p, 0) & " "
                                Loop
                                PList(p, 1) = Stats(i, 9)
                                PList(p, 4) = Stats(i, 4)
                                PList(p, 5) = CDbl(Stats(i, 3)) + CDbl(Stats(i, 4))
                                PList(p, 6) = Stats(i, 3)
                                PList(p, 7) = Stats(i, 8)
                                PList(p, 8) = Stats(i, 10)
                                PList(p, 9) = Stats(i, 2)
                                PList(p, 10) = Stats(i, 7)
                                PList(p, 11) = Stats(i, 5)
                                PList(p, 12) = Stats(i, 6)
                                PList(p, 13) = Stats(i, 14)
                                PList(p, 14) = Stats(i, 15)
                                PList(p, 15) = Stats(i, 16)
                                PList(p, 16) = Stats(i, 17)
                                PList(p, 17) = Stats(i, 18)
                                PList(p, 18) = Stats(i, 19)
                                PList(p, 19) = Stats(i, 21)
                                PList(p, 20) = Stats(i, 22)
                                PList(p, 21) = Stats(i, 23)
                                PList(p, 22) = Stats(i, 24)
                                PList(p, 23) = Stats(i, 25)
                                PList(p, 24) = Stats(i, 26)
                                PList(p, 25) = Stats(i, 27)
                                PList(p, 26) = Stats(i, 28)
                                PList(p, 27) = Stats(i, 29)
                                PList(p, 28) = Stats(i, 30)
                                If InStr(1, PList(p, 0), "Skillchain: ") = 0 Then
                                    If CDbl(Stats(i, 9)) > dHigh Then
                                        dHigh = CDbl(Stats(i, 9))
                                    End If
                                    If CDbl(Stats(i, 9)) < dLow And CDbl(Stats(i, 9)) <> 0 Then
                                        dLow = CDbl(Stats(i, 9))
                                    End If
                                End If
                                Exit For
                            End If
                        Next
                        FoundP1 = False
                        For p = 0 To UBound(GrandPList)
                            If Trim(GrandPList(p, 0)) = Trim(Stats(i, 0)) Then
                                FoundP1 = True
                                GrandPList(p, 1) = CDbl(GrandPList(p, 1)) + CDbl(Stats(i, 9))
                                GrandPList(p, 4) = CDbl(GrandPList(p, 4)) + CDbl(Stats(i, 4))
                                GrandPList(p, 5) = CDbl(GrandPList(p, 5)) + CDbl(Stats(i, 3)) + CDbl(Stats(i, 4))
                                GrandPList(p, 6) = CDbl(GrandPList(p, 6)) + CDbl(Stats(i, 3))
                                GrandPList(p, 7) = CDbl(GrandPList(p, 7)) + CDbl(Stats(i, 8))
                                GrandPList(p, 9) = CDbl(GrandPList(p, 9)) + CDbl(Stats(i, 2))
                                GrandPList(p, 10) = CDbl(GrandPList(p, 10)) + CDbl(Stats(i, 7))
                                GrandPList(p, 11) = CDbl(GrandPList(p, 11)) + CDbl(Stats(i, 5))
                                GrandPList(p, 12) = CDbl(GrandPList(p, 12)) + CDbl(Stats(i, 6))
                                GrandPList(p, 13) = CDbl(GrandPList(p, 13)) + CDbl(Stats(i, 14))
                                GrandPList(p, 14) = CDbl(GrandPList(p, 14)) + CDbl(Stats(i, 15))
                                GrandPList(p, 15) = CDbl(GrandPList(p, 15)) + CDbl(Stats(i, 16))
                                GrandPList(p, 16) = CDbl(GrandPList(p, 16)) + CDbl(Stats(i, 17))
                                GrandPList(p, 19) = CDbl(GrandPList(p, 19)) + CDbl(Stats(i, 21))
                                GrandPList(p, 20) = CDbl(GrandPList(p, 20)) + CDbl(Stats(i, 22))
                                GrandPList(p, 21) = CDbl(GrandPList(p, 21)) + CDbl(Stats(i, 23))
                                GrandPList(p, 22) = CDbl(GrandPList(p, 22)) + CDbl(Stats(i, 24))
                                GrandPList(p, 23) = CDbl(GrandPList(p, 23)) + CDbl(Stats(i, 25))
                                GrandPList(p, 24) = CDbl(GrandPList(p, 24)) + CDbl(Stats(i, 26))
                                GrandPList(p, 25) = CDbl(GrandPList(p, 25)) + CDbl(Stats(i, 27))
                                GrandPList(p, 26) = CDbl(GrandPList(p, 26)) + CDbl(Stats(i, 28))
                                GrandPList(p, 27) = CDbl(GrandPList(p, 27)) + CDbl(Stats(i, 29))
                                GrandPList(p, 28) = CDbl(GrandPList(p, 28)) + CDbl(Stats(i, 30))
                                If CDbl(Stats(i, 18)) > CDbl(GrandPList(p, 17)) Then
                                    GrandPList(p, 17) = CDbl(Stats(i, 18))
                                End If
                                If CDbl(Stats(i, 19)) < CDbl(GrandPList(p, 18)) Then
                                    GrandPList(p, 18) = CDbl(Stats(i, 19))
                                End If
                                Exit For
                            End If
                        Next
                        If FoundP1 = False Then
                            For p = 0 To UBound(GrandPList)
                                If Trim(GrandPList(p, 0)) = "" Then
                                    GrandPList(p, 0) = Stats(i, 0)
                                    GrandPList(p, 1) = Stats(i, 9)
                                    GrandPList(p, 4) = Stats(i, 4)
                                    GrandPList(p, 5) = CDbl(Stats(i, 3)) + CDbl(Stats(i, 4))
                                    GrandPList(p, 6) = Stats(i, 3)
                                    GrandPList(p, 7) = Stats(i, 8)
                                    GrandPList(p, 9) = Stats(i, 2)
                                    GrandPList(p, 10) = Stats(i, 7)
                                    GrandPList(p, 11) = Stats(i, 5)
                                    GrandPList(p, 12) = Stats(i, 6)
                                    GrandPList(p, 13) = Stats(i, 14)
                                    GrandPList(p, 14) = Stats(i, 15)
                                    GrandPList(p, 15) = Stats(i, 16)
                                    GrandPList(p, 16) = Stats(i, 17)
                                    GrandPList(p, 17) = Stats(i, 18)
                                    GrandPList(p, 18) = Stats(i, 19)
                                    GrandPList(p, 19) = Stats(i, 21)
                                    GrandPList(p, 20) = Stats(i, 22)
                                    GrandPList(p, 21) = Stats(i, 23)
                                    GrandPList(p, 22) = Stats(i, 24)
                                    GrandPList(p, 23) = Stats(i, 25)
                                    GrandPList(p, 24) = Stats(i, 26)
                                    GrandPList(p, 25) = Stats(i, 27)
                                    GrandPList(p, 26) = Stats(i, 28)
                                    GrandPList(p, 27) = Stats(i, 29)
                                    GrandPList(p, 28) = Stats(i, 30)
                                    Exit For
                                End If
                            Next
                        End If
                    End If
                    For StatClear = 0 To 30
                        Stats(i, StatClear) = ""
                    Next
                End If
            Next


            For pl = 0 To UBound(PList)
              If PList(pl, 0) <> "" Then
                  For p = 0 To UBound(Players)
                    If Players(p, 0) = "" Then
                        Players(p, 0) = PList(pl, 0)
                        Players(p, 1) = "1"
                        Players(p, 2) = PList(pl, 1)
                        If TotalDMG <> 0 Then Players(p, 3) = (PList(pl, 1) / TotalDMG) * 100
                        If CDbl(PList(pl, 4)) <> 0 And CDbl(PList(pl, 5)) <> 0 Then
                            Players(p, 4) = Format(Round((CDbl(PList(pl, 4)) / CDbl(PList(pl, 5))) * 100, 2), "0#.#0")
                        Else
                            Players(p, 4) = "00.00"
                        End If
                        Exit For
                    ElseIf Players(p, 0) = PList(pl, 0) Then
                        Players(p, 1) = Players(p, 1) + 1
                        Players(p, 2) = CDbl(Players(p, 2)) + CDbl(PList(pl, 1))
                        If CDbl(PList(pl, 1)) <> 0 Then
                            Players(p, 3) = CDbl(Players(p, 3)) + CDbl(((CDbl(PList(pl, 1)) / CDbl(TotalDMG)) * 100))
                        Else
                            Players(p, 3) = "0"
                        End If
                        If CDbl(PList(pl, 4)) <> 0 And CDbl(PList(pl, 5)) <> 0 Then
                            Players(p, 4) = CDbl(Players(p, 4)) + Format(Round((CDbl(PList(pl, 4)) / CDbl(PList(pl, 5))) * 100, 2), "0#.#0")
                        End If
                        Exit For
                    End If
                  Next
              End If
            Next
            For p = 0 To UBound(PList)
                If PList(p, 0) <> "" Then
                    If TotalDMG <> 0 Then
                        PList(p, 2) = Round((CDbl(PList(p, 1)) / TotalDMG) * 100, 2)
                    Else
                        PList(p, 2) = 0
                    End If
                    If CDbl(PList(p, 1)) = dHigh Then
                        PList(p, 3) = "H"
                    ElseIf CDbl(PList(p, 1)) = dLow Then
                        PList(p, 3) = "L"
                    End If
                Else
                    Exit For
                End If
            Next
            SelStart = Len(RTB_Report.Text)
      
            If TotalDMG <> 0 Then
              ExpLine = ff
              Do Until InStr(1, NextLine, "experience points.") <> 0
                If ExpLine + 1 <= UBound(FullFile) Then
                    NextLine = FullFile(ExpLine + 1)
                Else
                    NextLine = ""
                    Exit Do
                End If
                ExpChecks = ExpChecks + 1
                ExpLine = ExpLine + 1
                If ExpChecks > 160 Then
                    NextLine = ""
                    Exit Do
                End If
              Loop
              If NextLine <> "" Then
                MyPos = InStr(1, NextLine, "gains ")
                MyPos2 = InStr(1, NextLine, "exp")
                ExpGain = "(" & Mid$(NextLine, MyPos + 6, MyPos2 - (MyPos + 7)) & " exp)"
              Else
                ExpGain = ""
              End If
              ExpLine = 0
              ExpChecks = 0
              With RTB_Report
                .SelText = UniqueMOB & " - " & Replace(P1Opp, "the ", "The ") & ExpGain & MonsterCheck & " " & FightComment & vbNewLine
                .SelStart = SelStart
                .SelLength = Len(UniqueMOB & " - " & Replace(P1Opp, "the ", "The ") & ExpGain & MonsterCheck & " " & FightComment)
                .SelBold = True
                .SelColor = &H40097
                .SelStart = Len(.Text)
                .SelColor = vbBlack
                .SelBold = True
                .SelUnderline = True
                .SelText = ColumnText & vbNewLine
                .SelUnderline = False
                .SelBold = False
                .SelStart = Len(.Text)
              End With
              With RTB_Averages
                .Text = ""
                .SelBold = True
                .SelText = "Experience" & vbNewLine
                .SelBold = False
                If TotalExp <> 0 And StartTime <> "12:00:00 AM" And StopTime <> "12:00:00 AM" Then
                  .SelText = "Start: " & StartTime & " / Stop: " & StopTime & vbNewLine & "Total Exp: " & TotalExp & vbNewLine & "Per Hour: " & Round(TotalExp / DateDiff("n", StartTime, StopTime), 2) * 60 & vbNewLine & "Per Minute: " & Round(TotalExp / DateDiff("n", StartTime, StopTime), 2) & vbNewLine & vbNewLine
                ElseIf TotalExp <> 0 And StartTime <> "12:00:00 AM" Then
                  .SelText = "Start: " & StartTime & " / Stop: " & Now & vbNewLine & "Total Exp: " & TotalExp & vbNewLine & "Per Hour: " & Round(TotalExp / DateDiff("n", StartTime, Now), 2) * 60 & vbNewLine & "Per Minute: " & Round(TotalExp / DateDiff("n", StartTime, Now), 2) & vbNewLine & vbNewLine
                Else
                  .SelText = "Start: " & StartTime & vbNewLine & "Total Exp: " & TotalExp & vbNewLine & "Per Hour: 0" & vbNewLine & "Per Minute: 0" & vbNewLine & vbNewLine
                End If
              End With
              
              For p = 0 To UBound(Players)
                  If Players(p, 0) <> "" Then
                    With RTB_Averages
                      SelStart = Len(.Text)
                      .SelText = Players(p, 0)
                      .SelStart = SelStart
                      .SelLength = Len(Players(p, 0))
                      .SelBold = True
                      .SelStart = Len(.Text)
                      .SelBold = False
                      AvgAcc = Round(Players(p, 4) / Players(p, 1), 2)
                      EstDPS = ""
                      For dp = 0 To UBound(DPS)
                        If DPS(dp, 0) = Trim(Players(p, 0)) Then
                            If DPS(dp, 0) <> "" Then
                                If DPS(dp, 1) <> "0" And DPS(dp, 2) <> "0" And DPS(dp, 2) <> "" And DPS(dp, 1) <> "" Then
                                    EstDPS = Round(DPS(dp, 1) / DPS(dp, 2), 2) & " (" & DPS(dp, 2) & " seconds / " & DPS(dp, 1) & " dmg)"
                                Else
                                    EstDPS = 0
                                End If
                                Exit For
                            End If
                        End If
                      Next
                      If Players(p, 3) = "0" Or Players(p, 3) = "" Then Players(p, 3) = "1"
                      If Players(p, 2) = "0" Or Players(p, 2) = "" Then Players(p, 2) = "1"
                      .SelText = vbNewLine & ResizePart("Total Fights:", 1500) & vbTab & Players(p, 1) & vbNewLine & ResizePart("Average Damage:", 1500) & vbTab & Round(CDbl(Players(p, 2)) / CDbl(Players(p, 1)), 3) & vbNewLine & ResizePart("Average Percent:", 1500) & vbTab & Round(CDbl(Players(p, 3)) / CDbl(Players(p, 1)), 3) & vbNewLine & ResizePart("Average Accuracy:", 1500) & vbTab & AvgAcc & vbNewLine & ResizePart("Estimated DPS:", 1500) & vbTab & EstDPS & vbNewLine & vbNewLine
                    End With
                  End If
              Next
              RTB_Averages.SelStart = 0
              LastPlayer = False
              For p = 0 To UBound(PList)
                  If PList(p, 0) <> "" Then
                    If p = UBound(PList) Then
                        LastPlayer = True
                    ElseIf PList(p + 1, 0) = "" Then
                        LastPlayer = True
                    End If
                    SelStart = Len(RTB_Report.Text)
                    Erase Part
                    Part(0) = ResizePart(Replace(Trim(PList(p, 0)), "Skillchain: ", "SC: "), 1000)
                    Part(1) = ResizePart(Trim(PList(p, 1)), 525)
                    If ReportOptions(18, 0) = 1 Then Part(2) = ResizePart(Replace(Format(PList(p, 2), "0.#0"), "100.00", "100") & "%", 525)
                    If ReportOptions(0, 0) = 1 Then Part(3) = ResizePart(CDbl(PList(p, 9)) + CDbl(PList(p, 10)), 525)
                    If ReportOptions(1, 0) = 1 Then Part(4) = ResizePart(PList(p, 11), 525)
                    If ReportOptions(2, 0) = 1 Then Part(5) = ResizePart(PList(p, 12), 525)
                    If ReportOptions(3, 0) = 1 Then Part(6) = ResizePart(PList(p, 17) & "/" & PList(p, 18), 525)
                    If ReportOptions(4, 0) = 1 Then
                        If PList(p, 4) <> 0 Then
                            Part(7) = ResizePart(CStr(Round(CDbl(PList(p, 9)) / CDbl(PList(p, 4)), 2)), 525)
                        Else
                            Part(7) = ResizePart("0", 525)
                        End If
                    End If
                    If ReportOptions(5, 0) = 1 Then
                        If PList(p, 7) <> 0 And PList(p, 4) <> 0 Then
                            Part(8) = ResizePart(CStr(Round((CDbl(PList(p, 7)) / CDbl(PList(p, 4)) * 100), 2)) & "%", 525)
                        Else
                            Part(8) = ResizePart("0.00%", 525)
                        End If
                    End If
                    If ReportOptions(6, 0) = 1 Then Part(9) = ResizePart(PList(p, 7), 525)
                    If ReportOptions(21, 0) = 1 Then Part(10) = ResizePart(PList(p, 26), 525)
                    If ReportOptions(22, 0) = 1 Then Part(11) = ResizePart(PList(p, 27), 525)
                    If ReportOptions(7, 0) = 1 Then
                        If CDbl(PList(p, 4)) <> 0 And CDbl(PList(p, 5)) <> 0 Then
                            Part(12) = ResizePart(Format(Round((CDbl(PList(p, 4)) / CDbl(PList(p, 5))) * 100, 2), "0.#0") & "% ", 525)
                        Else
                            Part(12) = ResizePart("0.00%", 525)
                        End If
                    End If
                    If ReportOptions(8, 0) = 1 Then Part(13) = ResizePart(PList(p, 4) & "/" & PList(p, 6), 525)
                    If ReportOptions(9, 0) = 1 Then
                        If CDbl(PList(p, 14)) <> 0 Then
                            Part(14) = ResizePart(CStr(Round((CDbl(PList(p, 14)) / (CDbl(PList(p, 14)) + CDbl(PList(p, 15))) * 100), 2)) & "%", 525)
                        Else
                            Part(14) = ResizePart("0.00%", 525)
                        End If
                    End If
                    If ReportOptions(10, 0) = 1 Then Part(15) = ResizePart(PList(p, 15) & "/" & PList(p, 14), 525)
                    If ReportOptions(11, 0) = 1 Then Part(16) = ResizePart(PList(p, 19), 525)
                    If ReportOptions(12, 0) = 1 Then Part(17) = ResizePart(PList(p, 20), 525)
                    If ReportOptions(13, 0) = 1 Then Part(18) = ResizePart(PList(p, 21), 525)
                    If ReportOptions(14, 0) = 1 Then Part(19) = ResizePart(PList(p, 22), 525)
                    If ReportOptions(15, 0) = 1 Then Part(20) = ResizePart(PList(p, 23), 525)
                    If ReportOptions(20, 0) = 1 Then Part(21) = ResizePart(PList(p, 25), 525)
                    If ReportOptions(23, 0) = 1 Then Part(22) = ResizePart(PList(p, 28), 525)
                    If ReportOptions(16, 0) = 1 Then Part(23) = ResizePart(PList(p, 16), 525)
                    If ReportOptions(17, 0) = 1 Then Part(24) = ResizePart(PList(p, 13), 525)
                    If ReportOptions(19, 0) = 1 Then Part(25) = ResizePart(PList(p, 24), 525)
                    
                    CurrentSel = ""
                    For Parts = 0 To 25
                        If Part(Parts) <> "" Then
                            CurrentSel = CurrentSel & Part(Parts) & vbTab
                        End If
                    Next
                    If LastPlayer Then RTB_Report.SelUnderline = True
                    RTB_Report.SelText = CurrentSel & vbNewLine
                    RTB_Report.SelUnderline = False
                      
                      If InStr(1, PList(p, 0), "Skillchain: ") = 0 Then
                        With RTB_Report
                          .SelStart = SelStart
                          .SelLength = Len(CurrentSel)
                          If PList(p, 3) = "H" Then
                              .SelColor = vbBlue
                          ElseIf PList(p, 3) = "L" Then
                              .SelColor = vbRed
                          End If
                          .SelStart = Len(.Text)
                        End With
                      End If
                        'ADD TO SINGLE USER RPT
                        For u = 0 To UBound(UserLog)
                          If LCase(Trim(PList(p, 0))) = LCase(Trim(UserLog(u, 0))) Then
  
                              If Mid$(UserLog(u, 7), 1, 1) = "1" Then UserLog(u, 1) = UserLog(u, 1) & ResizeSimplePart(PList(p, 1), 10) & vbTab
                              If Mid$(UserLog(u, 7), 3, 1) = "1" Then UserLog(u, 1) = UserLog(u, 1) & ResizeSimplePart(PList(p, 9), 10) & vbTab
                              If Mid$(UserLog(u, 7), 5, 1) = "1" Then UserLog(u, 1) = UserLog(u, 1) & ResizeSimplePart(PList(p, 10), 10) & vbTab
                              If Mid$(UserLog(u, 7), 7, 1) = "1" Then UserLog(u, 1) = UserLog(u, 1) & ResizeSimplePart(PList(p, 11), 10) & vbTab
                              If Mid$(UserLog(u, 7), 9, 1) = "1" Then UserLog(u, 1) = UserLog(u, 1) & ResizeSimplePart(PList(p, 12), 10) & vbTab
                              If Mid$(UserLog(u, 7), 11, 1) = "1" Then UserLog(u, 1) = UserLog(u, 1) & ResizeSimplePart(PList(p, 7), 10) & vbTab
                              If Mid$(UserLog(u, 7), 13, 1) = "1" Then
                                PartA = ResizeSimplePart(CStr(PList(p, 4)) & "/" & CStr(PList(p, 6)), 10)
                                UserLog(u, 1) = UserLog(u, 1) & PartA & vbTab
                              End If
                              If CDbl(PList(p, 4)) <> 0 And CDbl(PList(p, 5)) <> 0 Then
                                If Mid$(UserLog(u, 7), 15, 1) = "1" Then
                                  PartA = ResizeSimplePart(Format(Round((CDbl(PList(p, 4)) / CDbl(PList(p, 5))) * 100, 2), "0#.#0") & "%", 10)
                                  UserLog(u, 1) = UserLog(u, 1) & PartA & vbTab
                                End If
                              Else
                                UserLog(u, 1) = UserLog(u, 1) & ResizeSimplePart("0.00%", 10) & vbTab
                              End If
                              UserLog(u, 1) = UserLog(u, 1) & ResizeSimplePart(Replace(P1Opp, "the ", "The ") & MonsterCheck, 23)
                              If Mid$(UserLog(u, 7), 17, 1) = "1" Then
                                UserLog(u, 1) = UserLog(u, 1) & vbTab & FightComment & vbNewLine
                              Else
                                UserLog(u, 1) = UserLog(u, 1) & vbNewLine
                              End If
                              UserLog(u, 2) = CDbl(UserLog(u, 2)) + CDbl(PList(p, 1)) 'Total Damage
                              UserLog(u, 3) = CDbl(UserLog(u, 3)) + CDbl(PList(p, 4)) 'Total Hits
                              UserLog(u, 4) = CDbl(UserLog(u, 4)) + CDbl(PList(p, 6)) 'Total Misses
                              UserLog(u, 5) = CDbl(UserLog(u, 5)) + CDbl(PList(p, 7)) 'Total Crits
                              UserLog(u, 6) = CDbl(UserLog(u, 6)) + 1 'Total Fights
                              UserLog(u, 8) = CDbl(UserLog(u, 8)) + CDbl(PList(p, 9)) 'Total Base dmg
                              UserLog(u, 9) = CDbl(UserLog(u, 9)) + CDbl(PList(p, 10)) 'Total Crit dmg
                              UserLog(u, 10) = CDbl(UserLog(u, 10)) + CDbl(PList(p, 11)) 'Total Skill dmg
                              UserLog(u, 11) = CDbl(UserLog(u, 11)) + CDbl(PList(p, 12)) 'Total Spell dmg
                          End If
                        Next
                  Else
                      Exit For
                  End If
              Next
              SelStart = Len(RTB_Report.Text)
              
              
            Erase Part
            Part(0) = ResizePart("Totals:", 1000)
            Part(1) = ResizePart(CStr(TotalDMG), 525)
            If ReportOptions(18, 0) = 1 Then Part(2) = ResizePart("100%", 525)
            If ReportOptions(0, 0) = 1 Then Part(3) = ResizePart(CStr(TotalBase), 525)
            If ReportOptions(1, 0) = 1 Then Part(4) = ResizePart(CStr(TotalSkill), 525)
            If ReportOptions(2, 0) = 1 Then Part(5) = ResizePart(CStr(TotalSpell), 525)
            If ReportOptions(3, 0) = 1 Then Part(6) = ResizePart("", 525)
            If ReportOptions(4, 0) = 1 Then Part(7) = ResizePart("", 525)
            If ReportOptions(5, 0) = 1 Then Part(8) = ResizePart("", 525)
            If ReportOptions(6, 0) = 1 Then Part(9) = ResizePart("", 525)
            If ReportOptions(21, 0) = 1 Then Part(10) = ResizePart("", 525)
            If ReportOptions(22, 0) = 1 Then Part(11) = ResizePart("", 525)
            If ReportOptions(7, 0) = 1 Then Part(12) = ResizePart("", 525)
            If ReportOptions(8, 0) = 1 Then Part(13) = ResizePart("", 525)
            If ReportOptions(9, 0) = 1 Then Part(14) = ResizePart("", 525)
            If ReportOptions(10, 0) = 1 Then Part(15) = ResizePart("", 525)
            If ReportOptions(11, 0) = 1 Then Part(16) = ResizePart("", 525)
            If ReportOptions(12, 0) = 1 Then Part(17) = ResizePart("", 525)
            If ReportOptions(13, 0) = 1 Then Part(18) = ResizePart("", 525)
            If ReportOptions(14, 0) = 1 Then Part(19) = ResizePart("", 525)
            If ReportOptions(15, 0) = 1 Then Part(20) = ResizePart("", 525)
            If ReportOptions(20, 0) = 1 Then Part(21) = ResizePart("", 525)
            If ReportOptions(23, 0) = 1 Then Part(22) = ResizePart("", 525)
            If ReportOptions(16, 0) = 1 Then Part(23) = ResizePart(CStr(TotalTaken), 525)
            If ReportOptions(17, 0) = 1 Then Part(24) = ResizePart(CStr(TotalHP), 525)
            If ReportOptions(19, 0) = 1 Then Part(25) = ResizePart(CStr(TotalHPH), 525)
                    
                    
              CurrentSel = ""
              For Parts = 0 To 25
                  If Part(Parts) <> "" Then
                      CurrentSel = CurrentSel & Part(Parts) & vbTab
                  End If
              Next
              With RTB_Report
                .SelText = CurrentSel & vbNewLine & vbNewLine
                .SelStart = SelStart
                .SelLength = Len(CurrentSel)
                .SelBold = True
                .SelStart = Len(.Text)
              End With
            End If

            If CustomMode = False Then
                For i = 0 To comboMOB.ListCount - 1
                    If comboMOB.List(i) = Replace(P1Opp, "the ", "The ") Then
                        AddMOB = False
                        Exit For
                    Else
                        AddMOB = True
                    End If
                Next
                If AddMOB Or comboMOB.ListCount = 0 Then
                    comboMOB.AddItem Replace(P1Opp, "the ", "The ")
                End If
                If FightComment <> "" Then
                    listResults.AddItem UniqueMOB & " " & Replace(P1Opp, "the ", "The ") & " (" & TotalDMG & ") - " & FightComment
                Else
                    listResults.AddItem UniqueMOB & " " & Replace(P1Opp, "the ", "The ") & " (" & TotalDMG & ")"
                End If
                listResults.Selected(listResults.ListCount - 1) = True
                Print #EditFile, CurrentLine
            End If
        Else
            For i = 0 To UBound(Stats)
                If LCase(Stats(i, 1)) = LCase(P1Opp) Or LCase(Stats(i, 0)) = LCase(P1Opp) Then
                    If mnuPlayer.Checked = True Then
                        If InStr(1, Stats(i, 0), " ") = 0 Or InStr(1, Stats(i, 0), "Skillchain") <> 0 Then
                            SkipIt = False
                        Else
                            SkipIt = True
                        End If
                    ElseIf mnuMonster.Checked = True Then
                        If InStr(1, Stats(i, 0), " ") <> 0 And InStr(1, Stats(i, 0), "Skillchain") = 0 Then
                            SkipIt = False
                        Else
                            SkipIt = True
                        End If
                    Else
                        SkipIt = False
                    End If
                    If SkipIt = False Then
                        For StatClear = 0 To 30
                            Stats(i, StatClear) = ""
                        Next
                    End If
                End If
            Next
        End If

        If TotalDMG <> 0 And GenerateHTML And ExportOptions(18, 0) = 0 Then
            HTMLCodeNew = ""
            HTMLCodeNew = HTMLCodeNew & "<CENTER><TABLE BORDER=1 CELLPADDING=0 CELLSPACING=0 style=""PADDING-LEFT: 3px;PADDING-RIGHT: 3px;BORDER-COLLAPSE:collapse;font-family:verdana;font-size:7pt;color:black"">" & vbNewLine
            HTMLCodeNew = HTMLCodeNew & "<TR><TH colSpan=25 align=""Left"" BGColor=""7CB1CB"">" & Replace(P1Opp, "the ", "The ") & " - (ID: " & UniqueMOB & ")</font></TH></TR>" & vbNewLine
            HTMLCodeNew = HTMLCodeHeader
        End If
        For p = 0 To UBound(PList)
            If PList(p, 0) <> "" Then
                If GenerateHTML And ExportOptions(18, 0) = 0 Then
                    HTMLCodeNew = HTMLCodeNew & "<TR style=""BACKGROUND-COLOR:#dae6ec"">" & vbNewLine
                    HTMLCodeNew = HTMLCodeNew & "<TD BGCOLOR=""#b8ced9""><b>" & Replace(Trim$(PList(p, 0)), "Skillchain: ", "SC:") & "</b></TD>" & vbNewLine 'PLAYER NAME
                    If ExportOptions(0, 0) = 1 Then HTMLCodeNew = HTMLCodeNew & "<TD>" & CDbl(PList(p, 9)) + CDbl(PList(p, 10)) & "</TD>" & vbNewLine                        'BASIC DMG
                    If ExportOptions(1, 0) = 1 Then HTMLCodeNew = HTMLCodeNew & "<TD>" & PList(p, 11) & "</TD>" & vbNewLine                        'SKILL DMG
                    If ExportOptions(2, 0) = 1 Then HTMLCodeNew = HTMLCodeNew & "<TD>" & PList(p, 12) & "</TD>" & vbNewLine                        'SPELL DMG
                    If ExportOptions(21, 0) = 1 Then HTMLCodeNew = HTMLCodeNew & "<TD>" & PList(p, 26) & "</TD>" & vbNewLine                        'EFFECT DMG
                    If ExportOptions(22, 0) = 1 Then HTMLCodeNew = HTMLCodeNew & "<TD>" & PList(p, 27) & "</TD>" & vbNewLine                        'WS #
                    If ExportOptions(3, 0) = 1 Then HTMLCodeNew = HTMLCodeNew & "<TD>" & PList(p, 17) & "/" & PList(p, 18) & "</TD>" & vbNewLine                        'High/Low
                    If ExportOptions(4, 0) = 1 Then
                        If PList(p, 5) <> 0 And (CDbl(PList(p, 9)) + CDbl(PList(p, 10))) <> 0 Then 'Average
                            HTMLCodeNew = HTMLCodeNew & "<TD>" & Round((CDbl(PList(p, 9)) + CDbl(PList(p, 10))) / PList(p, 4), 2) & "</TD>" & vbNewLine
                        Else
                            HTMLCodeNew = HTMLCodeNew & "<TD>0</TD>" & vbNewLine
                        End If
                    End If
                    If ExportOptions(5, 0) = 1 Then
                        If PList(p, 7) <> "0" Then 'CRIT %
                            HTMLCodeNew = HTMLCodeNew & "<TD>" & Round((CDbl(PList(p, 7)) / CDbl(PList(p, 4))) * 100, 2) & "%</TD>" & vbNewLine
                        Else
                            HTMLCodeNew = HTMLCodeNew & "<TD>0%</TD>" & vbNewLine
                        End If
                    End If
                    If ExportOptions(6, 0) = 1 Then HTMLCodeNew = HTMLCodeNew & "<TD>" & PList(p, 7) & "</TD>" & vbNewLine                        'CRIT COUNT
                    If ExportOptions(7, 0) = 1 Then
                        If PList(p, 4) <> "0" Then 'HIT %
                            HTMLCodeNew = HTMLCodeNew & "<TD>" & Round((CDbl(PList(p, 4)) / CDbl(PList(p, 5))) * 100, 2) & "%</TD>" & vbNewLine
                        Else
                            HTMLCodeNew = HTMLCodeNew & "<TD>0%</TD>" & vbNewLine
                        End If
                    End If
                    If ExportOptions(8, 0) = 1 Then HTMLCodeNew = HTMLCodeNew & "<TD>" & PList(p, 4) & "/" & PList(p, 6) & "</TD>" & vbNewLine                        'HIT/MISS
                    If ExportOptions(9, 0) = 1 Then
                        If PList(p, 14) <> "0" Then 'Avoid %
                            HTMLCodeNew = HTMLCodeNew & "<TD>" & Round((CDbl(PList(p, 14)) / (CDbl(PList(p, 14)) + CDbl(PList(p, 15)))) * 100, 2) & "%</TD>" & vbNewLine
                        Else
                            HTMLCodeNew = HTMLCodeNew & "<TD>0%</TD>" & vbNewLine
                        End If
                    End If
                    If ExportOptions(10, 0) = 1 Then HTMLCodeNew = HTMLCodeNew & "<TD>" & PList(p, 15) & "/" & PList(p, 14) & "</TD>" & vbNewLine                        'TAKE/Avoid
                    If ExportOptions(11, 0) = 1 Then HTMLCodeNew = HTMLCodeNew & "<TD>" & PList(p, 19) & "</TD>" & vbNewLine                        'Evades
                    If ExportOptions(12, 0) = 1 Then HTMLCodeNew = HTMLCodeNew & "<TD>" & PList(p, 20) & "</TD>" & vbNewLine                        'Parries
                    If ExportOptions(13, 0) = 1 Then HTMLCodeNew = HTMLCodeNew & "<TD>" & PList(p, 21) & "</TD>" & vbNewLine                        'Blocks
                    If ExportOptions(23, 0) = 1 Then HTMLCodeNew = HTMLCodeNew & "<TD>" & PList(p, 28) & "</TD>" & vbNewLine                        'Counters
                    If ExportOptions(20, 0) = 1 Then HTMLCodeNew = HTMLCodeNew & "<TD>" & PList(p, 25) & "</TD>" & vbNewLine                        'Anti
                    If ExportOptions(14, 0) = 1 Then HTMLCodeNew = HTMLCodeNew & "<TD>" & PList(p, 22) & "</TD>" & vbNewLine                        'Absorbs
                    If ExportOptions(15, 0) = 1 Then HTMLCodeNew = HTMLCodeNew & "<TD>" & PList(p, 23) & "</TD>" & vbNewLine                        'Avoids
                    If ExportOptions(16, 0) = 1 Then HTMLCodeNew = HTMLCodeNew & "<TD>" & PList(p, 16) & "</TD>" & vbNewLine                        'DMG TAKEN
                    If ExportOptions(17, 0) = 1 Then HTMLCodeNew = HTMLCodeNew & "<TD>" & PList(p, 13) & "</TD>" & vbNewLine                        'HP REC'D
                    If ExportOptions(19, 0) = 1 Then HTMLCodeNew = HTMLCodeNew & "<TD>" & PList(p, 24) & "</TD>" & vbNewLine                        'HP Healed
                    HTMLCodeNew = HTMLCodeNew & "<TD BGCOLOR=""#b8ced9""><B>" & PList(p, 1) & "</b> <FONT style=""font-family:small fonts;font-size:6pt"">(" & PList(p, 2) & "%)</TD></TR>" & vbNewLine 'TOTAL AND % OF DMG
                End If
                For StatClear = 0 To 28
                    PList(p, StatClear) = ""
                Next
            End If
        Next
        If TotalDMG <> 0 And GenerateHTML And ExportOptions(18, 0) = 0 Then
            HTMLCodeNew = HTMLCodeNew & "<TR style=""BACKGROUND-COLOR:#7CB1CB"">" & vbNewLine
            HTMLCodeNew = HTMLCodeNew & "<TD><B>Totals</B></TD>" & vbNewLine
            If ExportOptions(0, 0) = 1 Then HTMLCodeNew = HTMLCodeNew & "<TD><B>" & Format(TotalBase, "#,###") & "</B> <FONT style=""font-family:small fonts;font-size:6pt""><br>(" & Round((TotalBase / TotalDMG) * 100, 2) & "%)</TD>" & vbNewLine                'TOTAL BASIC
            If ExportOptions(1, 0) = 1 Then HTMLCodeNew = HTMLCodeNew & "<TD><B>" & Format(TotalSkill, "#,###") & "</B> <FONT style=""font-family:small fonts;font-size:6pt""><br>(" & Round((TotalSkill / TotalDMG) * 100, 2) & "%)</TD>" & vbNewLine                'TOTAL SKILL
            If ExportOptions(2, 0) = 1 Then HTMLCodeNew = HTMLCodeNew & "<TD><B>" & Format(TotalSpell, "#,###") & "</B> <FONT style=""font-family:small fonts;font-size:6pt""><br>(" & Round((TotalSpell / TotalDMG) * 100, 2) & "%)</TD>" & vbNewLine                'TOTAL SPELL
            If ExportOptions(21, 0) = 1 Then HTMLCodeNew = HTMLCodeNew & "<TD><B>" & Format(TotalEffect, "#,###") & "</B> <FONT style=""font-family:small fonts;font-size:6pt""><br>(" & Round((TotalEffect / TotalDMG) * 100, 2) & "%)</TD>" & vbNewLine                'EFFECT
            If ExportOptions(22, 0) = 1 Then HTMLCodeNew = HTMLCodeNew & "<TD></TD>" & vbNewLine                'WS COUNT
            If ExportOptions(3, 0) = 1 Then HTMLCodeNew = HTMLCodeNew & "<TD></TD>" & vbNewLine                'High/Low
            If ExportOptions(4, 0) = 1 Then HTMLCodeNew = HTMLCodeNew & "<TD></TD>" & vbNewLine                'Average
            If ExportOptions(5, 0) = 1 Then HTMLCodeNew = HTMLCodeNew & "<TD></TD>" & vbNewLine                'CRIT %
            If ExportOptions(6, 0) = 1 Then HTMLCodeNew = HTMLCodeNew & "<TD></TD>" & vbNewLine                'CRIT COUNT
            If ExportOptions(7, 0) = 1 Then HTMLCodeNew = HTMLCodeNew & "<TD></TD>" & vbNewLine                'HIT %
            If ExportOptions(8, 0) = 1 Then HTMLCodeNew = HTMLCodeNew & "<TD></TD>" & vbNewLine                'HIT/MISS
            If ExportOptions(9, 0) = 1 Then HTMLCodeNew = HTMLCodeNew & "<TD></TD>" & vbNewLine                'Avoid %
            If ExportOptions(10, 0) = 1 Then HTMLCodeNew = HTMLCodeNew & "<TD></TD>" & vbNewLine                'TAKE/Avoid
            If ExportOptions(11, 0) = 1 Then HTMLCodeNew = HTMLCodeNew & "<TD></TD>" & vbNewLine                'Evades
            If ExportOptions(12, 0) = 1 Then HTMLCodeNew = HTMLCodeNew & "<TD></TD>" & vbNewLine                'Parries
            If ExportOptions(13, 0) = 1 Then HTMLCodeNew = HTMLCodeNew & "<TD></TD>" & vbNewLine                'Blocks
            If ExportOptions(23, 0) = 1 Then HTMLCodeNew = HTMLCodeNew & "<TD></TD>" & vbNewLine                'Counters
            If ExportOptions(20, 0) = 1 Then HTMLCodeNew = HTMLCodeNew & "<TD></TD>" & vbNewLine                'Anti
            If ExportOptions(14, 0) = 1 Then HTMLCodeNew = HTMLCodeNew & "<TD></TD>" & vbNewLine                'Absorbs
            If ExportOptions(15, 0) = 1 Then HTMLCodeNew = HTMLCodeNew & "<TD></TD>" & vbNewLine                'Avoids
            If ExportOptions(16, 0) = 1 Then HTMLCodeNew = HTMLCodeNew & "<TD><B>" & Format(TotalTaken, "#,###") & "</B></TD>" & vbNewLine                'TOTAL TAKEN
            If ExportOptions(17, 0) = 1 Then HTMLCodeNew = HTMLCodeNew & "<TD><B>" & Format(TotalHP, "#,###") & "</B></TD>" & vbNewLine                'TOTAL HP REC'D
            If ExportOptions(19, 0) = 1 Then HTMLCodeNew = HTMLCodeNew & "<TD><B>" & Format(TotalHPH, "#,###") & "</B></TD>" & vbNewLine                'TOTAL HP GIVEN
            HTMLCodeNew = HTMLCodeNew & "<TD><B>" & Format(TotalDMG, "#,###") & "</B></TD>" & vbNewLine 'TOTAL DMG DEALT
            HTMLCodeNew = HTMLCodeNew & "</TR>" & vbNewLine
            If FightComment <> "" Then
                HTMLCodeNew = HTMLCodeNew & "<TR><TH colSpan=25 align=""Left"" BGColor=""7CB1CB"">Comment: " & FightComment & "</TR>" & vbNewLine
            End If
            HTMLCodeNew = HTMLCodeNew & "</TABLE><P></CENTER>"
            HTMLCode = HTMLCodeNew & HTMLCode
        End If
  
        For i = 0 To UBound(Stats)
            If Stats(i, 0) <> "" Then
                If Stats(i, 20) = "" Then
                    Stats(i, 20) = "1"
                ElseIf CDbl(Stats(i, 20)) < 5 Then
                    Stats(i, 20) = CDbl(Stats(i, 20)) + 1
                Else
                    For StatClear = 0 To 30
                        Stats(i, StatClear) = ""
                    Next
                End If
            End If
        Next
        PrevTotalDMG = PrevTotalDMG + TotalDMG
        PrevSwings = PrevSwings + TotalSwing
        PrevTotalTaken = PrevTotalTaken + TotalTaken
        PrevTakenSwings = PrevTakenSwings + TotalSwingTaken
        EffList = EffList & CStr(UniqueMOB) & ", "
        UniqueMOB = UniqueMOB + 1
        TotalSwingTaken = 0
        TotalDMG = 0
        TotalTaken = 0
        TotalHP = 0
        TotalHPH = 0
        TotalHit = 0
        TotalSwing = 0
        TotalEffect = 0
        dHigh = 0
        dLow = 10000
        TotalBase = 0
        TotalSpell = 0
        TotalSkill = 0
        FightComment = ""
        MonsterCheck = ""
        CurrentFight = ""
    ElseIf LineType = "bf" Then
        If InStr(1, LCase(CurrentLine), "decent") Then
            MonsterCheck = "(DC)"
        ElseIf InStr(1, LCase(CurrentLine), "very tough") Then
            MonsterCheck = "(VT)"
        ElseIf InStr(1, LCase(CurrentLine), "incredibly tough") Then
            MonsterCheck = "(IT)"
        ElseIf InStr(1, LCase(CurrentLine), "tough") Then
            MonsterCheck = "(T)"
        ElseIf InStr(1, LCase(CurrentLine), "weak") Then
            MonsterCheck = "(TW)"
        ElseIf InStr(1, LCase(CurrentLine), "easy") Then
            MonsterCheck = "(EP)"
        ElseIf InStr(1, LCase(CurrentLine), "even") Then
            MonsterCheck = "(EM)"
        End If
    ElseIf InStr(1, "09,0a,01,02,0c,04,0d,05,0e,06", LineType) Then
          ChatText(UBound(ChatText)) = CurrentLine
          ReDim Preserve ChatText(UBound(ChatText) + 1)
    ElseIf (Left$(LCase(CurrentLine), 18) = "parser start dps") Then
        ReadDPS_Start = True
        BeginDPS = True
        If CustomMode = False Then
            Print #EditFile, CurrentLine
        End If
    ElseIf (Left$(LCase(CurrentLine), 18) = "parser start exp") Then
        ReadEXP_Start = True
        If CustomMode = False Then
            Print #EditFile, CurrentLine
        End If
    ElseIf (Left$(LCase(CurrentLine), 20) = "parser start fight") Then
        Read_Start = True
        If CustomMode = False Then
            Print #EditFile, CurrentLine
        End If
    ElseIf (Left$(LCase(CurrentLine), 19) = "parser stop fight") Then
        Read_Stop = True
        If CustomMode = False Then
            Print #EditFile, CurrentLine
        End If
    ElseIf (Left$(LCase(CurrentLine), 17) = "parser stop dps") Then
        BeginDPS = False
        ReadDPS_Stop = True
        If CustomMode = False Then
            Print #EditFile, CurrentLine
        End If
    ElseIf (Left$(LCase(CurrentLine), 18) = "parser clear dps") Then
        BeginDPS = False
        ReadDPS_Stop = False
        ReadDPS_Start = False
        Erase DPS
    ElseIf (Left$(LCase(CurrentLine), 17) = "parser stop exp") Then
        ReadEXP_Stop = True
        If CustomMode = False Then
            Print #EditFile, CurrentLine
        End If
    ElseIf (Left$(LCase(CurrentLine), 19) = "parser start fish") Then
        FishRPT
        Erase FishFound
        ReDim FishFound(0)
        ReadFISH_Start = True
        If CustomMode = False Then
            Print #EditFile, CurrentLine
        End If
    ElseIf (Left$(LCase(CurrentLine), 18) = "parser stop fish") Then
        FishRPT
        Erase FishFound
        ReDim FishFound(0)
        If CustomMode = False Then
            Print #EditFile, CurrentLine
        End If
    ElseIf InStr(1, CurrentLine, "Earth:") And LineType = "8c" Then
        If Read_Start Then
            EffList = ""
            PrevTotalDMG = 0
            PrevSwings = 0
            PrevTotalTaken = 0
            PrevTakenSwings = 0
            BeginDPS = True
            StartTimeDPS = Replace(Mid(CurrentLine, 10, Len(CurrentLine) - 13), ".", "")
            ReadDPS_Start = False
            StopDPS = False
            FightStartTime = Replace(Mid(CurrentLine, 10, Len(CurrentLine) - 13), ".", "")
            Read_Start = False
        ElseIf Read_Stop Then
            BeginDPS = False
            StopTimeDPS = Replace(Mid(CurrentLine, 10, Len(CurrentLine) - 13), ".", "")
            For i = 0 To UBound(DPS)
                If DPS(i, 2) = "" Then DPS(i, 2) = "0"
                DPS(i, 2) = CDbl(DPS(i, 2)) + CDbl(DateDiff("s", StartTimeDPS, StopTimeDPS))
                DPS(i, 0) = DPS(i, 0)
                DPS(i, 1) = DPS(i, 1)
                DPS(i, 2) = DPS(i, 2)
            Next
            ReadDPS_Stop = False
            StopDPS = True
            FightStopTime = Replace(Mid(CurrentLine, 10, Len(CurrentLine) - 13), ".", "")
            With RTB_Report
                TotalSeconds = DateDiff("s", FightStartTime, FightStopTime)
                .SelStart = Len(.Text) - 2
                .SelBold = True
                .SelText = vbNewLine & Mid(EffList, 1, Len(EffList) - 2) & " - Efficiency" & vbNewLine
                .SelBold = False
                .SelColor = &HC00000
                .SelText = "Time Taken: "
                .SelColor = &HC0&
                .SelText = Format(FightStopTime - FightStartTime & vbNewLine, "HH:MM:SS") & vbNewLine
                .SelColor = &HC00000
                .SelText = "Damage Per Second: "
                .SelColor = &HC0&
                .SelText = Round(PrevTotalDMG / TotalSeconds, 2) & vbNewLine
                .SelColor = &HC00000
                .SelText = "Attacks Per Second: "
                .SelColor = &HC0&
                .SelText = Round(PrevSwings / TotalSeconds, 2) & vbNewLine
                .SelColor = &HC00000
                .SelText = "Damage Taken Per Second: "
                .SelColor = &HC0&
                .SelText = Round(PrevTotalTaken / TotalSeconds, 2) & vbNewLine
                .SelColor = &HC00000
                .SelText = "Attacks Taken Per Second: "
                .SelColor = &HC0&
                .SelText = Round(PrevTakenSwings / TotalSeconds, 2) & vbNewLine
                .SelStart = Len(.Text)
            End With
            Read_Stop = False
        End If
        If ReadEXP_Start Then
            TotalExp = 0
            StartTime = Replace(Mid(CurrentLine, 10, Len(CurrentLine) - 13), ".", "")
            ReadEXP_Start = False
            StopEXP = False
        ElseIf ReadEXP_Stop Then
            StopTime = Replace(Mid(CurrentLine, 10, Len(CurrentLine) - 13), ".", "")
            ReadEXP_Stop = False
            StopEXP = True
        End If
        If ReadFISH_Start Then
            MyPos = InStr(1, PrevlineB, ",")
            If MyPos <> 0 Then
                FishHeader = Trim(Left(Mid(PrevlineB, MyPos + 2), Len(Mid(PrevlineB, MyPos + 2)) - 2))
                FishHeader = FishHeader & " - " & Left(Trim(Mid(PrevlineA, 3)), Len(Trim(Mid(PrevlineA, 3))) - 2)
                FishHeader = FishHeader & " - Earth: " & Trim(Mid(CurrentLine, 24, Len(CurrentLine) - 26))
            End If
            ReadFISH_Start = False
        End If
        If ReadDPS_Start Then
            StartTimeDPS = Replace(Mid(CurrentLine, 10, Len(CurrentLine) - 13), ".", "")
            ReadDPS_Start = False
            StopDPS = False
        ElseIf ReadDPS_Stop Then
            StopTimeDPS = Replace(Mid(CurrentLine, 10, Len(CurrentLine) - 13), ".", "")
            For i = 0 To UBound(DPS)
                If DPS(i, 2) = "" Then DPS(i, 2) = "0"
                DPS(i, 2) = CDbl(DPS(i, 2)) + CDbl(DateDiff("s", StartTimeDPS, StopTimeDPS))
                DPS(i, 0) = DPS(i, 0)
                DPS(i, 1) = DPS(i, 1)
                DPS(i, 2) = DPS(i, 2)
            Next
            ReadDPS_Stop = False
            StopDPS = True
        End If
        If StopEXP Or StopDPS Then
            RTB_Averages.Text = ""
            RTB_Averages.SelBold = True
            RTB_Averages.SelText = "Experience" & vbNewLine
            RTB_Averages.SelBold = False
            If TotalExp <> 0 And StartTime <> "12:00:00 AM" And StopTime <> "12:00:00 AM" Then
              RTB_Averages.SelText = "Start: " & StartTime & " / Stop: " & StopTime & vbNewLine & "Total Exp: " & TotalExp & vbNewLine & "Per Hour: " & Round(TotalExp / DateDiff("n", StartTime, StopTime), 2) * 60 & vbNewLine & "Per Minute: " & Round(TotalExp / DateDiff("n", StartTime, StopTime), 2) & vbNewLine & vbNewLine
            ElseIf TotalExp <> 0 And StartTime <> "12:00:00 AM" Then
              RTB_Averages.SelText = "Start: " & StartTime & " / Stop: " & Now & vbNewLine & "Total Exp: " & TotalExp & vbNewLine & "Per Hour: " & Round(TotalExp / DateDiff("n", StartTime, Now), 2) * 60 & vbNewLine & "Per Minute: " & Round(TotalExp / DateDiff("n", StartTime, Now), 2) & vbNewLine & vbNewLine
            Else
              RTB_Averages.SelText = "Start: " & StartTime & vbNewLine & "Total Exp: " & TotalExp & vbNewLine & "Per Hour: 0" & vbNewLine & "Per Minute: 0" & vbNewLine & vbNewLine
            End If

            For p = 0 To UBound(Players)
                If Players(p, 0) <> "" Then
                    With RTB_Averages
                        SelStart = Len(.Text)
                        .SelText = Players(p, 0)
                        .SelStart = SelStart
                        .SelLength = Len(Players(p, 0))
                        .SelBold = True
                        .SelStart = Len(.Text)
                        .SelBold = False
                        AvgAcc = Round(Players(p, 4) / Players(p, 1), 2)
                        EstDPS = ""
                        For dp = 0 To UBound(DPS)
                          If DPS(dp, 0) = Trim(Players(p, 0)) Then
                              If DPS(dp, 0) <> "" Then
                                    If DPS(dp, 1) <> "0" And DPS(dp, 2) <> "0" And DPS(dp, 2) <> "" And DPS(dp, 1) <> "" Then
                                        EstDPS = Round(DPS(dp, 1) / DPS(dp, 2), 2) & " (" & DPS(dp, 2) & " seconds / " & DPS(dp, 1) & " dmg)"
                                    Else
                                        EstDPS = 0
                                    End If
                                  Exit For
                              End If
                          End If
                        Next
                        If Players(p, 3) = "0" Or Players(p, 3) = "" Then Players(p, 3) = "1"
                        If Players(p, 2) = "0" Or Players(p, 2) = "" Then Players(p, 2) = "1"
                        .SelText = vbNewLine & ResizePart("Total Fights:", 1500) & vbTab & Players(p, 1) & vbNewLine & ResizePart("Average Damage:", 1500) & vbTab & Round(CDbl(Players(p, 2)) / CDbl(Players(p, 1)), 3) & vbNewLine & ResizePart("Average Percent:", 1500) & vbTab & Round(CDbl(Players(p, 3)) / CDbl(Players(p, 1)), 3) & vbNewLine & ResizePart("Average Accuracy:", 1500) & vbTab & AvgAcc & vbNewLine & ResizePart("Estimated DPS:", 1500) & vbTab & EstDPS & vbNewLine & vbNewLine
                    End With
                    'RTB_Averages.SelText = vbNewLine & "Total Fights: " & vbTab & Players(p, 1) & vbNewLine & "Average Damage: " & vbTab & Round(CDbl(Players(p, 2)) / CDbl(Players(p, 1)), 3) & vbNewLine & "Average Percent: " & vbTab & Round(CDbl(Players(p, 3)) / CDbl(Players(p, 1)), 3) & vbNewLine & "Average Accuracy: " & AvgAcc & vbTab & vbNewLine & "Estimated DPS: " & vbTab & EstDPS & vbNewLine & vbNewLine
                End If
            Next
            RTB_Averages.SelStart = 0
        End If
        If CustomMode = False Then
            Print #EditFile, CurrentLine
        End If
    ElseIf InStr(1, CurrentLine, "experience points.") And LineType = "79" Then
        MyPos = InStr(1, CurrentLine, "gains ")
        MyPos2 = InStr(1, CurrentLine, "exp")
        If StopEXP = False Then
            TotalExp = TotalExp + CDbl(Mid$(CurrentLine, MyPos + 6, MyPos2 - (MyPos + 7)))
        End If
        If CustomMode = False Then
            Print #EditFile, CurrentLine
        End If
    ElseIf (Left$(LCase(CurrentLine), 14) = "parser clear") And CustomMode = False Then
        Close #EditFile
        ClearEdit = True
        mnuClear_Click
        CustomMode = False
        ClearEdit = False
        EditFile = FreeFile
        Open App.Path & "\EditFile.log" For Append As #EditFile
    ElseIf (Left$(LCase(CurrentLine), 13) = "parser save") Then
      MyPos = InStr(1, CurrentLine, ".rtf")
      MyPos2 = InStrRev(CurrentLine, " ", MyPos)
      SaveFileName = Mid$(CurrentLine, MyPos2 + 1, (MyPos + 4) - (MyPos2 + 1))
      If InStr(1, LCase(CurrentLine), "save report") Then
          RTB_Report.SaveFile SaveFileName, rtfText
      ElseIf InStr(1, LCase(CurrentLine), "save player") Then
          If InStr(1, LCase(CurrentLine), "save player1") Then
              comboUser.ListIndex = 0
          ElseIf InStr(1, LCase(CurrentLine), "save player2") Then
              comboUser.ListIndex = 1
          ElseIf InStr(1, LCase(CurrentLine), "save player3") Then
              comboUser.ListIndex = 2
          ElseIf InStr(1, LCase(CurrentLine), "save player4") Then
              comboUser.ListIndex = 3
          ElseIf InStr(1, LCase(CurrentLine), "save player5") Then
              comboUser.ListIndex = 4
          ElseIf InStr(1, LCase(CurrentLine), "save player6") Then
              comboUser.ListIndex = 5
          End If
          comboUser_Click
          RTB_User.SaveFile SaveFileName, rtfText
      ElseIf InStr(1, LCase(CurrentLine), "save summary") Then
          RTB_Averages.SaveFile SaveFileName, rtfText
      ElseIf InStr(1, LCase(CurrentLine), "save details") Then
          RTB_Details.SaveFile SaveFileName, rtfText
      End If
    ElseIf (Left$(LCase(CurrentLine), 15) = "parser player") Then
      MyPos = InStr(1, CurrentLine, "'")
      MyPos2 = InStr(MyPos + 1, CurrentLine, "'")
      If MyPos <> 0 And MyPos2 <> 0 Then
        NewPlayer = Mid$(CurrentLine, MyPos + 1, MyPos2 - (MyPos + 1))
      Else
        NewPlayer = ""
      End If
      If Trim(NewPlayer) <> "" Then
          If InStr(1, LCase(CurrentLine), "player1") Then
              i = 0
              comboUser.List(0) = NewPlayer
              UserLog(0, 0) = NewPlayer
              SaveSetting App.Title, "Settings", "Player1", NewPlayer
              comboUser.ListIndex = 0
              comboUser_Click
          ElseIf InStr(1, LCase(CurrentLine), "player2") Then
              i = 1
              comboUser.List(1) = NewPlayer
              UserLog(1, 0) = NewPlayer
              SaveSetting App.Title, "Settings", "Player2", NewPlayer
              comboUser.ListIndex = 1
              comboUser_Click
          ElseIf InStr(1, LCase(CurrentLine), "player3") Then
              i = 2
              comboUser.List(2) = NewPlayer
              UserLog(2, 0) = NewPlayer
              SaveSetting App.Title, "Settings", "Player3", NewPlayer
              comboUser.ListIndex = 2
              comboUser_Click
          ElseIf InStr(1, LCase(CurrentLine), "player4") Then
              i = 3
              comboUser.List(3) = NewPlayer
              UserLog(3, 0) = NewPlayer
              SaveSetting App.Title, "Settings", "Player4", NewPlayer
              comboUser.ListIndex = 3
              comboUser_Click
          ElseIf InStr(1, LCase(CurrentLine), "player5") Then
              i = 4
              comboUser.List(4) = NewPlayer
              UserLog(4, 0) = NewPlayer
              SaveSetting App.Title, "Settings", "Player5", NewPlayer
              comboUser.ListIndex = 4
              comboUser_Click
          ElseIf InStr(1, LCase(CurrentLine), "player6") Then
              i = 5
              comboUser.List(5) = NewPlayer
              UserLog(5, 0) = NewPlayer
              SaveSetting App.Title, "Settings", "Player6", NewPlayer
              comboUser.ListIndex = 5
              comboUser_Click
          End If
          comboUser.List(i) = UserLog(i, 0)
          UserLog(i, 1) = ""
          UserLog(i, 2) = "0"
          UserLog(i, 3) = "0"
          UserLog(i, 4) = "0"
          UserLog(i, 5) = "0"
          UserLog(i, 6) = "0"
          UserLog(i, 7) = GetSetting(App.Title, "Settings", "PlayerOptions" & i + 1, Default:="1,0,0,0,0,1,1,1,0")
          UserLog(i, 8) = "0"
          UserLog(i, 9) = "0"
          UserLog(i, 10) = "0"
          UserLog(i, 11) = "0"
      End If
  ElseIf (Left$(LCase(CurrentLine), 18) = "parser comment '") Then
      MyPos = InStr(1, CurrentLine, "'")
      MyPos2 = InStr(MyPos + 1, CurrentLine, "'")
      If MyPos2 <> 0 Then
        FightComment = Mid$(CurrentLine, MyPos + 1, MyPos2 - (MyPos + 1))
      Else
        FightComment = "Invalid format."
      End If
      
      If CustomMode = False Then
          Print #EditFile, CurrentLine
      End If
  ElseIf (Left$(LCase(CurrentLine), 23) = "parser fish comment '") Then
      MyPos = InStr(1, CurrentLine, "'")
      MyPos2 = InStr(MyPos + 1, CurrentLine, "'")
      If MyPos2 <> 0 Then
        FishComment = Mid$(CurrentLine, MyPos + 1, MyPos2 - (MyPos + 1))
      Else
        FishComment = "Invalid format."
      End If
      If CustomMode = False Then
          Print #EditFile, CurrentLine
      End If
  ElseIf (Left$(CurrentLine, 12) = "yYou find") Then
      'LOOT
        MyPos = InStr(1, CurrentLine, " ")
        MyPos2 = InStr(MyPos + 1, CurrentLine, " ")
        LootItem = Mid$(CurrentLine, MyPos + 3, MyPos2 - (MyPos + 3))
        PrevLoot = False
        For lf = 0 To UBound(LootFound)
          If InStr(1, LootFound(lf), LootItem) Then
              PrevLoot = True
              Exit For
          End If
        Next
        If PrevLoot Then
          MyPos = InStr(1, LootFound(lf), " - ")
          LootFound(lf) = CDbl(Left$(LootFound(lf), MyPos)) + 1 & " - " & LootItem
        Else
          LootFound(UBound(LootFound)) = "1 - " & LootItem
          ReDim Preserve LootFound(UBound(LootFound) + 1)
        End If
  ElseIf LineType = "7f" And InStr(1, CurrentLine, " obtains ") And InStr(1, CurrentLine, " gil.") = 0 Then
      'PLAYER LOOT
        MyPos = InStr(1, CurrentLine, " ")
        LootPlayer = Mid(CurrentLine, 5, MyPos - 5)
        MyPos = InStr(1, CurrentLine, " ")
        MyPos2 = InStr(MyPos + 1, CurrentLine, ".")
        LootItem = Mid$(CurrentLine, MyPos + 3, MyPos2 - (MyPos + 3))
        PrevLoot = False
        For lf = 0 To UBound(PlayerLoot)
          If InStr(1, PlayerLoot(lf), LootItem & ";" & LootPlayer) Then
              PrevLoot = True
              Exit For
          End If
        Next
        If PrevLoot Then
          MyPos = InStr(1, PlayerLoot(lf), " - ")
          PlayerLoot(lf) = CDbl(Left$(PlayerLoot(lf), MyPos)) + 1 & " - " & LootItem & ";" & LootPlayer
        Else
          PlayerLoot(UBound(PlayerLoot)) = "1 - " & LootItem & ";" & LootPlayer
          ReDim Preserve PlayerLoot(UBound(PlayerLoot) + 1)
        End If
  ElseIf InStr(1, LCase(CurrentLine), " obtains ") <> 0 And InStr(1, LCase(CurrentLine), " gil.") <> 0 And Mid$(CurrentLine, 3, 1) <> "<" And Mid$(CurrentLine, 3, 1) <> ">" And Mid$(CurrentLine, 3, 1) <> "(" And InStr(1, CurrentLine, " : ") = 0 And LineType <> "0f" Then
      'GIL
        MyPos = InStr(1, CurrentLine, " obtains ")
        MyPos2 = InStr(1, CurrentLine, " gil.")
        LootItem = "Gil"
        GilAmt = CDbl(Mid$(CurrentLine, MyPos + 9, MyPos2 - (MyPos + 9)))
        PrevLoot = False
        For lf = 0 To UBound(LootFound)
          If InStr(1, LootFound(lf), LootItem) Then
              PrevLoot = True
              Exit For
          End If
        Next
        If PrevLoot Then
          MyPos = InStr(1, LootFound(lf), " - ")
          LootFound(lf) = CDbl(Left$(LootFound(lf), MyPos)) + GilAmt & " - " & LootItem
        Else
          LootFound(UBound(LootFound)) = GilAmt & " - " & LootItem
          ReDim Preserve LootFound(UBound(LootFound) + 1)
        End If
  ElseIf (InStr(1, LCase(CurrentLine), "obtained: ") <> 0 Or InStr(1, LCase(CurrentLine), "you lost your catch.") <> 0 Or InStr(1, LCase(CurrentLine), "you didn't catch anything.") <> 0) And Mid$(CurrentLine, 3, 1) <> "<" And Mid$(CurrentLine, 3, 1) <> ">" And Mid$(CurrentLine, 3, 1) <> "(" And InStr(1, CurrentLine, " : ") = 0 And LineType <> "0f" Then
   
    If InStr(1, LCase(CurrentLine), "obtained: ") And LineType = "94" Then
        MyPos = InStr(1, CurrentLine, "obtained: ")
        MyPos2 = InStr(1, CurrentLine, ".")
        FishItem = Mid$(CurrentLine, MyPos + 15, MyPos2 - (MyPos + 17))
        PrevFish = False
        For lf = 0 To UBound(FishFound)
          If InStr(1, FishFound(lf), FishItem) Then
              PrevFish = True
              Exit For
          End If
        Next
        If PrevFish Then
          MyPos = InStr(1, FishFound(lf), " - ")
          FishFound(lf) = CDbl(Left$(FishFound(lf), MyPos)) + 1 & " - " & FishItem
        Else
          FishFound(UBound(FishFound)) = "1 - " & FishItem
          ReDim Preserve FishFound(UBound(FishFound) + 1)
        End If
    ElseIf InStr(1, LCase(CurrentLine), "you lost your catch.") Then
        FishItem = "catches lost"
        PrevFish = False
        For lf = 0 To UBound(FishFound)
          If InStr(1, FishFound(lf), FishItem) Then
              PrevFish = True
              Exit For
          End If
        Next
        If PrevFish Then
          MyPos = InStr(1, FishFound(lf), " - ")
          FishFound(lf) = CDbl(Left$(FishFound(lf), MyPos)) + 1 & " - " & FishItem
        Else
          FishFound(UBound(FishFound)) = "1 - " & FishItem
          ReDim Preserve FishFound(UBound(FishFound) + 1)
        End If
    ElseIf InStr(1, LCase(CurrentLine), "you didn't catch anything.") Then
        FishItem = "didn't catch anything"
        PrevFish = False
        For lf = 0 To UBound(FishFound)
          If InStr(1, FishFound(lf), FishItem) Then
              PrevFish = True
              Exit For
          End If
        Next
        If PrevFish Then
          MyPos = InStr(1, FishFound(lf), " - ")
          FishFound(lf) = CDbl(Left$(FishFound(lf), MyPos)) + 1 & " - " & FishItem
        Else
          FishFound(UBound(FishFound)) = "1 - " & FishItem
          ReDim Preserve FishFound(UBound(FishFound) + 1)
        End If
    End If
  End If
Next
Done:
If CustomMode = False Then
    Close #EditFile
End If
If GenerateHTML Then
    SummaryHTML HTMLCode
End If
Exit Sub

err_handler:
ErrorCount = ErrorCount + 1
f = FreeFile
Dim ReportError As String
Open App.Path & "\error_log.txt" For Append As f
    ReportError = ReportError & vbNewLine & "Error: " & Err.Number & vbNewLine & "Description: " & Err.Description & vbNewLine & "Line: " & Erl & vbNewLine & "FFXI Log Line: " & CurrentLine & vbNewLine & "Previous FFXI Log Line: " & PrevlineA & vbNewLine & vbNewLine
    Write #f, ReportError
Close #f
HasErrors = True
Err.Clear
If ErrorCount >= 25 Then
    lblStatus.Caption = "Too many errors - Parsing stopped for this log."
    Exit Sub
Else
    Resume Next
End If
End Sub

Private Sub ResetUsers()
Dim i
For i = 0 To UBound(UserLog)
    UserLog(i, 0) = GetSetting(App.Title, "Settings", "Player" & i + 1, Default:="Player " & i + 1)
    comboUser.List(i) = UserLog(i, 0)
    UserLog(i, 1) = ""
    UserLog(i, 2) = "0"
    UserLog(i, 3) = "0"
    UserLog(i, 4) = "0"
    UserLog(i, 5) = "0"
    UserLog(i, 6) = "0"
    UserLog(i, 7) = GetSetting(App.Title, "Settings", "PlayerOptions" & i + 1, Default:="1,0,0,0,0,1,1,1,0")
    UserLog(i, 8) = "0"
    UserLog(i, 9) = "0"
    UserLog(i, 10) = "0"
    UserLog(i, 11) = "0"
Next
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
Public Function ResizeSimplePart(Part As String, Size As Long) As String
Do Until Len(Part) >= Size
    Part = Part & " "
Loop
ResizeSimplePart = Part
End Function
Public Sub StartNew()
10    On Error GoTo err_handler
20    If fileList.ListCount <> 0 Then
          Dim FSO As FileSystemObject
          Dim fo As Integer
          Dim MyDate As Date
          Dim FullFile() As String, LineType As String, CurrentFile As String
          Dim Index As Long
30        dLow = 10000

40        Set FSO = New FileSystemObject
50        If FSO.FolderExists(App.Path & "\FFXI_Logs") = False Then
60            FSO.CreateFolder App.Path & "\FFXI_Logs"
70        End If
80        If FSO.FolderExists(App.Path & "\FFXI_Gather") = False Then
90            FSO.CreateFolder App.Path & "\FFXI_Gather"
100       End If
110       If FSO.FileExists(App.Path & "\EditFile.log") = True Then
120           FSO.DeleteFile (App.Path & "\EditFile.log")
130       End If
140       lblStatus.Caption = "Errors: " & HasErrors & " - " & "Parsing Data..."
150       DoEvents
160       fileListBox.Clear

170       If Gather = True Then
180           If FSO.FileExists(SingleFile) = True Then
190               FSO.DeleteFile (SingleFile)
200           End If
              Dim EditFile
210           EditFile = FreeFile
220           Open SingleFile For Append As #EditFile
230       End If
    
240       If OpenSingle = False Then
250           For i = 0 To fileList.ListCount - 1
260               fileList.ListIndex = i
270               fileListBox.AddItem Format(FileDateTime(dirList.Path & "\" & fileList.FileName), "MM/DD HhNnSs") & " - " & fileList.Path & "\" & fileList.FileName
280           Next
290       End If
    
300       If OpenSingle Then
310           Erase FullFile
320           Index = 0
330           f = FreeFile
340           Open SingleFile For Input As f
350             Do Until EOF(f)
360               Line Input #f, CurrentLine
370               ReDim Preserve FullFile(Index)
380               FullFile(Index) = CurrentLine
390               Index = Index + 1
400             Loop
410           Close #f
420           If Index <> 0 Then
430             ParseLog FullFile, False, False
440           End If
450           cmdRecalc.Enabled = True
460           cmdExport.Enabled = True
470           If optionResults(1).Value = True Then
480               comboUser_Click
490           Else
500               comboDisplay_Click
510           End If
520           lblStatus.Caption = "Errors: " & HasErrors & " - " & "Finished Parsing Data."
530           Exit Sub
540       Else
550           Erase FullFile
560           Index = 0
570           For fo = 0 To fileListBox.ListCount - 1
580             fileListBox.ListIndex = fo
590             f = FreeFile
    
600             RTB.LoadFile Mid$(fileListBox.Text, 16)
610             RTB.Text = Mid(RTB.Text, 101)
620             RTB.Text = Replace(RTB.Text, Chr(0), vbNewLine)
630             MyPos = InStrRev(fileListBox.Text, "\")
640             If Gather = False Then
650               CurrentFile = App.Path & "\FFXI_Logs" & Mid(fileListBox.Text, MyPos)
660             Else
670               CurrentFile = App.Path & "\FFXI_Gather" & Mid(fileListBox.Text, MyPos)
680             End If
690             RTB.SaveFile CurrentFile, rtfText
    
700             MyDate = Left$(fileListBox.Text, 5) & Format(Date, "/YYYY") & " " & Format(Format(Mid$(fileListBox.Text, 7, 6), "00:00:00"), "Hh:Nn:Ss AM/PM")
710             ResetTimeFile CurrentFile, MyDate
720             Open CurrentFile For Input As f
730               Do Until EOF(f)
740                   Line Input #f, CurrentLine
750                   LineType = Left(CurrentLine, 2)
760                   If Mid(CurrentLine, 51, 2) = "01" And Index <> 0 Then
770                       FullFile(Index - 1) = Left(FullFile(Index - 1), Len(FullFile(Index - 1)) - 3) & Mid(CurrentLine, 56) & " " & LineType
780                   Else
790                       ReDim Preserve FullFile(Index)
800                       FullFile(Index) = Mid(CurrentLine, 54) & " " & LineType
810                       Index = Index + 1
820                   End If
830               Loop
840             Close #f
850           Next
860           If Gather = False Then
870             If Index <> 0 Then
880               ParseLog FullFile, False, False
890             End If
900             cmdRecalc.Enabled = True
910             cmdExport.Enabled = True
920           Else
930               If Index <> 0 Then
940                   For fo = 0 To UBound(FullFile)
950                       Print #EditFile, FullFile(fo)
960                   Next
970               End If
980           End If
990       End If
    
1000      fileListBox.ListIndex = fileListBox.ListCount - 1
1010      LastItem = fileListBox.Text
1020      timerRead.Enabled = True
1030      cmdOpen.Caption = "Stop"
1040      If optionResults(1).Value = True Then
1050    comboUser_Click
1060      Else
1070    comboDisplay_Click
1080      End If
1090      lblStatus.Caption = "Errors: " & HasErrors & " - " & "Waiting for new log...."
1100      If Gather = True Then
1110    Close #EditFile
1120      End If
1130  Else
1140      MsgBox "No log files found in this folder. Please select another folder." & vbNewLine & vbNewLine & "Usually: ""C:\Program Files\PlayOnline\SquareEnix\FINAL FANTASY XI\TEMP""", vbInformation, "Error"
1150      cmdOpen.Caption = "Start"
1160      lblStatus.Caption = "Errors: " & HasErrors & " - " & "Waiting."
1170      timerRead.Enabled = False
1180  End If

1190  Set FSO = Nothing
1200  Exit Sub

err_handler:
1210  MsgBox "Error: " & Err.Number & vbNewLine & "Description: " & Err.Description & vbNewLine & "Line: " & Erl, vbOKOnly, "Error"
1220  Err.Clear
1230  Exit Sub
End Sub

Private Sub SummaryHTML(HTMLCode As String)
Dim TotalDMG As Long, TotalBase As Long, TotalSkill As Long, TotalSpell As Long, TotalTaken As Long, TotalHP As Long, TotalHPH As Long, TotalEffect As Long
Dim HTMLFile, TotalFights As Long

HTMLFile = FreeFile
Open App.Path & "\" & ExportFile For Output As HTMLFile
    
For i = 0 To listResults.ListCount - 1
    If listResults.Selected(i) Then
        TotalFights = TotalFights + 1
    End If
Next

SummaryCode = SummaryCode & "<style type=""text/css"">"
SummaryCode = SummaryCode & "TD {BORDER-RIGHT: #7CB1CB 1px solid; BORDER-TOP: #7CB1CB 1px solid; BORDER-LEFT: #7CB1CB 1px solid; BORDER-BOTTOM: #7CB1CB 1px solid}"
SummaryCode = SummaryCode & "</style>"

SummaryCode = SummaryCode & "<CENTER><TABLE BORDER=1 CELLPADDING=0 CELLSPACING=0 style=""PADDING-LEFT: 3px;PADDING-RIGHT: 3px;BORDER-COLLAPSE:collapse;font-family:verdana;font-size:7pt;color:black;BORDER-RIGHT: #7CB1CB 1px solid; BORDER-TOP: #7CB1CB 1px solid; BORDER-LEFT: #7CB1CB 1px solid; BORDER-BOTTOM: #7CB1CB 1px solid"">" & vbNewLine
SummaryCode = SummaryCode & "<TR><TH colSpan=25 align=""Left"" BGColor=""7CB1CB"">Summary - " & TotalFights & " battles.</font></TH></TR>" & vbNewLine
SummaryCode = SummaryCode & "<TR style=""FONT-WEIGHT:bold;BACKGROUND-COLOR:#b8ced9"">"
SummaryCode = SummaryCode & "<TD WIDTH=75></TD>"
If ExportOptions(0, 0) = 1 Then SummaryCode = SummaryCode & "<TD WIDTH=50>Basic</TD>"
If ExportOptions(1, 0) = 1 Then SummaryCode = SummaryCode & "<TD WIDTH=50>Skill</TD>"
If ExportOptions(2, 0) = 1 Then SummaryCode = SummaryCode & "<TD WIDTH=50>Spell</TD>"
If ExportOptions(21, 0) = 1 Then SummaryCode = SummaryCode & "<TD WIDTH=50>Effect</TD>"
If ExportOptions(22, 0) = 1 Then SummaryCode = SummaryCode & "<TD WIDTH=50>WS #</TD>"
If ExportOptions(3, 0) = 1 Then SummaryCode = SummaryCode & "<TD WIDTH=50>High/Low</TD>"
If ExportOptions(4, 0) = 1 Then SummaryCode = SummaryCode & "<TD WIDTH=50>Average</TD>"
If ExportOptions(5, 0) = 1 Then SummaryCode = SummaryCode & "<TD WIDTH=50>Crit %</TD>"
If ExportOptions(6, 0) = 1 Then SummaryCode = SummaryCode & "<TD WIDTH=50>Crits</TD>"
If ExportOptions(7, 0) = 1 Then SummaryCode = SummaryCode & "<TD WIDTH=50>Hit %</TD>"
If ExportOptions(8, 0) = 1 Then SummaryCode = SummaryCode & "<TD WIDTH=50>Hit/Miss</TD>"
If ExportOptions(9, 0) = 1 Then SummaryCode = SummaryCode & "<TD WIDTH=50>Avoid %</TD>"
If ExportOptions(10, 0) = 1 Then SummaryCode = SummaryCode & "<TD WIDTH=50>Take/Avoid</TD>"
If ExportOptions(11, 0) = 1 Then SummaryCode = SummaryCode & "<TD WIDTH=50>Evades</TD>"
If ExportOptions(12, 0) = 1 Then SummaryCode = SummaryCode & "<TD WIDTH=50>Parries</TD>"
If ExportOptions(13, 0) = 1 Then SummaryCode = SummaryCode & "<TD WIDTH=50>Blocks</TD>"
If ExportOptions(23, 0) = 1 Then SummaryCode = SummaryCode & "<TD WIDTH=50>Counters</TD>"
If ExportOptions(20, 0) = 1 Then SummaryCode = SummaryCode & "<TD WIDTH=50>Anticipates</TD>"
If ExportOptions(14, 0) = 1 Then SummaryCode = SummaryCode & "<TD WIDTH=50>Absorbs</TD>"
If ExportOptions(15, 0) = 1 Then SummaryCode = SummaryCode & "<TD WIDTH=50>Avoids</TD>"
If ExportOptions(16, 0) = 1 Then SummaryCode = SummaryCode & "<TD WIDTH=75>DMG Taken</TD>"
If ExportOptions(17, 0) = 1 Then SummaryCode = SummaryCode & "<TD WIDTH=75>HP Rec'd</TD>"
If ExportOptions(19, 0) = 1 Then SummaryCode = SummaryCode & "<TD WIDTH=75>HP Healed</TD>"
SummaryCode = SummaryCode & "<TD WIDTH=150>TTL DMG</TD></TR>" & vbNewLine
For p = 0 To UBound(GrandPList)
    If GrandPList(p, 0) <> "" Then
        TotalDMG = TotalDMG + CDbl(GrandPList(p, 9)) + CDbl(GrandPList(p, 10)) + CDbl(GrandPList(p, 11)) + CDbl(GrandPList(p, 12))
    End If
Next
For p = 0 To UBound(GrandPList)
    If GrandPList(p, 0) <> "" Then
        SummaryCode = SummaryCode & "<TR style=""BACKGROUND-COLOR:#dae6ec"">" & vbNewLine
        SummaryCode = SummaryCode & "<TD BGCOLOR=""#b8ced9""><b>" & Replace(Trim$(GrandPList(p, 0)), "Skillchain: ", "SC:") & "</b></TD>" & vbNewLine 'PLAYER NAME
        If ExportOptions(0, 0) = 1 Then SummaryCode = SummaryCode & "<TD>" & CDbl(GrandPList(p, 9)) + CDbl(GrandPList(p, 10)) & "</TD>" & vbNewLine                'BASIC DMG
        If ExportOptions(1, 0) = 1 Then SummaryCode = SummaryCode & "<TD>" & GrandPList(p, 11) & "</TD>" & vbNewLine                'SKILL DMG
        If ExportOptions(2, 0) = 1 Then SummaryCode = SummaryCode & "<TD>" & GrandPList(p, 12) & "</TD>" & vbNewLine                'SPELL DMG
        If ExportOptions(21, 0) = 1 Then SummaryCode = SummaryCode & "<TD>" & GrandPList(p, 26) & "</TD>" & vbNewLine                'EFFECT
        If ExportOptions(22, 0) = 1 Then SummaryCode = SummaryCode & "<TD>" & GrandPList(p, 27) & "</TD>" & vbNewLine                'WS #
        If ExportOptions(3, 0) = 1 Then SummaryCode = SummaryCode & "<TD>" & GrandPList(p, 17) & "/" & GrandPList(p, 18) & "</TD>" & vbNewLine                'High/Low
        If ExportOptions(4, 0) = 1 Then
            If GrandPList(p, 5) <> 0 And (CDbl(GrandPList(p, 9)) + CDbl(GrandPList(p, 10))) <> 0 Then 'Average
                SummaryCode = SummaryCode & "<TD>" & Round((CDbl(GrandPList(p, 9)) + CDbl(GrandPList(p, 10))) / GrandPList(p, 4), 2) & "</TD>" & vbNewLine
            Else
                SummaryCode = SummaryCode & "<TD>0</TD>" & vbNewLine
            End If
        End If
        If ExportOptions(5, 0) = 1 Then
            If GrandPList(p, 7) <> "0" Then 'CRIT %
                SummaryCode = SummaryCode & "<TD>" & Round((CDbl(GrandPList(p, 7)) / CDbl(GrandPList(p, 4))) * 100, 2) & "%</TD>" & vbNewLine
            Else
                SummaryCode = SummaryCode & "<TD>0%</TD>" & vbNewLine
            End If
        End If
        If ExportOptions(6, 0) = 1 Then SummaryCode = SummaryCode & "<TD>" & GrandPList(p, 7) & "</TD>" & vbNewLine                'CRIT COUNT
        If ExportOptions(7, 0) = 1 Then
            If GrandPList(p, 4) <> "0" Then 'HIT %
                SummaryCode = SummaryCode & "<TD>" & Round((CDbl(GrandPList(p, 4)) / CDbl(GrandPList(p, 5))) * 100, 2) & "%</TD>" & vbNewLine
            Else
                SummaryCode = SummaryCode & "<TD>0%</TD>" & vbNewLine
            End If
        End If
        If ExportOptions(8, 0) = 1 Then SummaryCode = SummaryCode & "<TD>" & GrandPList(p, 4) & "/" & GrandPList(p, 6) & "</TD>" & vbNewLine                'HIT/MISS
        If ExportOptions(9, 0) = 1 Then
            If GrandPList(p, 14) <> "0" Then 'Avoid %
                SummaryCode = SummaryCode & "<TD>" & Round((CDbl(GrandPList(p, 14)) / (CDbl(GrandPList(p, 14)) + CDbl(GrandPList(p, 15)))) * 100, 2) & "%</TD>" & vbNewLine
            Else
                SummaryCode = SummaryCode & "<TD>0%</TD>" & vbNewLine
            End If
        End If
        If ExportOptions(10, 0) = 1 Then SummaryCode = SummaryCode & "<TD>" & GrandPList(p, 15) & "/" & GrandPList(p, 14) & "</TD>" & vbNewLine                'TAKE/Avoid
        If ExportOptions(11, 0) = 1 Then SummaryCode = SummaryCode & "<TD>" & GrandPList(p, 19) & "</TD>" & vbNewLine                'Evades
        If ExportOptions(12, 0) = 1 Then SummaryCode = SummaryCode & "<TD>" & GrandPList(p, 20) & "</TD>" & vbNewLine                'Parries
        If ExportOptions(13, 0) = 1 Then SummaryCode = SummaryCode & "<TD>" & GrandPList(p, 21) & "</TD>" & vbNewLine                'Blocks
        If ExportOptions(23, 0) = 1 Then SummaryCode = SummaryCode & "<TD>" & GrandPList(p, 28) & "</TD>" & vbNewLine                'Counters
        If ExportOptions(20, 0) = 1 Then SummaryCode = SummaryCode & "<TD>" & GrandPList(p, 25) & "</TD>" & vbNewLine                'Anti
        If ExportOptions(14, 0) = 1 Then SummaryCode = SummaryCode & "<TD>" & GrandPList(p, 22) & "</TD>" & vbNewLine                'Absorbs
        If ExportOptions(15, 0) = 1 Then SummaryCode = SummaryCode & "<TD>" & GrandPList(p, 23) & "</TD>" & vbNewLine                'Avoids
        If ExportOptions(16, 0) = 1 Then SummaryCode = SummaryCode & "<TD>" & GrandPList(p, 16) & "</TD>" & vbNewLine                'DMG TAKEN
        If ExportOptions(17, 0) = 1 Then SummaryCode = SummaryCode & "<TD>" & GrandPList(p, 13) & "</TD>" & vbNewLine                'HP REC'D
        If ExportOptions(19, 0) = 1 Then SummaryCode = SummaryCode & "<TD>" & GrandPList(p, 24) & "</TD>" & vbNewLine                'HP healed
        SummaryCode = SummaryCode & "<TD BGCOLOR=""#b8ced9""><B>" & GrandPList(p, 1) & "</b> <FONT style=""font-family:small fonts;font-size:6pt"">(" & Round(((GrandPList(p, 1) / TotalDMG) * 100), 2) & "%)</TD></TR>" & vbNewLine 'TOTAL AND % OF DMG
        TotalBase = TotalBase + CDbl(GrandPList(p, 9)) + CDbl(GrandPList(p, 10))
        TotalSpell = TotalSpell + CDbl(GrandPList(p, 12))
        TotalSkill = TotalSkill + CDbl(GrandPList(p, 11))
        TotalTaken = TotalTaken + CDbl(GrandPList(p, 16))
        TotalEffect = TotalEffect + CDbl(GrandPList(p, 26))
        TotalHP = TotalHP + CDbl(GrandPList(p, 13))
        TotalHPH = TotalHPH + CDbl(GrandPList(p, 24))
        GrandPList(p, 0) = ""
        GrandPList(p, 1) = ""
        GrandPList(p, 2) = ""
        GrandPList(p, 3) = ""
        GrandPList(p, 4) = ""
        GrandPList(p, 5) = ""
        GrandPList(p, 6) = ""
        GrandPList(p, 7) = ""
        GrandPList(p, 8) = ""
        GrandPList(p, 9) = ""
        GrandPList(p, 10) = ""
        GrandPList(p, 11) = ""
        GrandPList(p, 12) = ""
        GrandPList(p, 13) = ""
        GrandPList(p, 14) = ""
        GrandPList(p, 15) = ""
        GrandPList(p, 16) = ""
        GrandPList(p, 17) = ""
        GrandPList(p, 18) = ""
        GrandPList(p, 19) = ""
        GrandPList(p, 20) = ""
        GrandPList(p, 21) = ""
        GrandPList(p, 22) = ""
        GrandPList(p, 23) = ""
        GrandPList(p, 24) = ""
        GrandPList(p, 25) = ""
        GrandPList(p, 26) = ""
        GrandPList(p, 27) = ""
        GrandPList(p, 28) = ""
    End If
Next
SummaryCode = SummaryCode & "<TR style=""BACKGROUND-COLOR:#7CB1CB"">" & vbNewLine
SummaryCode = SummaryCode & "<TD><B>Totals</B></TD>" & vbNewLine
If ExportOptions(0, 0) = 1 Then SummaryCode = SummaryCode & "<TD><B>" & Format(TotalBase, "#,###") & "</B> <FONT style=""font-family:small fonts;font-size:6pt""><br>(" & Round((TotalBase / TotalDMG) * 100, 2) & "%)</TD>" & vbNewLine        'TOTAL BASIC
If ExportOptions(1, 0) = 1 Then SummaryCode = SummaryCode & "<TD><B>" & Format(TotalSkill, "#,###") & "</B> <FONT style=""font-family:small fonts;font-size:6pt""><br>(" & Round((TotalSkill / TotalDMG) * 100, 2) & "%)</TD>" & vbNewLine        'TOTAL SKILL
If ExportOptions(2, 0) = 1 Then SummaryCode = SummaryCode & "<TD><B>" & Format(TotalSpell, "#,###") & "</B> <FONT style=""font-family:small fonts;font-size:6pt""><br>(" & Round((TotalSpell / TotalDMG) * 100, 2) & "%)</TD>" & vbNewLine        'TOTAL SPELL
If ExportOptions(21, 0) = 1 Then SummaryCode = SummaryCode & "<TD><B>" & Format(TotalEffect, "#,###") & "</B> <FONT style=""font-family:small fonts;font-size:6pt""><br>(" & Round((TotalEffect / TotalDMG) * 100, 2) & "%)</TD>" & vbNewLine        'TOTAL EFFECT
If ExportOptions(22, 0) = 1 Then SummaryCode = SummaryCode & "<TD></TD>" & vbNewLine        'WS COUNT
If ExportOptions(3, 0) = 1 Then SummaryCode = SummaryCode & "<TD></TD>" & vbNewLine        'High/Low
If ExportOptions(4, 0) = 1 Then SummaryCode = SummaryCode & "<TD></TD>" & vbNewLine        'Average
If ExportOptions(5, 0) = 1 Then SummaryCode = SummaryCode & "<TD></TD>" & vbNewLine       'CRIT %
If ExportOptions(6, 0) = 1 Then SummaryCode = SummaryCode & "<TD></TD>" & vbNewLine        'CRIT COUNT
If ExportOptions(7, 0) = 1 Then SummaryCode = SummaryCode & "<TD></TD>" & vbNewLine        'HIT %
If ExportOptions(8, 0) = 1 Then SummaryCode = SummaryCode & "<TD></TD>" & vbNewLine        'HIT/MISS
If ExportOptions(9, 0) = 1 Then SummaryCode = SummaryCode & "<TD></TD>" & vbNewLine        'Avoid %
If ExportOptions(10, 0) = 1 Then SummaryCode = SummaryCode & "<TD></TD>" & vbNewLine        'TAKE/Avoid
If ExportOptions(11, 0) = 1 Then SummaryCode = SummaryCode & "<TD></TD>" & vbNewLine        'Evades
If ExportOptions(12, 0) = 1 Then SummaryCode = SummaryCode & "<TD></TD>" & vbNewLine        'Parries
If ExportOptions(13, 0) = 1 Then SummaryCode = SummaryCode & "<TD></TD>" & vbNewLine        'Blocks
If ExportOptions(23, 0) = 1 Then SummaryCode = SummaryCode & "<TD></TD>" & vbNewLine        'Counters
If ExportOptions(20, 0) = 1 Then SummaryCode = SummaryCode & "<TD></TD>" & vbNewLine        'Anti
If ExportOptions(14, 0) = 1 Then SummaryCode = SummaryCode & "<TD></TD>" & vbNewLine        'Absorbs
If ExportOptions(15, 0) = 1 Then SummaryCode = SummaryCode & "<TD></TD>" & vbNewLine        'Avoids
If ExportOptions(16, 0) = 1 Then SummaryCode = SummaryCode & "<TD><B>" & Format(TotalTaken, "#,###") & "</B></TD>" & vbNewLine        'TOTAL TAKEN
If ExportOptions(17, 0) = 1 Then SummaryCode = SummaryCode & "<TD><B>" & Format(TotalHP, "#,###") & "</B></TD>" & vbNewLine        'TOTAL HP REC'D
If ExportOptions(19, 0) = 1 Then SummaryCode = SummaryCode & "<TD><B>" & Format(TotalHPH, "#,###") & "</B></TD>" & vbNewLine        'TOTAL given
SummaryCode = SummaryCode & "<TD><B>" & Format(TotalDMG, "#,###") & "</B></TD>" & vbNewLine 'TOTAL DMG DEALT
SummaryCode = SummaryCode & "</TR>" & vbNewLine
SummaryCode = SummaryCode & "</TABLE><P></CENTER>"
HTMLCode = SummaryCode & HTMLCode
Print #HTMLFile, HTMLCode
Close #HTMLFile
End Sub

Private Sub cmdCustom_Click()
Dim i
For i = 0 To listResults.ListCount - 1
    If InStr(1, LCase(listResults.List(i)), LCase(comboMOB.Text)) Then
        listResults.Selected(i) = True
    Else
        listResults.Selected(i) = False
    End If
Next
End Sub

Private Sub cmdExport_Click()
frmExport.Show
End Sub

Private Sub cmdOpen_Click()
If cmdOpen.Caption = "Start" Then
    frmOpen.Left = Me.Left + 200
    frmOpen.Top = Me.Top + 200
    frmOpen.Visible = True
ElseIf cmdOpen.Caption = "Stop" Then
    FishRPT
    Erase FishFound
    ReDim FishFound(0)
    timerRead.Enabled = False
    cmdOpen.Caption = "Start"
    If Gather Then
        lblStatus.Caption = "Stopped - File saved as '" & SingleFile & "'"
    Else
        lblStatus.Caption = "Stopped - Waiting."
    End If
End If
End Sub





Private Sub cmdOptions_Click()
cmdOpen.SetFocus
PopupMenu mnuOptions
End Sub


Private Sub cmdRecalc_Click()
Screen.MousePointer = vbHourglass
ClearEdit = False
NotClearFile = True
mnuClear_Click
NotClearFile = False
ClearEdit = True
Dim FullFile() As String, CurrentLine As String, f, Index As Long, LineType As String, PrevLine As String
f = FreeFile
Erase FullFile
Index = 0

Open App.Path & "\EditFile.log" For Input As f
  Do Until EOF(f)
      Input #f, CurrentLine
      If Len(CurrentLine) = 2 And LineType = "" Then
          LineType = Right$(PrevLine, 2)
      End If
      If Trim(CurrentLine) <> "" And Left$(CurrentLine, 2) = "" Or Left$(CurrentLine, 3) = "but" Then
          If Left$(CurrentLine, 3) = "but" Then
              FullFile(Index - 1) = Trim(Left$(FullFile(Index - 1), Len(FullFile(Index - 1)) - 2)) & ", " & CurrentLine & " " & LineType
              LineType = ""
          Else
              ReDim Preserve FullFile(Index)
              FullFile(Index) = CurrentLine & " " & LineType
              Index = Index + 1
              LineType = ""
          End If
      End If
      PrevLine = CurrentLine
  Loop
Close #f

ParseLog FullFile, True, False
MsgBox "Done.", vbInformation, "Recalculate"
Screen.MousePointer = vbDefault
End Sub

Private Sub cmdSelect_Click()
Dim i
For i = 0 To listResults.ListCount - 1
    listResults.Selected(i) = True
Next
End Sub

Private Sub cmdUnselect_Click()
Dim i
For i = 0 To listResults.ListCount - 1
    listResults.Selected(i) = False
Next
End Sub





Private Sub comboDisplay_Click()
Dim i, lf, AddLoot As String, Players() As String, PlayerName As String, FoundPlayer As Boolean, PlayerCount As Long
ReDim Players(0)
optionResults(0).Value = True
If comboDisplay.Text = "Report" Then
    frameEdit.Visible = False
    RTB_Report.Visible = True
    RTB_User.Visible = False
    RTB_Averages.Visible = False
    RTB_Details.Visible = False
    frameChat.Visible = False
    RTB_Fish.Visible = False
ElseIf comboDisplay.Text = "Fishing" Then
    If RTB_Fish.Text = "" Then
        FishRPT
    End If
    frameEdit.Visible = False
    RTB_Report.Visible = False
    RTB_User.Visible = False
    RTB_Averages.Visible = False
    RTB_Details.Visible = False
    RTB_Fish.Visible = True
    frameChat.Visible = False
ElseIf comboDisplay.Text = "Chat" Then
    RTB_Chat.Text = ""
    For i = 0 To optionChat.UBound
        optionChat(i).Value = False
    Next
    RTB_Fish.Visible = False
    frameEdit.Visible = False
    RTB_Report.Visible = False
    RTB_User.Visible = False
    RTB_Averages.Visible = False
    RTB_Details.Visible = False
    frameChat.Visible = True
ElseIf comboDisplay.Text = "Summary" Then
    RTB_Fish.Visible = False
    frameEdit.Visible = False
    RTB_Report.Visible = False
    RTB_User.Visible = False
    RTB_Averages.Visible = True
    RTB_Details.Visible = False
    frameChat.Visible = False
ElseIf comboDisplay.Text = "Details" Then
    RTB_Fish.Visible = False
    frameEdit.Visible = False
    RTB_Report.Visible = False
    RTB_User.Visible = False
    RTB_Averages.Visible = False
    RTB_Details.Visible = True
    frameChat.Visible = False
ElseIf comboDisplay.Text = "Edit" Then
    RTB_Fish.Visible = False
    frameEdit.Visible = True
    RTB_Report.Visible = False
    RTB_User.Visible = False
    RTB_Averages.Visible = False
    RTB_Details.Visible = False
    frameChat.Visible = False
ElseIf comboDisplay.Text = "Loot!" Then
    RTB_Fish.Visible = False
    frameEdit.Visible = False
    RTB_Report.Visible = False
    RTB_User.Visible = True
    RTB_Averages.Visible = False
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
    cmdCustom_Click
End If
End Sub


Public Sub comboUser_Click()
On Error Resume Next
Dim PartA As String, PartB As String, PartC As String
RTB_User.Text = ""
RTB_User.TextRTF = ""
RTB_User.SelStart = 0
RTB_User.SelLength = Len(RTB_User.TextRTF)
RTB_User.SelFontName = "Courier New"
RTB_User.SelColor = vbBlack
RTB_User.SelText = UserLog(comboUser.ListIndex, 0) & " - " & Format(Date, "MM/DD/YYYY") & vbNewLine & vbNewLine
RTB_User.SelColor = vbBlack
If Mid$(UserLog(comboUser.ListIndex, 7), 1, 1) = "1" Then
    PartA = "TTL DMG:"
    Do Until Len(PartA) >= 10
        PartA = PartA & " "
    Loop
    RTB_User.SelColor = vbBlack
    RTB_User.SelText = PartA & vbTab
End If
If Mid$(UserLog(comboUser.ListIndex, 7), 3, 1) = "1" Then
    PartA = "Base DMG:"
    Do Until Len(PartA) >= 10
        PartA = PartA & " "
    Loop
    RTB_User.SelColor = vbBlack
    RTB_User.SelText = PartA & vbTab
End If
If Mid$(UserLog(comboUser.ListIndex, 7), 5, 1) = "1" Then
    PartA = "Crit DMG:"
    Do Until Len(PartA) >= 10
        PartA = PartA & " "
    Loop
    RTB_User.SelColor = vbBlack
    RTB_User.SelText = PartA & vbTab
End If
If Mid$(UserLog(comboUser.ListIndex, 7), 7, 1) = "1" Then
    PartA = "Skill DMG:"
    Do Until Len(PartA) >= 10
        PartA = PartA & " "
    Loop
    RTB_User.SelColor = vbBlack
    RTB_User.SelText = PartA & vbTab
End If
If Mid$(UserLog(comboUser.ListIndex, 7), 9, 1) = "1" Then
    PartA = "Spell DMG:"
    Do Until Len(PartA) >= 10
        PartA = PartA & " "
    Loop
    RTB_User.SelColor = vbBlack
    RTB_User.SelText = PartA & vbTab
End If
If Mid$(UserLog(comboUser.ListIndex, 7), 11, 1) = "1" Then
    PartA = "Crit Hits:"
    Do Until Len(PartA) >= 10
        PartA = PartA & " "
    Loop
    RTB_User.SelColor = vbBlack
    RTB_User.SelText = PartA & vbTab
End If
If Mid$(UserLog(comboUser.ListIndex, 7), 13, 1) = "1" Then
    PartA = "Hit/Miss:"
    Do Until Len(PartA) >= 10
        PartA = PartA & " "
    Loop
    RTB_User.SelColor = vbBlack
    RTB_User.SelText = PartA & vbTab
End If
If Mid$(UserLog(comboUser.ListIndex, 7), 15, 1) = "1" Then
    PartA = "Accuracy:"
    Do Until Len(PartA) >= 10
        PartA = PartA & " "
    Loop
    RTB_User.SelColor = vbBlack
    RTB_User.SelText = PartA & vbTab
End If
    RTB_User.SelColor = vbBlack
RTB_User.SelText = "Enemy:                 "

If Mid$(UserLog(comboUser.ListIndex, 7), 17, 1) = "1" Then
    RTB_User.SelText = vbTab & "User Comment:" & vbNewLine
    RTB_User.SelColor = vbBlack
Else
    RTB_User.SelColor = vbBlack
    RTB_User.SelText = vbNewLine
End If

RTB_User.SelColor = vbBlack
RTB_User.SelStart = 0
RTB_User.SelLength = Len(RTB_User.Text)
RTB_User.SelColor = vbBlack
RTB_User.SelBold = True
RTB_User.SelStart = Len(RTB_User.Text)
RTB_User.SelBold = False


RTB_User.SelColor = vbBlack
RTB_User.SelText = UserLog(comboUser.ListIndex, 1)
RTB_User.SelBold = True

'8 Total Base DMG
'9 Total Crit DMG
'10 Total Skill DMG
'11 Total Spell DMG

If Mid$(UserLog(comboUser.ListIndex, 7), 1, 1) = "1" Then
    PartA = UserLog(comboUser.ListIndex, 2)
    Do Until Len(PartA) >= 10
        PartA = PartA & " "
    Loop
    RTB_User.SelColor = vbBlack
    RTB_User.SelText = PartA & vbTab
End If
If Mid$(UserLog(comboUser.ListIndex, 7), 3, 1) = "1" Then
    PartA = UserLog(comboUser.ListIndex, 8)
    Do Until Len(PartA) >= 10
        PartA = PartA & " "
    Loop
    RTB_User.SelColor = vbBlack
    RTB_User.SelText = PartA & vbTab
End If
If Mid$(UserLog(comboUser.ListIndex, 7), 5, 1) = "1" Then
    PartA = UserLog(comboUser.ListIndex, 9)
    Do Until Len(PartA) >= 10
        PartA = PartA & " "
    Loop
    RTB_User.SelColor = vbBlack
    RTB_User.SelText = PartA & vbTab
End If
If Mid$(UserLog(comboUser.ListIndex, 7), 7, 1) = "1" Then
    PartA = UserLog(comboUser.ListIndex, 10)
    Do Until Len(PartA) >= 10
        PartA = PartA & " "
    Loop
    RTB_User.SelColor = vbBlack
    RTB_User.SelText = PartA & vbTab
End If
If Mid$(UserLog(comboUser.ListIndex, 7), 9, 1) = "1" Then
    PartA = UserLog(comboUser.ListIndex, 11)
    Do Until Len(PartA) >= 10
        PartA = PartA & " "
    Loop
    RTB_User.SelColor = vbBlack
    RTB_User.SelText = PartA & vbTab
End If
If Mid$(UserLog(comboUser.ListIndex, 7), 11, 1) = "1" Then
    PartA = UserLog(comboUser.ListIndex, 5)
    Do Until Len(PartA) >= 10
        PartA = PartA & " "
    Loop
    RTB_User.SelColor = vbBlack
    RTB_User.SelText = PartA & vbTab
End If
If Mid$(UserLog(comboUser.ListIndex, 7), 13, 1) = "1" Then
    PartA = UserLog(comboUser.ListIndex, 3) & "/" & UserLog(comboUser.ListIndex, 4)
    Do Until Len(PartA) >= 10
        PartA = PartA & " "
    Loop
    RTB_User.SelColor = vbBlack
    RTB_User.SelText = PartA & vbTab
End If
If Mid$(UserLog(comboUser.ListIndex, 7), 15, 1) = "1" Then
    PartA = Format(Round((CDbl(UserLog(comboUser.ListIndex, 3)) / CDbl(UserLog(comboUser.ListIndex, 3) + CDbl(UserLog(comboUser.ListIndex, 4)))) * 100, 2), "0#.#0") & "%"
    Do Until Len(PartA) >= 10
        PartA = PartA & " "
    Loop
    RTB_User.SelColor = vbBlack
    RTB_User.SelText = PartA & vbTab
End If

'RTB_User.SelText = PartA & vbTab & PartB & vbTab & PartC & vbTab & UserLog(comboUser.ListIndex, 5)
frameEdit.Visible = False
RTB_Report.Visible = False
frameChat.Visible = False
RTB_User.Visible = True
RTB_Fish.Visible = False
RTB_Averages.Visible = False
RTB_Details.Visible = False
optionResults(1).Value = True
End Sub



Private Sub dirList_Change()
fileList.Path = dirList.Path
End Sub



Private Sub Form_Load()
Dim i
On Error Resume Next
ReDim LootFound(0)
ReDim FishFound(0)
ReDim PlayerLoot(0)
ReDim ChatText(0)
Me.Caption = "FFXIP " & App.Major & "." & App.Minor & "." & App.Revision
comboUser.ListIndex = 0
If GetSetting(App.Title, "Settings", "AutoCheck", Default:="") = "" Then
    If MsgBox("OK to always check for updates?", vbYesNo + vbQuestion, "Version Check") = vbYes Then
        SaveSetting App.Title, "Settings", "AutoCheck", "1"
    Else
        SaveSetting App.Title, "Settings", "AutoCheck", "0"
    End If
End If

Dim ColumnText As String
For i = 0 To UBound(ReportOptions)
    If i = 0 Or i = 1 Or i = 2 Or i = 7 Or i = 18 Then
        ReportOptions(i, 0) = GetSetting(App.Title, "Settings", "Report" & i, Default:=1)
    Else
        ReportOptions(i, 0) = GetSetting(App.Title, "Settings", "Report" & i, Default:=0)
    End If
Next
mnuUpdate.Checked = GetSetting(App.Title, "Settings", "AutoCheck", Default:="0")
mnuTray.Checked = GetSetting(App.Title, "Settings", "TrayIcon", Default:="1")
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
        If MyVersion <> Mid$(MyUpdate, 10, MyPosA - 10) Then
            lblUpdate.Visible = True
            lblUpdate.Caption = "Version " & Mid$(MyUpdate, 10, MyPosA - 10) & " now available! - Click here for fixes/changes"
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
Set FSO = Nothing

SingleUser = GetSetting(App.Title, "Settings", "User", Default:="User")
optionResults(1).Caption = SingleUser
SingleDmg = "0"
SingleCrit = "0"
ResetUsers

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
If MsgBox("If your PC explodes while using this program, it's not my fault." & vbNewLine & vbNewLine & "Please read the Notes/Known issues on the website before emailing me!" & vbNewLine & vbNewLine & "********TaruTaru heal Galka - Galka eat TaruTaru********", vbOKCancel + vbInformation, "Notice!") = vbCancel Then
    Unload Me
End If
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
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
frameChat.Width = Me.Width - 300
frameEdit.Width = Me.Width - 300
listResults.Width = frameEdit.Width - 2300
RTB_Report.Width = Me.Width - 300
RTB_User.Width = Me.Width - 300
RTB_Chat.Width = Me.Width - 300
RTB_Fish.Width = Me.Width - 300
RTB_Details.Width = Me.Width - 300
lblUpdate.Width = Me.Width - 300
RTB_Averages.Width = Me.Width - 300
cmdRecalc.Left = listResults.Width + 150
cmdUnselect.Left = listResults.Width + 150
cmdSelect.Left = listResults.Width + 150
cmdCustom.Left = listResults.Width + 150
cmdExport.Left = listResults.Width + 150
comboMOB.Left = listResults.Width + 150
lblInfo(0).Left = listResults.Width + 250
lblInfo(1).Left = listResults.Width + 250
lblInfo(0).Top = (frameEdit.Height - lblInfo(0).Height) - 200

If lblUpdate.Visible = True Then
    frameEdit.Height = Me.Height - 1565
    frameChat.Height = Me.Height - 1565
    listResults.Height = frameEdit.Height - 300
    RTB_Chat.Height = frameChat.Height - 320
    RTB_Report.Height = Me.Height - 1565
    RTB_User.Height = Me.Height - 1565
    RTB_Fish.Height = Me.Height - 1565
    RTB_Averages.Height = Me.Height - 1565
    RTB_Details.Height = Me.Height - 1565
Else
    frameEdit.Height = Me.Height - 1305
    frameChat.Height = Me.Height - 1305
    listResults.Height = frameEdit.Height - 300
    RTB_Chat.Height = frameChat.Height - 320
    RTB_Report.Height = Me.Height - 1305
    RTB_User.Height = Me.Height - 1305
    RTB_Fish.Height = Me.Height - 1305
    RTB_Averages.Height = Me.Height - 1305
    RTB_Details.Height = Me.Height - 1305
End If
Shape1.Width = Me.Width - 1330
lblStatus.Width = Me.Width - 2050
cmdOptions.Left = Me.Width - 1100

If lblUpdate.Visible = True Then
    optionResults(0).Top = RTB_Report.Top + RTB_Report.Height + 355
    optionResults(1).Top = RTB_Report.Top + RTB_Report.Height + 355
    comboDisplay.Top = RTB_Report.Top + RTB_Report.Height + 290
    comboUser.Top = RTB_Report.Top + RTB_Report.Height + 290
    lblChange.Top = RTB_Report.Top + RTB_Report.Height + 355
    cmdOptions.Top = RTB_Report.Top + RTB_Report.Height + 290
Else
    optionResults(0).Top = RTB_Report.Top + RTB_Report.Height + 90
    optionResults(1).Top = RTB_Report.Top + RTB_Report.Height + 90
    comboDisplay.Top = RTB_Report.Top + RTB_Report.Height + 40
    comboUser.Top = RTB_Report.Top + RTB_Report.Height + 40
    lblChange.Top = RTB_Report.Top + RTB_Report.Height + 90
    cmdOptions.Top = RTB_Report.Top + RTB_Report.Height + 40
End If
lblUpdate.Top = RTB_Report.Top + RTB_Report.Height + 30
lblAbout.Left = cmdOptions.Left - 550
lblAbout.Top = lblChange.Top
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
SaveSetting App.Title, "Settings", "User", SingleUser
SaveSetting App.Title, "Settings", "UserB", SingleUserB
End Sub



Private Sub lblAbout_Click()
frmAbout.Show
End Sub

Private Sub lblChange_Click()
frmUsers.Show
End Sub

Private Sub lblStatus_Change()
If InStr(1, lblStatus, "True") Then
    lblStatus.ForeColor = vbRed
Else
    lblStatus.ForeColor = vbBlack
End If
End Sub

Private Sub lblUpdate_Click()
If InStr(1, lblUpdate.Caption, "Click here") Then
    If MsgBox(Updates & vbNewLine & vbNewLine & "Visit website?" & vbNewLine & "(www.frontiernet.net/~Spyle/FFXI/ffxi.html)", vbYesNo, "Update Info") = vbYes Then
        ShellExecute Me.hwnd, vbNullString, "http://www.frontiernet.net/~Spyle/FFXI/ffxi.html", vbNullString, "C:\", SW_SHOWNORMAL
    End If
End If
End Sub

Public Sub mnuClear_Click()
If NotClearFile = False Then
    Dim FSO As FileSystemObject
    Set FSO = New FileSystemObject
    
    If FSO.FileExists(App.Path & "\EditFile.log") = True Then
        FSO.DeleteFile (App.Path & "\EditFile.log")
    End If
    Set FSO = Nothing
End If

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
FishHeader = ""
FishComment = ""

FightStartTime = "12:00:00 AM"
FightStopTime = "12:00:00 AM"
StartTime = "12:00:00 AM"
StopTime = "12:00:00 AM"
StartTimeDPS = "12:00:00 AM"
StopTimeDPS = "12:00:00 AM"
If ClearEdit Then
    listResults.Clear
End If
Erase PList
Erase GrandPList
Erase Stats
Erase Players
Erase UserLog
Erase DPS
Erase ChatText
ReDim ChatText(0)
Erase LootFound
ReDim LootFound(0)
Erase FishFound
ReDim FishFound(0)
Erase PlayerLoot
ReDim PlayerLoot(0)
HTMLCode = ""
SummaryCode = ""
UniqueMOB = 0
P1Special = ""
P1 = ""
P1Opp = ""
P1Stat = ""
P1Takes = ""
P1Uses = ""
HasErrors = False
Casts = False
Critical = False
FightComment = ""
MonsterCheck = ""
TotalExp = 0
ErrorCount = 0
dHigh = 0
dLow = 10000
TotalExp = 0

RTB_Report.Text = ""
RTB_Fish.Text = ""
RTB_Chat.Text = ""
RTB_User.Text = ""
RTB_Averages.Text = ""
RTB_Details.Text = ""
ResetUsers
End Sub

Private Sub mnuExit_Click()
If Me.WindowState <> 0 Then Me.WindowState = 0
Me.Visible = True
Me.SetFocus
Unload Me
End Sub

Private Sub mnuMonster_Click()
If mnuMonster.Checked = False Then
    mnuMonster.Checked = True
    mnuPlayer.Checked = False
Else
    mnuMonster.Checked = False
End If
End Sub

Private Sub mnuPlayer_Click()
If mnuPlayer.Checked = False Then
    mnuPlayer.Checked = True
    mnuMonster.Checked = False
Else
    mnuPlayer.Checked = False
End If
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

Private Sub mnuUser_Click()
lblChange_Click
End Sub

Private Sub optionChat_Click(Index As Integer)
Dim i, LineType As String, OldLabel As String

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
            RTB_Chat.SelText = Replace(Replace(Mid$(ChatText(i), 3, Len(ChatText(i)) - 4), "ï'", "["), "ï(", "]")
            RTB_Chat.SelText = vbNewLine
            RTB_Chat.SelStart = Len(RTB_Chat.Text)
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

Private Sub RTB_Averages_DblClick()
CD_Save.Flags = cdlOFNOverwritePrompt
CD_Save.Filter = "Rich Text File (*.rtf)|*.rtf|Plain Text File (*.txt)|*.txt"

CD_Save.ShowSave
If CD_Save.FileName <> "0" And CD_Save.FileName <> "" Then
    If Right$(CD_Save.FileName, 3) = "rtf" Then
        RTB_Averages.SaveFile CD_Save.FileName
    Else
        RTB_Averages.SaveFile CD_Save.FileName, rtfText
    End If
End If
End Sub


Private Sub RTB_Chat_DblClick()
CD_Save.Flags = cdlOFNOverwritePrompt
CD_Save.Filter = "Rich Text File (*.rtf)|*.rtf|Plain Text File (*.txt)|*.txt"

CD_Save.ShowSave
If CD_Save.FileName <> "0" And CD_Save.FileName <> "" Then
    If Right$(CD_Save.FileName, 3) = "rtf" Then
        RTB_Chat.SaveFile CD_Save.FileName
    Else
        RTB_Chat.SaveFile CD_Save.FileName, rtfText
    End If
End If
End Sub


Private Sub RTB_Details_DblClick()
CD_Save.Flags = cdlOFNOverwritePrompt
CD_Save.Filter = "Rich Text File (*.rtf)|*.rtf|Plain Text File (*.txt)|*.txt"

CD_Save.ShowSave
If CD_Save.FileName <> "0" And CD_Save.FileName <> "" Then
    If Right$(CD_Save.FileName, 3) = "rtf" Then
        RTB_Details.SaveFile CD_Save.FileName
    Else
        RTB_Details.SaveFile CD_Save.FileName, rtfText
    End If
End If
End Sub








Private Sub RTB_Fish_DblClick()
CD_Save.Flags = cdlOFNOverwritePrompt
CD_Save.Filter = "Rich Text File (*.rtf)|*.rtf|Plain Text File (*.txt)|*.txt"

CD_Save.ShowSave
If CD_Save.FileName <> "0" And CD_Save.FileName <> "" Then
    If Right$(CD_Save.FileName, 3) = "rtf" Then
        RTB_Fish.SaveFile CD_Save.FileName
    Else
        RTB_Fish.SaveFile CD_Save.FileName, rtfText
    End If
End If
End Sub


Private Sub RTB_Report_DblClick()
CD_Save.Flags = cdlOFNOverwritePrompt
CD_Save.Filter = "Rich Text File (*.rtf)|*.rtf|Plain Text File (*.txt)|*.txt"

CD_Save.ShowSave
If CD_Save.FileName <> "0" And CD_Save.FileName <> "" Then
    If Right$(CD_Save.FileName, 3) = "rtf" Then
        RTB_Report.SaveFile CD_Save.FileName
    Else
        RTB_Report.SaveFile CD_Save.FileName, rtfText
    End If
End If
End Sub




Private Sub RTB_User_DblClick()
CD_Save.Flags = cdlOFNOverwritePrompt
CD_Save.Filter = "Rich Text File (*.rtf)|*.rtf|Plain Text File (*.txt)|*.txt"

CD_Save.ShowSave
If CD_Save.FileName <> "0" And CD_Save.FileName <> "" Then
    If Right$(CD_Save.FileName, 3) = "rtf" Then
        RTB_User.SaveFile CD_Save.FileName
    Else
        RTB_User.SaveFile CD_Save.FileName, rtfText
    End If
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


Private Sub timerRead_Timer()
Dim z As Integer, o As Integer
Dim MyDate As Date
Dim FullFile() As String, CurrentLine As String, LineType As String, PrevLine As String, MyPosAdd As Integer, MyPos As Long, CurrentFile As String
Dim Index As Long

If fileList.ListCount <> 0 Then

    If lblStatus.Caption = "Too many errors - Parsing stopped for this log." Then
    Else
        lblStatus.Caption = "Errors: " & HasErrors & " - " & "Waiting for new log...."
    End If
    DoEvents
    
    fileListBox.Clear

    fileList.Refresh
    For i = 0 To fileList.ListCount - 1
        fileList.ListIndex = i
        fileListBox.AddItem Format(FileDateTime(dirList.Path & "\" & fileList.FileName), "MM/DD HhNnSs") & " - " & fileList.Path & "\" & fileList.FileName
    Next
    
    fileListBox.ListIndex = fileListBox.ListCount - 1
    If LastItem <> fileListBox.Text Then
    
        If Gather = True Then
            Dim EditFile
            EditFile = FreeFile
            Open SingleFile For Append As #EditFile
        End If
    
        RTB_Report.SelStart = Len(RTB_Report.Text)
        lblStatus.Caption = "Errors: " & HasErrors & " - " & "Parsing Data...."
        DoEvents
        f = FreeFile

        RTB.LoadFile Mid$(fileListBox.Text, 16)
        RTB.Text = Mid(RTB.Text, 101)
        RTB.Text = Replace(RTB.Text, Chr(0), vbNewLine)
        MyPos = InStrRev(fileListBox.Text, "\")
        If Gather = False Then
          CurrentFile = App.Path & "\FFXI_Logs" & Mid(fileListBox.Text, MyPos)
        Else
          CurrentFile = App.Path & "\FFXI_Gather" & Mid(fileListBox.Text, MyPos)
        End If
        RTB.SaveFile CurrentFile, rtfText

        MyDate = Left$(fileListBox.Text, 5) & Format(Date, "/YYYY") & " " & Format(Format(Mid$(fileListBox.Text, 7, 6), "00:00:00"), "Hh:Nn:Ss AM/PM")
        ResetTimeFile CurrentFile, MyDate
        
        Erase FullFile
        Index = 0
        Open CurrentFile For Input As f
          Do Until EOF(f)
            Line Input #f, CurrentLine
            LineType = Left(CurrentLine, 2)
            If Mid(CurrentLine, 51, 2) = "01" And Index <> 0 Then
                FullFile(Index - 1) = Left(FullFile(Index - 1), Len(FullFile(Index - 1)) - 3) & Mid(CurrentLine, 56) & " " & LineType
            Else
                ReDim Preserve FullFile(Index)
                FullFile(Index) = Mid(CurrentLine, 54) & " " & LineType
                Index = Index + 1
            End If
          Loop
        Close #f
        If Gather = False Then
          If Index <> 0 Then
            ParseLog FullFile, False, False
          End If
        Else
            If Index <> 0 Then
                For i = 0 To UBound(FullFile)
                    Print #EditFile, FullFile(i)
                Next
            End If
        End If

    lblStatus.Caption = "Errors: " & HasErrors & " - " & "Waiting for new log...."
    fileListBox.ListIndex = fileListBox.ListCount - 1
    LastItem = fileListBox.Text
    If optionResults(1).Value = True Then
        comboUser_Click
    End If
    If Gather = True Then
        Close #EditFile
    End If
End If
Else
    MsgBox "No log files found in this folder. Please select another folder.", vbInformation, "Error"
    cmdOpen.Caption = "Start"
    lblStatus.Caption = "Errors: " & HasErrors & " - " & "Waiting."
    timerRead.Enabled = False
    frmOpen.Visible = False
    frmOpen.Left = Me.Left + 200
    frmOpen.Top = Me.Top + 200
    frmOpen.Visible = True
End If
Exit Sub
End Sub




