VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmOpen 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Start FFXIP"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6540
   Icon            =   "frmOpen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   6540
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog CD1 
      Left            =   3015
      Top             =   2655
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Logs (*.txt;*.log)|*.txt;*.log"
   End
   Begin VB.Frame Frame2 
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   4140
      TabIndex        =   10
      Top             =   45
      Width           =   2355
      Begin VB.OptionButton optionAction 
         Caption         =   "Gather log files to date."
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   5
         ToolTipText     =   $"frmOpen.frx":1E72
         Top             =   810
         Width           =   2040
      End
      Begin VB.TextBox txtGather 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   90
         MaxLength       =   20
         TabIndex        =   7
         ToolTipText     =   "Gathered Logs will be saved to the FFXIP directory."
         Top             =   2115
         Width           =   2130
      End
      Begin VB.OptionButton optionAction 
         Caption         =   "Open saved log file."
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   6
         ToolTipText     =   "Use to open a log file that you have gathered."
         Top             =   1080
         Width           =   1860
      End
      Begin VB.OptionButton optionAction 
         Caption         =   "Parse immediately."
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   3
         ToolTipText     =   "Will immediately parse existing FFXI logs and any new ones that appear."
         Top             =   270
         Value           =   -1  'True
         Width           =   1860
      End
      Begin VB.OptionButton optionAction 
         Caption         =   "Gather log files."
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   4
         ToolTipText     =   $"frmOpen.frx":1EFE
         Top             =   540
         Width           =   2130
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Hold mouse over options for descriptions."
         ForeColor       =   &H00FF0000&
         Height          =   420
         Left            =   270
         TabIndex        =   12
         Top             =   1395
         Width           =   1680
      End
      Begin VB.Label Label1 
         Caption         =   "Save gathered logs as:"
         Height          =   195
         Left            =   90
         TabIndex        =   11
         Top             =   1890
         Width           =   2085
      End
   End
   Begin VB.Frame frameFolder 
      Caption         =   "Select FFXI Log Folder"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   45
      TabIndex        =   9
      Top             =   45
      Width           =   4065
      Begin VB.DirListBox dirList 
         Height          =   1890
         Left            =   90
         TabIndex        =   2
         Top             =   540
         Width           =   3885
      End
      Begin VB.DriveListBox drvList 
         Height          =   315
         Left            =   90
         TabIndex        =   1
         Top             =   225
         Width           =   3930
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   330
      Left            =   5310
      TabIndex        =   8
      Top             =   2610
      Width           =   1185
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   330
      Left            =   45
      TabIndex        =   0
      Top             =   2610
      Width           =   1185
   End
End
Attribute VB_Name = "frmOpen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
Dim FSO As FileSystemObject
Set FSO = New FileSystemObject
If FSO.FileExists(App.Path & "\error_log.txt") = True Then
    FSO.DeleteFile App.Path & "\error_log.txt"
End If
Dim ErrorFile
Dim SaveFile
Set ErrorFile = FSO.CreateTextFile(App.Path & "\error_log.txt", True)
ErrorFile.WriteLine ("FFXI Parser Error Log")
ErrorFile.Close
Set ErrorFile = Nothing

frmRead.dirList.Path = dirList.Path
SaveSetting App.Title, "Settings", "LogPath", dirList.Path
frmRead.listResults.Clear
frmRead.mnuClear_Click

If optionAction(0).Value = True Then
    OpenSingle = False
    Gather = False
    frmRead.StartNew
    Unload Me
ElseIf optionAction(1).Value = True Then
    If Trim(txtGather.Text) <> "" Then
        OpenSingle = False
        Gather = True
        If FSO.FileExists(App.Path & "\" & txtGather.Text) = True Then
            FSO.DeleteFile App.Path & "\" & txtGather.Text
        End If
        Set SaveFile = FSO.CreateTextFile(App.Path & "\" & txtGather.Text, True)
        SaveFile.Close
        Set SaveFile = Nothing
        SingleFile = App.Path & "\" & txtGather.Text
        frmRead.StartNew
        Unload Me
    Else
        MsgBox "Please enter a file name.", vbInformation, "Gather"
    End If
ElseIf optionAction(3).Value = True Then
    OpenSingle = False
    GatherDate = True
    Gather = True
    If FSO.FileExists(App.Path & "\" & Format(Date, "MM-DD-YYYY") & ".log") = True Then
        FSO.DeleteFile App.Path & "\" & Format(Date, "MM-DD-YYYY") & ".log"
    End If
    Set SaveFile = FSO.CreateTextFile(App.Path & "\" & Format(Date, "MM-DD-YYYY") & ".log", True)
    SaveFile.Close
    Set SaveFile = Nothing
    SingleFile = App.Path & "\" & Format(Date, "MM-DD-YYYY") & ".log"
    frmRead.StartNew
    Unload Me
Else
    OpenSingle = True
    Gather = False
    CD1.InitDir = App.Path
    CD1.ShowOpen
    If CD1.FileName <> "" Then
        Screen.MousePointer = vbHourglass
        DoEvents
        SingleFile = CD1.FileName
        If InStr(1, CD1.FileName, "EditFile.log") = 0 Then
            frmRead.StartNew
            Unload Me
        Else
            MsgBox "Unable to open '" & SingleFile & "'", vbCritical, "Error"
        End If
        Screen.MousePointer = vbDefault
    End If
End If
Set FSO = Nothing
End Sub


Private Sub drvList_Change()
On Error Resume Next
dirList.Path = drvList.Drive
End Sub


Private Sub Form_Activate()
txtGather = Format(Date, "MM-DD-YYYY") & ".log"
dirList.Path = frmRead.dirList.Path
If GetSetting(App.Title, "Settings", "NewUserA", Default:=True) Then
    MsgBox "Select the folder where the FFXI log files are located." & vbNewLine & vbNewLine & "Usually: ""C:\Program Files\PlayOnline\SquareEnix\FINAL FANTASY XI\TEMP""", vbInformation, "Folder Select"
    SaveSetting App.Title, "Settings", "NewUserA", False
End If
End Sub

Private Sub txtGather_Validate(Cancel As Boolean)
If txtGather.Text = "" Then txtGather.Text = "GatheredLogs.log"
If InStr(1, txtGather, ".") <> 0 Then
    If Right$(txtGather, 4) <> ".txt" And Right$(txtGather, 4) <> ".log" Then
        MsgBox "Invalid Format." & vbNewLine & vbNewLine & "Example:" & vbNewLine & "GatheredLogs.log", vbInformation, "Gather"
        Cancel = True
    End If
Else
    txtGather.Text = Trim(txtGather.Text) & ".log"
End If
End Sub


