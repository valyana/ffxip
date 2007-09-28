VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About"
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3405
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAbout.frx":1E72
   ScaleHeight     =   1785
   ScaleWidth      =   3405
   Begin VB.Label lblAbout 
      BackStyle       =   0  'Transparent
      Caption         =   "www.frontiernet.net/~Spyle/FFXI/ffxi.html"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   90
      MouseIcon       =   "frmAbout.frx":4808
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   1485
      Width           =   3210
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Old webpage: "
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   90
      TabIndex        =   2
      Top             =   1305
      Width           =   1320
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Original by Spyle, modified by Valyana"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   90
      TabIndex        =   1
      Top             =   1080
      Width           =   3225
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "Version 5.0.0"
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   90
      TabIndex        =   0
      Top             =   855
      Width           =   1005
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, _
                                                                               ByVal lpOperation As String, _
                                                                               ByVal lpFile As String, _
                                                                               ByVal lpParameters As String, _
                                                                               ByVal lpDirectory As String, _
                                                                               ByVal nShowCmd As Long) _
                                                                               As Long



Private Sub Form_Click()
Unload Me
End Sub


Private Sub Form_Load()
Me.Left = frmRead.Left + ((frmRead.Width / 2) - frmAbout.Width / 2)
Me.Top = frmRead.Top + ((frmRead.Height / 2) - frmAbout.Height / 2)
lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub Label1_Click()
Unload Me
End Sub

Private Sub Label2_Click()
Unload Me
End Sub


Private Sub lblAbout_Click()
'ShellExecute Me.hwnd, vbNullString, "http://www.frontiernet.net/~Spyle/FFXI/ffxi.html", vbNullString, "C:\", SW_SHOWNORMAL
End Sub


Private Sub lblVersion_Click()
Unload Me
End Sub


