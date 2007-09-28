VERSION 5.00
Begin VB.Form frmGraph 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8910
   FillColor       =   &H00FFFFFF&
   Icon            =   "frmGraph.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   8910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lblKey 
      BackStyle       =   0  'Transparent
      Caption         =   "None"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Index           =   5
      Left            =   7605
      TabIndex        =   11
      Top             =   630
      Width           =   1230
   End
   Begin VB.Label lblKey 
      BackStyle       =   0  'Transparent
      Caption         =   "None"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Index           =   4
      Left            =   7605
      TabIndex        =   10
      Top             =   450
      Width           =   1230
   End
   Begin VB.Label lblKey 
      BackStyle       =   0  'Transparent
      Caption         =   "None"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Index           =   3
      Left            =   7605
      TabIndex        =   9
      Top             =   270
      Width           =   1230
   End
   Begin VB.Label lblKey 
      BackStyle       =   0  'Transparent
      Caption         =   "None"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Index           =   2
      Left            =   6210
      TabIndex        =   8
      Top             =   630
      Width           =   1230
   End
   Begin VB.Label lblKey 
      BackStyle       =   0  'Transparent
      Caption         =   "None"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Index           =   1
      Left            =   6210
      TabIndex        =   7
      Top             =   450
      Width           =   1230
   End
   Begin VB.Label lblKey 
      BackStyle       =   0  'Transparent
      Caption         =   "None"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Index           =   0
      Left            =   6210
      TabIndex        =   6
      Top             =   270
      Width           =   1230
   End
   Begin VB.Shape shapeKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   105
      Index           =   5
      Left            =   7470
      Top             =   675
      Width           =   105
   End
   Begin VB.Shape shapeKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   105
      Index           =   4
      Left            =   7470
      Top             =   495
      Width           =   105
   End
   Begin VB.Shape shapeKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   105
      Index           =   3
      Left            =   7470
      Top             =   315
      Width           =   105
   End
   Begin VB.Shape shapeKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   105
      Index           =   2
      Left            =   6075
      Top             =   675
      Width           =   105
   End
   Begin VB.Shape shapeKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   105
      Index           =   1
      Left            =   6075
      Top             =   495
      Width           =   105
   End
   Begin VB.Shape shapeKey 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   105
      Index           =   0
      Left            =   6075
      Top             =   315
      Width           =   105
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      Height          =   555
      Left            =   6030
      Top             =   270
      Width           =   2850
   End
   Begin VB.Shape Shape1 
      Height          =   4380
      Left            =   315
      Top             =   225
      Width           =   8610
   End
   Begin VB.Label lblGraph 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   195
      Index           =   4
      Left            =   0
      TabIndex        =   5
      Top             =   4410
      Width           =   690
   End
   Begin VB.Label lblGraph 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   195
      Index           =   3
      Left            =   0
      TabIndex        =   4
      Top             =   540
      Width           =   330
   End
   Begin VB.Label lblGraph 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   195
      Index           =   2
      Left            =   0
      TabIndex        =   3
      Top             =   360
      Width           =   330
   End
   Begin VB.Label lblGraph 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   195
      Index           =   1
      Left            =   0
      TabIndex        =   2
      Top             =   180
      Width           =   330
   End
   Begin VB.Label lblGraph 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   195
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   330
   End
   Begin VB.Label lblAmt 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   195
      Left            =   3870
      TabIndex        =   0
      Top             =   0
      Width           =   915
   End
End
Attribute VB_Name = "frmGraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MultiplyBy As Long

Public Sub OpenNew(UniqueMOB As Long)
On Error GoTo err_handler
Me.Show
Dim Results(5, 1) As String
Dim i, WhichColor As ColorConstants, PerSeg As Long, ThisSeg As Long, TotalHits As Long, LastX As Long, LastY As Long, X1 As Long, X2 As Long, Y1 As Long, Y2 As Long
Dim ThisLine As String, CurrentNumber As Long, WhichNumber As Long, FirstLine As Boolean, CurrentHitLine As String
Dim HighHit As Long, TotalPlayers As Integer
Dim CurrentPlayer As String

For i = 0 To frmRead.comboPlayer.ListCount - 1
    MyPos = InStr(1, frmRead.comboPlayer.List(i), "-")
    If UniqueMOB = Left(frmRead.comboPlayer.List(i), MyPos - 1) And InStr(1, LCase(frmRead.comboPlayer.List(i)), "skillchain") = 0 Then
        MyPos2 = InStr(MyPos + 1, frmRead.comboPlayer.List(i), "-")
        CurrentPlayer = Mid(frmRead.comboPlayer.List(i), MyPos + 1, MyPos2 - (MyPos + 1))
        CurrentHitLine = Mid(frmRead.comboPlayer.List(i), MyPos2 + 1)
        Results(TotalPlayers, 0) = CurrentPlayer
        Results(TotalPlayers, 1) = CurrentHitLine
        TotalPlayers = TotalPlayers + 1
        If TotalPlayers = 6 Then Exit For
    End If
Next

For i = 0 To UBound(Results)
    ThisLine = Results(i, 1)
    Do Until ThisLine = ""
        MyPos = InStr(1, ThisLine, ",")
        If MyPos <> 0 Then
            CurrentNumber = Left(ThisLine, MyPos - 1)
            ThisLine = Mid(ThisLine, MyPos + 1)
        Else
            CurrentNumber = ThisLine
            ThisLine = ""
        End If
        If CurrentNumber > HighHit Then
            HighHit = CurrentNumber
        End If
    Loop
Next

If HighHit <= 50 Then
    MultiplyBy = 75
    lblGraph(0).Caption = 50
    lblGraph(0).Top = (Me.ScaleHeight - (50 * MultiplyBy))
ElseIf HighHit <= 100 Then
    MultiplyBy = 40
    lblGraph(0).Caption = 100
    lblGraph(0).Top = (Me.ScaleHeight - (100 * MultiplyBy))
ElseIf HighHit <= 150 Then
    MultiplyBy = 28
    lblGraph(0).Caption = 150
    lblGraph(0).Top = (Me.ScaleHeight - (150 * MultiplyBy))
ElseIf HighHit <= 200 Then
    MultiplyBy = 20
    lblGraph(0).Caption = 200
    lblGraph(0).Top = (Me.ScaleHeight - (200 * MultiplyBy))
ElseIf HighHit <= 250 Then
    MultiplyBy = 16
    lblGraph(0).Caption = 250
    lblGraph(0).Top = (Me.ScaleHeight - (250 * MultiplyBy))
ElseIf HighHit <= 300 Then
    MultiplyBy = 14
    lblGraph(0).Caption = 300
    lblGraph(0).Top = (Me.ScaleHeight - (300 * MultiplyBy))
ElseIf HighHit <= 350 Then
    MultiplyBy = 12
    lblGraph(0).Caption = 350
    lblGraph(0).Top = (Me.ScaleHeight - (350 * MultiplyBy))
ElseIf HighHit <= 400 Then
    MultiplyBy = 10
    lblGraph(0).Caption = 400
    lblGraph(0).Top = (Me.ScaleHeight - (400 * MultiplyBy))
End If
X1 = 315
X2 = Me.ScaleWidth
Y1 = lblGraph(0).Top
Y2 = lblGraph(0).Top
Me.Line (X1, Y1)-(X2, Y2)

lblGraph(1).Caption = (lblGraph(0) / 4) * 3
lblGraph(1).Top = Me.ScaleHeight - (((Me.ScaleHeight - lblGraph(0).Top) / 4) * 3)
Y1 = lblGraph(1).Top
Y2 = lblGraph(1).Top
Line (X1, Y1)-(X2, Y2)

lblGraph(2).Caption = (lblGraph(0) / 2)
lblGraph(2).Top = Me.ScaleHeight - (((Me.ScaleHeight - lblGraph(0).Top) / 4) * 2)
Y1 = lblGraph(2).Top
Y2 = lblGraph(2).Top
Line (X1, Y1)-(X2, Y2)

lblGraph(3).Caption = (lblGraph(0) / 4)
lblGraph(3).Top = Me.ScaleHeight - (((Me.ScaleHeight - lblGraph(0).Top) / 4) * 1)
Y1 = lblGraph(3).Top
Y2 = lblGraph(3).Top
Line (X1, Y1)-(X2, Y2)

lblGraph(4).Caption = 0
lblGraph(3).Top = Me.ScaleHeight - (((Me.ScaleHeight - lblGraph(0).Top) / 4))

Me.DrawWidth = 3
For i = 0 To UBound(Results)
    If Results(i, 0) <> "" Then
        TotalHits = 0
        For X = 1 To Len(Results(i, 1))
            If Mid(Results(i, 1), X, 1) = "," Then
                TotalHits = TotalHits + 1
            End If
        Next
        PerSeg = Me.ScaleWidth / (TotalHits + 1)
        
        lblKey(i).Caption = Results(i, 0)
        If i = 0 Then
            WhichColor = vbRed
            shapeKey(i).BackColor = vbRed
        ElseIf i = 1 Then
            WhichColor = vbBlue
            shapeKey(i).BackColor = vbBlue
        ElseIf i = 2 Then
            WhichColor = vbGreen
            shapeKey(i).BackColor = vbGreen
        ElseIf i = 3 Then
            WhichColor = vbBlack
            shapeKey(i).BackColor = vbBlack
        ElseIf i = 4 Then
            WhichColor = vbCyan
            shapeKey(i).BackColor = vbCyan
        ElseIf i = 5 Then
            WhichColor = vbMagenta
            shapeKey(i).BackColor = vbMagenta
        End If
        LastY = 0
        ThisLine = Results(i, 1)
        WhichNumber = 0
        Do Until ThisLine = ""
            MyPos = InStr(1, ThisLine, ",")
            If MyPos <> 0 Then
                CurrentNumber = Left(ThisLine, MyPos - 1)
                ThisLine = Mid(ThisLine, MyPos + 1)
            Else
                CurrentNumber = ThisLine
                ThisLine = ""
            End If
            If LastY = 0 Then
                LastX = 0
                LastY = Me.ScaleHeight - (CurrentNumber * MultiplyBy)
                ThisSeg = PerSeg * (WhichNumber)
                X2 = ThisSeg + (PerSeg / 2)
                Y2 = Me.ScaleHeight - (CurrentNumber * MultiplyBy)
                LastX = ThisSeg + (PerSeg / 2)
                LastY = Me.ScaleHeight - (CurrentNumber * MultiplyBy)
            Else
                ThisSeg = PerSeg * (WhichNumber)
                X2 = ThisSeg + (PerSeg / 2)
                Y2 = Me.ScaleHeight - (CurrentNumber * MultiplyBy)
                Line (LastX, LastY)-(X2, Y2), WhichColor
                LastX = ThisSeg + (PerSeg / 2)
                LastY = Me.ScaleHeight - (CurrentNumber * MultiplyBy)
            End If
            WhichNumber = WhichNumber + 1
        Loop
    End If
Next
Exit Sub
err_handler:
Stop
Resume
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If MultiplyBy <> 0 Then
    lblAmt = Round((Me.ScaleHeight - (Y)) / MultiplyBy)
End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
Set frmGraph = Nothing
End Sub


