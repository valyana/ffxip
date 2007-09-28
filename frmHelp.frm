VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmHelp 
   Caption         =   "Parser Commands"
   ClientHeight    =   3720
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5310
   Icon            =   "frmHelp.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3720
   ScaleWidth      =   5310
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox comboList 
      Height          =   315
      ItemData        =   "frmHelp.frx":030A
      Left            =   990
      List            =   "frmHelp.frx":030C
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   45
      Width           =   4335
   End
   Begin RichTextLib.RichTextBox RTB_Help 
      Height          =   3300
      Left            =   0
      TabIndex        =   1
      Top             =   405
      Width           =   5325
      _ExtentX        =   9393
      _ExtentY        =   5821
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmHelp.frx":030E
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
   Begin VB.Label lbl 
      Caption         =   "Command:"
      Height          =   240
      Left            =   45
      TabIndex        =   2
      Top             =   90
      Width           =   1410
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub comboList_Click()
With RTB_Help
    .Text = ""
    If comboList.ListIndex = 0 Then
        .SelBold = True
        .SelText = "Basic usage for all commands" & vbNewLine
        .SelColor = vbBlue
        .SelText = """" & "/echo parser " & """"
        .SelColor = vbBlack
        .SelText = " and the command." & vbNewLine & vbNewLine
        
        .SelBold = True
        .SelText = "Command List: " & vbNewLine
        .SelText = """" & "start fight" & """" & " *" & vbNewLine
        .SelText = """" & "stop fight" & """" & " *" & vbNewLine
        .SelText = """" & "start dps" & """" & " *" & vbNewLine
        .SelText = """" & "stop dps" & """" & " *" & vbNewLine
        .SelText = """" & "clear dps" & """" & vbNewLine
        .SelText = """" & "start exp" & """" & " *" & vbNewLine
        .SelText = """" & "stop exp" & """" & " *" & vbNewLine
        .SelText = """" & "start fish" & """" & " *" & vbNewLine
        .SelText = """" & "stop fish" & """" & " *" & vbNewLine
        .SelText = """" & "clear" & """" & vbNewLine
        .SelText = """" & "save report" & """" & vbNewLine
        .SelText = """" & "save summary" & """" & vbNewLine
        .SelText = """" & "save details" & """" & vbNewLine
        .SelText = """" & "save player1-6" & """" & vbNewLine
        .SelText = """" & "beep" & """" & vbNewLine
        .SelText = """" & "window1-7" & """" & vbNewLine
        .SelText = """" & "start logging" & """" & vbNewLine
        .SelText = """" & "stop logging" & """" & vbNewLine
        .SelText = """" & "comment" & """" & vbNewLine
        .SelText = """" & "timer ##:##:## '1-10'" & """" & vbNewLine
        .SelText = """" & "fish comment" & """" & vbNewLine & vbNewLine
        .SelBold = True
        .SelText = "* must be followed by /clock" & vbNewLine
        .SelText = "The /clock command can follow anytime after the command that requires it." & vbNewLine & vbNewLine & vbNewLine
    End If
    
    If comboList.ListIndex = 1 Then
        .SelBold = True
        .SelColor = vbRed
        .SelText = "Fight Efficiency - Displayed per fight on Report." & vbNewLine
        .SelBold = True
        .SelText = "Commands used:" & vbNewLine
        .SelColor = vbBlue
        .SelText = """" & "start fight" & """" & vbNewLine
        .SelColor = vbBlue
        .SelText = """" & "stop fight" & """" & vbNewLine
        .SelText = """" & "/clock" & """" & vbNewLine & vbNewLine
        
        .SelBold = True
        .SelText = "Example:" & vbNewLine
        .SelText = "1. You begin fighting." & vbNewLine
        .SelText = "2. /echo parser start fight" & vbNewLine
        .SelText = "3. /clock" & vbNewLine
        .SelText = "4. You kill the monster." & vbNewLine
        .SelText = "5. /echo parser stop fight" & vbNewLine
        .SelText = "6. /clock" & vbNewLine & vbNewLine
        
        .SelBold = True
        .SelText = "Notes:" & vbNewLine
        .SelText = "This will also calculate the DPS on the summary so you don't have to use both commands." & vbNewLine & vbNewLine
    End If
        
    If comboList.ListIndex = 2 Then
        .SelBold = True
        .SelColor = vbRed
        .SelText = "Damage Per Second Calculation - Displayed on summary." & vbNewLine
        .SelBold = True
        .SelText = "Commands used:" & vbNewLine
        .SelColor = vbBlue
        .SelText = """" & "start dps" & """" & vbNewLine
        .SelColor = vbBlue
        .SelText = """" & "stop dps" & """" & vbNewLine
        .SelText = """" & "/clock" & """" & vbNewLine
        .SelColor = vbBlue
        .SelText = """" & "clear dps" & """" & vbNewLine & vbNewLine
        
        .SelBold = True
        .SelText = "Example:" & vbNewLine
        .SelText = "1. You begin fighting." & vbNewLine
        .SelText = "2. /echo parser start dps" & vbNewLine
        .SelText = "3. /clock" & vbNewLine
        .SelText = "4. You kill the monster." & vbNewLine
        .SelText = "5. /echo parser stop dps" & vbNewLine
        .SelText = "6. /clock" & vbNewLine
        .SelText = "7. You messed up, forgot to stop for example." & vbNewLine
        .SelText = "8. /echo parser clear dps" & vbNewLine & vbNewLine
    End If
    
    
    If comboList.ListIndex = 3 Then
        .SelBold = True
        .SelColor = vbRed
        .SelText = "Experience Calculation - Displayed on summary." & vbNewLine
        .SelBold = True
        .SelText = "Commands used:" & vbNewLine
        .SelColor = vbBlue
        .SelText = """" & "start exp" & """" & vbNewLine
        .SelColor = vbBlue
        .SelText = """" & "stop exp" & """" & vbNewLine
        .SelText = """" & "/clock" & """" & vbNewLine & vbNewLine
        
        .SelBold = True
        .SelText = "Example:" & vbNewLine
        .SelText = "1. You puller pulls the first monster." & vbNewLine
        .SelText = "2. /echo parser start exp" & vbNewLine
        .SelText = "3. /clock" & vbNewLine
        .SelText = "4. Your party breaks up." & vbNewLine
        .SelText = "5. /echo parser stop exp" & vbNewLine
        .SelText = "6. /clock" & vbNewLine & vbNewLine
        
        .SelBold = True
        .SelText = "Notes:" & vbNewLine
        .SelText = "When you do the start command the Total Exp on the summary report will reset to 0." & vbNewLine
        .SelText = "The stop command is not required to see the exp calculations, the parser will use the current system time until a stop command is received."
    End If
    
        
    If comboList.ListIndex = 4 Then
        .SelBold = True
        .SelColor = vbRed
        .SelText = "Fishing Report" & vbNewLine
        .SelBold = True
        .SelText = "Commands used:" & vbNewLine
        .SelColor = vbBlue
        .SelText = """" & "start fish" & """" & vbNewLine
        .SelColor = vbBlue
        .SelText = """" & "stop fish" & """" & vbNewLine
        .SelText = """" & "/clock" & """" & vbNewLine
        .SelColor = vbBlue
        .SelText = """" & "fish comment" & """" & vbNewLine & vbNewLine
        
        .SelBold = True
        .SelText = "Example:" & vbNewLine
        .SelText = "1. /echo parser start fish" & vbNewLine
        .SelText = "2. /clock" & vbNewLine
        .SelText = "3. /echo parser fish comment 'Using lugworm'" & " - This of course is optional" & vbNewLine
        .SelText = "4. You fish for as long as you want." & vbNewLine
        .SelText = "5. /echo parser stop fish" & vbNewLine
        .SelText = "6. /clock" & vbNewLine & vbNewLine
        
        .SelBold = True
        .SelText = "Notes:" & vbNewLine
        .SelText = """" & "stop fish" & """" & " is not required, you may start a new cycle with " & """" & "start fish" & """" & ". However if you are done for the day you must use the " & """" & "stop fish" & """" & "." & vbNewLine & vbNewLine
  
    End If
        
    
    If comboList.ListIndex = 5 Then
        .SelBold = True
        .SelColor = vbRed
        .SelText = "Save Reports - Saves specified report to FFXIP folder." & vbNewLine
        .SelBold = True
        .SelText = "Commands used:" & vbNewLine
        .SelColor = vbBlue
        .SelText = """" & "save report" & """" & vbNewLine
        .SelColor = vbBlue
        .SelText = """" & "save summary" & """" & vbNewLine
        .SelColor = vbBlue
        .SelText = """" & "save details" & """" & vbNewLine
        .SelColor = vbBlue
        .SelText = """" & "save player1-6" & """" & vbNewLine & vbNewLine
        
        .SelBold = True
        .SelText = "Examples:" & vbNewLine
        .SelText = "/echo parser save report MyReport.rtf" & vbNewLine & vbNewLine
        .SelText = "/echo parser save summary MySummary.rtf" & vbNewLine & vbNewLine
        .SelText = "/echo parser save details MyDetails.rtf" & vbNewLine & vbNewLine
        .SelText = "/echo parser save player1 Spyle.rtf" & vbNewLine & vbNewLine
    End If
        
    If comboList.ListIndex = 6 Then
        .SelBold = True
        .SelColor = vbRed
        .SelText = "Set Player Details." & vbNewLine
        .SelBold = True
        .SelText = "Commands used:" & vbNewLine
        .SelColor = vbBlue
        .SelText = """" & "player" & """" & vbNewLine & vbNewLine
        
        .SelText = "/echo parser player Name Job SubJob Level" & vbNewLine & vbNewLine
        .SelBold = True
        .SelText = "Examples:" & vbNewLine
        .SelText = "/echo parser player Spyle warrior ninja 75" & vbNewLine & vbNewLine
        .SelText = "/echo parser player Spyle war nin 74-75" & vbNewLine & vbNewLine
        .SelText = "/echo parser player Apricoth bard whitemage 75" & vbNewLine & vbNewLine
        .SelText = "/echo parser player Apricoth brd whm 75" & vbNewLine & vbNewLine
        
    End If
    
    If comboList.ListIndex = 7 Then
        .SelBold = True
        .SelColor = vbRed
        .SelText = "Beep Beep!" & vbNewLine
        .SelBold = True
        .SelText = "Commands used:" & vbNewLine
        .SelColor = vbBlue
        .SelText = """" & "beep" & """" & vbNewLine & vbNewLine
        
        .SelBold = True
        .SelText = "Example:" & vbNewLine
        .SelText = "1. /echo parser beep" & vbNewLine & vbNewLine
        
        .SelBold = True
        .SelText = "Notes:" & vbNewLine
        .SelText = "This will make your computer beep twice, letting you know it has reached this area of the log." & vbNewLine & vbNewLine
    End If
    
    If comboList.ListIndex = 8 Then
        .SelBold = True
        .SelColor = vbRed
        .SelText = "Change Window." & vbNewLine
        .SelBold = True
        .SelText = "Commands used:" & vbNewLine
        .SelColor = vbBlue
        .SelText = """" & "window1-7" & """" & vbNewLine & vbNewLine
        
        .SelBold = True
        .SelText = "Example:" & vbNewLine
        .SelText = "1. /echo parser window 2" & vbNewLine & vbNewLine
        
        .SelBold = True
        .SelText = "Notes:" & vbNewLine
        .SelText = "Changes the current window to Summary. Correlates with Function Keys in View." & vbNewLine & vbNewLine
    End If
    
    If comboList.ListIndex = 9 Then
        .SelBold = True
        .SelColor = vbRed
        .SelText = "Start/Stop Logging." & vbNewLine
        .SelBold = True
        .SelText = "Commands used:" & vbNewLine
        .SelColor = vbBlue
        .SelText = """" & "stop logging" & """" & vbNewLine
        .SelText = """" & "start logging" & """" & vbNewLine & vbNewLine
        
        .SelBold = True
        .SelText = "Example:" & vbNewLine
        .SelText = "1. /echo parser stop logging" & vbNewLine
        .SelText = "2. Do your vewy-wery secwet stuff!" & vbNewLine
        .SelText = "3. /echo parser start logging" & vbNewLine & vbNewLine
        
        .SelBold = True
        .SelText = "Notes:" & vbNewLine
        .SelText = "Once a stop logging command is received it will NOT parse/gather or do ANYTHING with any future lines until a start logging command is received." & vbNewLine & vbNewLine
    End If
    
    If comboList.ListIndex = 10 Then
        .SelBold = True
        .SelColor = vbRed
        .SelText = "Timer Function." & vbNewLine
        .SelBold = True
        .SelText = "Commands used:" & vbNewLine
        .SelColor = vbBlue
        .SelText = """" & "timer ##:##:## '1-10'" & """" & vbNewLine
        .SelText = """" & "/clock" & """" & vbNewLine & vbNewLine
        
        .SelBold = True
        .SelText = "Example Macro 1:" & vbNewLine
        .SelText = "/ja " & """" & "Warcry" & """" & " Spyle" & vbNewLine
        .SelText = "/echo parser timer 00:05:00 '1'" & vbNewLine
        .SelText = "/clock" & vbNewLine & vbNewLine
        .SelItalic = True
        .SelText = "This macro will make the parser play 1 beep or WAV 1, in 5 mins, when my Warcry is ready again" & vbNewLine & vbNewLine
        .SelBold = True
        .SelText = "Example Macro 2:" & vbNewLine
        .SelText = "/ja " & """" & "Mighty Strikes" & """" & " Spyle" & vbNewLine
        .SelText = "/echo parser timer 02:00:00 '5'" & vbNewLine
        .SelText = "/clock" & vbNewLine & vbNewLine
        .SelItalic = True
        .SelText = "This macro will make the parser play 5 beeps or WAV 5, in 2 hrs, when my Mighty Strikes are ready again" & vbNewLine & vbNewLine
        
        .SelBold = True
        .SelText = "Notes:" & vbNewLine
        .SelText = "You MUST use the format ##:##:##, if you simply put 0:05 then it will NOT work. You can have up to 10 timers running at one time." & vbNewLine & vbNewLine
        .SelText = "The '1-10' specifies either the WAV to be played, or the number of beeps. Depending on what option you have selected." & vbNewLine & vbNewLine
        .SelText = "What this command basically does is play a sound when your ability is ready. Of course, you don't have to use it just for that.. " & vbNewLine & vbNewLine
        .SelBold = True
        .SelText = "Important!" & vbNewLine
        .SelText = "If you do something quick, like 00:00:30, then chances are by the time the parser is able to read the log file, those 30 seconds will be long gone, making it useless for abilities that have a short wait."
    End If
    
    If comboList.ListIndex = 11 Then
        .SelBold = True
        .SelColor = vbRed
        .SelText = "Add Comments - Displayed in multiple areas per fight." & vbNewLine
        .SelBold = True
        .SelText = "Commands used:" & vbNewLine
        .SelColor = vbBlue
        .SelText = """" & "comment" & """" & vbNewLine & vbNewLine
        
        .SelBold = True
        .SelText = "Example:" & vbNewLine
        .SelText = "1. At anytime before or during the fight." & vbNewLine
        .SelText = "2. /echo parser comment 'Using Darksteel Axe+1'" & vbNewLine & vbNewLine
        
        .SelBold = True
        .SelText = "Notes:" & vbNewLine
        .SelText = "The ' is require before and after your comments." & vbNewLine & vbNewLine
        
        .SelBold = True
        .SelColor = vbRed
        .SelText = "Clear Parser - Clears all parsed data." & vbNewLine
        .SelBold = True
        .SelText = "Commands used:" & vbNewLine
        .SelColor = vbBlue
        .SelText = """" & "clear" & """" & vbNewLine & vbNewLine
    End If
    
    If comboList.ListIndex = 12 Then
        .SelBold = True
        .SelColor = vbRed
        .SelText = "Change Gather File." & vbNewLine
        .SelBold = True
        .SelText = "Commands used:" & vbNewLine
        .SelColor = vbBlue
        .SelText = """" & "gather" & """" & vbNewLine & vbNewLine
        
        .SelBold = True
        .SelText = "Example:" & vbNewLine
        .SelText = "1. /echo parser gather 'c:\mine.prs'" & vbNewLine & vbNewLine
        
    End If
    
End With
End Sub


Private Sub Form_Load()
Me.Left = frmRead.Left + 100
Me.Top = frmRead.Top + 100
comboList.AddItem "General"
comboList.AddItem "Fight Efficiency"
comboList.AddItem "Damage Per Second Calculation"
comboList.AddItem "Experience Calculation"
comboList.AddItem "Fishing Report"
comboList.AddItem "Save Reports"
comboList.AddItem "Set Player Details"
comboList.AddItem "Beep Beep!"
comboList.AddItem "Change Window"
comboList.AddItem "Start/Stop Logging"
comboList.AddItem "Timer Function"
comboList.AddItem "Other"
comboList.AddItem "Change Gather File"
comboList.ListIndex = 0
End Sub


Private Sub Form_Resize()
RTB_Help.Width = Me.ScaleWidth
RTB_Help.Height = Me.ScaleHeight - 400
comboList.Width = Me.ScaleWidth - 1000
End Sub


