Attribute VB_Name = "modParser"
Public Type udtCrafting
    Result As String
    DayType As String
    MoonPhase As String
    MoonPerc As String
    Count As Integer
    CurrentTime As String
    Direction As String
    CriticalFailure As Boolean
End Type
'Additional sorting features.
'Website to export raw data.

Public Type udtBasics
    Hit As Integer
    Miss As Integer
    Damage As Long
    High As Integer
    Low As Integer
    Uses As Integer
    List As String
    MPCost As Long
    HighSkillType As String
End Type
Public Type udtEvasion
    TotalEvasion As Integer
    Parry As Integer
    Block As Integer
    Absorb As Integer
    Anticipate As Integer
    Evade As Integer
    Miss As Integer
    Hit As Integer
    Damage As Long
End Type
Public Type udtHeal
    Healed As Long
    Recovered As Long
    HealedList As String
    RecoveredList As String
    MPCost As Long
End Type
Public Type udtStatistics
    BattleID As Integer
    Battles As Integer
    Attacker As String
    Defender As String
    Basic As udtBasics
    Ranged As udtBasics
    Counter As udtBasics
    Skill As udtBasics
    Ability As udtBasics
    Critical As udtBasics
    Effect As udtBasics
    Spell As udtBasics
    TotalMeleeHit As Integer
    TotalMeleeMiss As Integer
    TotalRangedHit As Integer
    TotalRangedMiss As Integer
    TotalDMG As Long
    Evasion As udtEvasion
    Heal As udtHeal
    Percent As Currency
    Accuracy As Currency
End Type
Public Type udtEfficiency
    TotalDMG As Long
    BasicDMG As Long
    ATK As Integer
    ATKTaken As Integer
    DMGTaken As Long
End Type
'Too small for HNM
Public BattleStats(50) As udtStatistics
Public BattleTotals As udtStatistics
Public SummaryStats() As udtStatistics
Public FullStats() As udtStatistics
Public EffTotals As udtEfficiency

Public Type udtSpells
    Name As String
    MPCost As Integer
End Type

Public SkillList As String
Public SpellList() As udtSpells
Public CraftingCSV() As udtCrafting

Public ReportOptions(37)
Public ErrorFile, ErrorCount As Integer, ReportError As String

Public TypeDone As String
Public Gather As Boolean, ParseGather As Boolean, GatherDate As Boolean, OpenSingle As Boolean, SingleFile As String, ExportFile As String, ClearEdit As Boolean, NotClearFile As Boolean, BeepSounds(10) As String, StartWithOpen As String, KeyLogs(9, 3) As String
', HiddenAds As Boolean


Public Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Public Const SND_ASYNC = &H1
Public Const SND_FILENAME = &H20000
Private Sub Main()
On Error Resume Next
If Command$ <> "" Then
    StartWithOpen = Replace(Command$, """", "")
End If
frmRead.Show
End Sub










