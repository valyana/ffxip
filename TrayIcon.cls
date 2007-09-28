VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TrayIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Contains all functions needed to tray an icon"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const WM_MOUSEMOVE = &H200
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4

Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205

Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Private theTray As NOTIFYICONDATA

Private m_Tooltip As String
Private m_Icon As StdPicture
Private m_Visible As Boolean

Public Event MouseDown(ByVal button As Integer)
Public Event MouseUp(ByVal button As Integer)
Public Event DblClick()

Public WithEvents OwnerForm As Form
Attribute OwnerForm.VB_VarHelpID = -1

Public Property Let Visible(ByVal blnData As Boolean)
    m_Visible = blnData
    
    If m_Visible = False Then
        Shell_NotifyIcon NIM_DELETE, theTray
    Else
        With theTray
            .cbSize = Len(theTray)
            .hIcon = m_Icon
            .hwnd = OwnerForm.hwnd
            .szTip = m_Tooltip
            .ucallbackMessage = WM_MOUSEMOVE
            .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
            .uId = 1&
        End With
        Shell_NotifyIcon NIM_ADD, theTray
    End If
End Property

Public Property Get Visible() As Boolean
    Visible = m_Visible
End Property

Public Property Let Icon(ByVal vData As StdPicture)
    Set m_Icon = vData
    'if icon is already in tray then refresh
    If m_Visible = True Then
        Modify
    End If
End Property

Public Property Get Icon() As StdPicture
    Set Icon = m_Icon
End Property

Public Property Let Tooltip(ByVal vData As String)
    'Add Null to the Tooltip
    m_Tooltip = vData & vbNullChar
    'if icon is already in tray then refresh
    If m_Visible = True Then
        Modify
    End If
End Property

Public Property Get Tooltip() As String
Attribute Tooltip.VB_UserMemId = 0
    'Strip Null
    Tooltip = Left$(m_Tooltip, Len(m_Tooltip) - 1)
End Property

Private Sub Modify()
    With theTray
        .cbSize = Len(theTray)
        .hIcon = m_Icon
        .hwnd = OwnerForm.hwnd
        .szTip = m_Tooltip
        .ucallbackMessage = WM_MOUSEMOVE
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uId = 1&
    End With
    Shell_NotifyIcon NIM_MODIFY, theTray
End Sub

Private Sub OwnerForm_MouseMove(button As Integer, Shift As Integer, X As Single, Y As Single)
    Static rec As Boolean, MSG As Long
    MSG = X / Screen.TwipsPerPixelX
    If rec = False Then
        rec = True
        Select Case MSG
            Case WM_LBUTTONDBLCLK:
                RaiseEvent DblClick
            Case WM_LBUTTONDOWN:
                RaiseEvent MouseDown(vbLeftButton)
            Case WM_LBUTTONUP:
                RaiseEvent MouseUp(vbLeftButton)
            Case WM_RBUTTONDBLCLK:
                RaiseEvent DblClick
            Case WM_RBUTTONDOWN:
                RaiseEvent MouseDown(vbRightButton)
            Case WM_RBUTTONUP:
                RaiseEvent MouseUp(vbRightButton)
        End Select
        rec = False
    End If
End Sub

