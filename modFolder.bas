Attribute VB_Name = "BrowseForFolder"
Option Explicit

Public Declare Function SHBrowseForFolder Lib "shell32.dll" _
(lpBrowseInfo As BROWSEINFO) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32.dll" _
(ByVal pidl As Long, ByVal pszPath As String) As Long
Public Type BROWSEINFO
    howner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    lImage As Long
End Type
Public bi As BROWSEINFO
Public pidl As Long
Public gMyFolder As String

'release the memory used by the browse for folder
Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)
Public Const LMEM_FIXED = &H0
Public Const LMEM_ZEROINIT = &H40
Public Const LPTR = (LMEM_FIXED Or LMEM_ZEROINIT)

'send a message to the browse for folder window
Public Declare Function SendMessage Lib "user32" _
   Alias "SendMessageA" _
   (ByVal hWnd As Long, _
   ByVal wMsg As Long, _
   ByVal wParam As Long, _
   lParam As Any) As Long
Public Const BFFM_INITIALIZED = 1
Public Const BFFM_SELECTIONCHANGED = 2

'allocate and free space for the folder parameter
' that is to be passed to browse for folder
Public Declare Function LocalAlloc Lib "kernel32" _
   (ByVal uFlags As Long, _
    ByVal uBytes As Long) As Long
Public Declare Function LocalFree Lib "kernel32" _
   (ByVal hMem As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" _
   Alias "RtlMoveMemory" _
   (pDest As Any, _
    pSource As Any, _
    ByVal dwLength As Long)

Public Const MAX_PATH = 260
Public Const WM_USER = &H400
Public Const BFFM_SETSELECTIONA As Long = (WM_USER + 102)
Public Const BFFM_SETSELECTIONW As Long = (WM_USER + 103)


Private Function BrowseCallbackProcStr(ByVal hWnd As Long, _
                                   ByVal uMsg As Long, _
                                   ByVal lParam As Long, _
                                   ByVal lpData As Long) As Long
'Called from the browse for folder window
'Sets the initial path to whatever has already been set
   Select Case uMsg
      Case BFFM_INITIALIZED
         Call SendMessage(hWnd, BFFM_SETSELECTIONA, _
            True, ByVal lpData)
         Case Else:
   End Select
End Function

Private Function FARPROC(ByVal pfn As Long) As Long
  'A dummy procedure that receives and
  '   returns the return value of the AddressOf operator.
  'Used to get a pointer (AddressOf) to the call back routine.
   FARPROC = pfn
End Function

Public Function GetFolderPath(frm As Form, Optional DefaultFolder As String = "C:\") As String
Dim lpSelPath As Long
Dim sPath As String * MAX_PATH
Dim pidl As Long
Dim iNull As Integer
Dim strFolderPath As String

    'Get the folder required.
    'Allocate it in some memory, with a pointer to it
    'sPath = "C:\Program Files\"
    If Right(DefaultFolder, 1) <> "\" Then
        sPath = DefaultFolder & "\"
    Else
        sPath = DefaultFolder
    End If
    
    lpSelPath = LocalAlloc(LPTR, Len(sPath) + 1)
    CopyMemory ByVal lpSelPath, ByVal sPath, Len(sPath) + 1

    With bi
        If IsNumeric(frm.hWnd) Then .howner = frm.hWnd
        .pidlRoot = 0
        .lpfn = FARPROC(AddressOf BrowseCallbackProcStr)
        .lParam = lpSelPath
        .lpszTitle = "Select FFXI Log Folder:" & Chr$(0)
    End With

    pidl = SHBrowseForFolder(bi)
    If pidl Then
        If SHGetPathFromIDList(pidl, sPath) Then
            strFolderPath = Trim(sPath)
            If InStr(strFolderPath, vbNullChar) Then
                strFolderPath = Left$(strFolderPath, Len(strFolderPath) - 1)
            End If
        End If
        CoTaskMemFree pidl
    End If
    LocalFree lpSelPath

'   strFolderPath now holds the path and folder actually selected
    GetFolderPath = strFolderPath
End Function


