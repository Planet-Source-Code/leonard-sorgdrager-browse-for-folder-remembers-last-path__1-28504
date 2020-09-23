Attribute VB_Name = "BrowseFolders"
Option Explicit


Public strDirName As String
Public counter As Integer
'Public Directory
'Public looped As Boolean

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2


Private Const MAX_PATH = 260

Private Declare Function SHBrowseForFolder Lib _
"shell32" (lpbi As BrowseInfo) As Long

Private Declare Function SHGetPathFromIDList Lib _
"shell32" (ByVal pidList As Long, ByVal lpBuffer _
As String) As Long

Private Declare Function lstrcat Lib "kernel32" _
Alias "lstrcatA" (ByVal lpString1 As String, ByVal _
lpString2 As String) As Long

Private Type BrowseInfo
   hWndOwner As Long
   pidlRoot As Long
   pszDisplayName As Long
   lpszTitle As Long
   ulFlags As Long
   lpfnCallback As Long
   lParam As Long
   iImage As Long
End Type

Public Const LMEM_FIXED = &H0   'added
Public Const LMEM_ZEROINIT = &H40   'added
Public Const LPTR = (LMEM_FIXED Or LMEM_ZEROINIT)   'added

Public Declare Function LocalAlloc Lib "kernel32" _
 (ByVal uFlags As Long, _
    ByVal uBytes As Long) As Long   'added

Public Declare Function LocalFree Lib "kernel32" _
  (ByVal hMem As Long) As Long  'added

Public Declare Function lstrcpyA Lib "kernel32" _
  (lpString1 As Any, lpString2 As Any) As Long  'added

Public Declare Function lstrlenA Lib "kernel32" _
(lpString As Any) As Long   'added

'windows-defined type OSVERSIONINFO
'Public Type OSVERSIONINFO   'added
'  OSVSize         As Long
'  dwVerMajor      As Long
'  dwVerMinor      As Long
'  dwBuildNumber   As Long
'  PlatformID      As Long
'  szCSDVersion    As String * 128
'End Type
'
'Public Const VER_PLATFORM_WIN32_NT = 2  'added
'Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" _
' (lpVersionInformation As OSVERSIONINFO) As Long    'added
 
 Public Declare Sub CopyMemory Lib "kernel32" _
    Alias "RtlMoveMemory" _
   (pDest As Any, _
    pSource As Any, _
    ByVal dwLength As Long) 'added

Public Const BFFM_INITIALIZED = 1   'added
Public Const WM_USER = &H400    'added

'Selects the specified folder. If the wParam
'parameter is FALSE, the lParam parameter is the
'PIDL of the folder to select , or it is the path
'of the folder if wParam is the C value TRUE (or 1).
'Note that after this message is sent, the browse
'dialog receives a subsequent BFFM_SELECTIONCHANGED
'message.
Public Const BFFM_SETSELECTIONA As Long = (WM_USER + 102)   'added
Public Const BFFM_SETSELECTIONW As Long = (WM_USER + 103)   'added

Public Declare Function SendMessage Lib "user32" _
   Alias "SendMessageA" _
   (ByVal hWnd As Long, _
   ByVal wMsg As Long, _
   ByVal wParam As Long, _
   lParam As Any) As Long   'added

Public Function Browse_Folder(sPath As String) As String    'added
Dim lpIDList As Long 'Declare Varibles
Dim sBuffer As String
Dim szTitle As String
Dim lpPath As Long  'added
Dim tBrowseInfo As BrowseInfo

szTitle = "Click on the directory containing " & _
"the images you wish to include in the Slide Show"
'Text to appear in the the gray area under the title bar
'telling you what to do

With tBrowseInfo
   .hWndOwner = Form1.hWnd 'Owner Form
   .pidlRoot = 0
   .lpszTitle = lstrcat(szTitle, "")
   .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
   .lpfnCallback = FARPROC(AddressOf BrowseCallbackProcStr) 'added
   
    lpPath = LocalAlloc(LPTR, Len(sPath) + 1)   'added
    CopyMemory ByVal lpPath, ByVal sPath, Len(sPath) + 1    'added
    .lParam = lpPath    'added
End With

lpIDList = SHBrowseForFolder(tBrowseInfo)

If (lpIDList) Then
   sBuffer = Space(MAX_PATH)
   SHGetPathFromIDList lpIDList, sBuffer
   sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
End If

strDirName = sBuffer
End Function

Public Function BrowseCallbackProcStr(ByVal hWnd As Long, _
                                      ByVal uMsg As Long, _
                                      ByVal lParam As Long, _
                                      ByVal lpData As Long) As Long

  'Callback for the Browse STRING method.

  'On initialization, set the dialog's
  'pre-selected folder from the pointer
  'to the path allocated as bi.lParam,
  'passed back to the callback as lpData param.

   Select Case uMsg
      Case BFFM_INITIALIZED

         Call SendMessage(hWnd, BFFM_SETSELECTIONA, _
                          True, ByVal lpData)

         Case Else:

   End Select

End Function

Public Function FARPROC(pfn As Long) As Long

  'A dummy procedure that receives and returns
  'the value of the AddressOf operator.

  'This workaround is needed as you can't assign
  'AddressOf directly to a member of a user-
  'defined type, but you can assign it to another
  'long and use that (as returned here)
  FARPROC = pfn

End Function


