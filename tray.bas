Attribute VB_Name = "Module1"
Public Const WM_LBUTTONDOWN = &H201
' // Tray notification definitions
Global IconIndex As Integer
Global FlashOn As Boolean

Type NOTIFYICONDATA
        cbSize As Long
        hwnd As Long
        uID As Long
        uFlags As Long
        uCallbackMessage As Long
        hIcon As Long
        szTip As String * 64
End Type

Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2

Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4

Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Type SHFILEINFO
        hIcon As Long                      '  out: icon
        iIcon As Long          '  out: icon index
        dwAttributes As Long               '  out: SFGAO_ flags
        szDisplayName As String * 255 'MAX_PATH '  out: display name (or path)
        szTypeName As String * 80         '  out: type name
End Type

Public Const SHGFI_ICON = &H100                         '  get icon
Public Const SHGFI_DISPLAYNAME = &H200                  '  get display name
Public Const SHGFI_TYPENAME = &H400                     '  get type name
Public Const SHGFI_ATTRIBUTES = &H800                   '  get attributes
Public Const SHGFI_ICONLOCATION = &H1000                '  get icon location
Public Const SHGFI_EXETYPE = &H2000                     '  return exe type
Public Const SHGFI_SYSICONINDEX = &H4000                '  get system icon index
Public Const SHGFI_LINKOVERLAY = &H8000                 '  put a link overlay on icon
Public Const SHGFI_SELECTED = &H10000                   '  show icon in selected state
Public Const SHGFI_LARGEICON = &H0                      '  get large icon
Public Const SHGFI_SMALLICON = &H1                      '  get small icon
Public Const SHGFI_OPENICON = &H2                       '  get open icon
Public Const SHGFI_SHELLICONSIZE = &H4                  '  get shell size icon
Public Const SHGFI_PIDL = &H8                           '  pszPath is a pidl
Public Const SHGFI_USEFILEATTRIBUTES = &H10             '  use passed dwFileAttribute

Declare Function SHGetFileInfo Lib "shell32.dll" Alias " SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long

Declare Function SHGetNewLinkInfo Lib "shell32.dll" Alias "SHGetNewLinkInfoA" (ByVal pszLinkto As String, ByVal pszDir As String, ByVal pszName As String, pfMustCopy As Long, ByVal uFlags As Long) As Long

Public Const SHGNLI_PIDL = &H1                          '  pszLinkTo is a pidl
Public Const SHGNLI_PREFIXNAME = &H2                    '  Make name "Shortcut to xxx"

' // End SHGetFileInfo
'Get the menu handle
Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
'make it pop up
Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal x As Long, ByVal y As Long, ByVal nReserved As Long, ByVal hwnd As Long, lpRect As Any) As Long
' Flags for TrackPopupMenu
Public Const TPM_LEFTBUTTON = &H0&
Public Const TPM_RIGHTBUTTON = &H2&
Public Const TPM_LEFTALIGN = &H0&
Public Const TPM_CENTERALIGN = &H4&
Public Const TPM_RIGHTALIGN = &H8&

Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As Any, ByVal lpWindowName As Any) As Long
Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

'Tray messages
Public Const TRAY_MSG_MOUSEMOVE = 7680
Public Const TRAY_MSG_LEFTBTN_DOWN = 7695
Public Const TRAY_MSG_LEFTBTN_UP = 7710
Public Const TRAY_MSG_LEFTBTN_DBLCLICK = 7725
Public Const TRAY_MSG_RIGHTBTN_DOWN = 7740
Public Const TRAY_MSG_RIGHTBTN_UP = 7755
Public Const TRAY_MSG_RIGHTBTN_DBLCLICK = 7770



