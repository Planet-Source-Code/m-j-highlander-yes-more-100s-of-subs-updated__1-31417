Option Explicit
Declare Function GetVersion Lib "Kernel" () As Long
Declare Function GetWindowsDirectory Lib "Kernel" (ByVal lpBuffer As String, ByVal nSize As Integer) As Integer
Declare Function GetSystemDirectory Lib "Kernel" (ByVal lpBuffer As String, ByVal nSize As Integer) As Integer

Global Const ATTR_READONLY = 1    'Read-only file
Global Const ATTR_VOLUME = 8  'Volume label
Global Const ATTR_ARCHIVE = 32    'File has changed since last back-up
Global Const ATTR_NORMAL = 0  'Normal files
Global Const ATTR_HIDDEN = 2  'Hidden files
Global Const ATTR_SYSTEM = 4  'System files
Global Const ATTR_DIRECTORY = 16  'Directory

Global Const ATTR_ALL_FILES = ATTR_NORMAL Or ATTR_HIDDEN Or ATTR_SYSTEM Or ATTR_READONLY Or ATTR_ARCHIVE


Function DOSVer () As String
End Function

Function FileExists (sFile As String) As Integer
Dim tmp As String
Dim iResult As Integer

iResult = 0
If Dir$(sFile, ATTR_ALL_FILES) <> "" Then
    'iResult = GetAttr(sFile) And ATTR_ALL_FILES
    FileExists = True
Else
    FileExists = False

End If

'If iResult = 0 Then   'file not found
'    FileExists = False
'Else
'    FileExists = True
'End If


End Function

Function GetSysDir () As String
    
    Dim SysDir As String
    Dim File As String
    Dim Res As Integer
    SysDir = Space$(20)
    Res = GetSystemDirectory(SysDir, 20)
    File = Left$(SysDir, InStr(1, SysDir, Chr$(0)) - 1)
    GetSysDir = Trim$(File) & "\"
    
End Function

Function GetWinDir () As String
    
    Dim WinDir As String
    Dim File As String
    Dim Res As Integer
    WinDir = Space$(20)
    Res = GetWindowsDirectory(WinDir, 20)
    File = Left$(WinDir, InStr(1, WinDir, Chr$(0)) - 1)
    GetWinDir = Trim$(File) & "\"
    
End Function

Function WinVer () As String
'Returns 3.95 in Windows95

Dim lTemp As Long
Dim lWinVer As Long

lTemp = GetVersion()
lWinVer = lTemp And &HFFFF&
WinVer = Format((lWinVer Mod 256) + ((lWinVer \ 256) / 100), "Fixed")

End Function

