Option Explicit

Declare Function mciSendString Lib "mmsystem" (ByVal lpstrCommand$, ByVal lpstrReturnStr As String, ByVal wReturnLen%, ByVal hCallBack%) As Long
Declare Function mciGetErrorString Lib "mmsystem" (ByVal dwError&, ByVal lpstrReturnStr As Any, ByVal wReturnLen%) As Long
Declare Function FindWindow Lib "User" (ByVal lpClassName As Any, ByVal lpWindowName As Any) As Integer
Declare Function WinExec Lib "Kernel" (ByVal lpCmdLine As String, ByVal nCmdShow As Integer) As Integer
Declare Function GetModuleUsage Lib "Kernel" (ByVal hModule As Integer) As Integer
Declare Function GetPrivateProfileString Lib "Kernel" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
Declare Function WritePrivateProfileString Lib "Kernel" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lplFileName As String) As Integer
Declare Function sndPlaySound Lib "MMSYSTEM" (ByVal lpszSoundName As Any, ByVal uFlags As Integer) As Integer
Declare Function SendMessage Lib "User" (ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As Any) As Long
Declare Function GetWindowLong Lib "User" (ByVal hWnd As Integer, ByVal nIndex As Integer) As Long
Declare Function SetWindowLong Lib "User" (ByVal hWnd As Integer, ByVal nIndex As Integer, ByVal dwNewLong As Long) As Long
Declare Function GetWindowsDirectory Lib "Kernel" (ByVal lpBuffer As String, ByVal nSize As Integer) As Integer
Declare Function GetSystemDirectory Lib "Kernel" (ByVal lpBuffer As String, ByVal nSize As Integer) As Integer
Declare Sub SetWindowPos Lib "User" (ByVal hWnd As Integer, ByVal hWndInsertAfter As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal CX As Integer, ByVal CY As Integer, ByVal wFlags As Integer)
Declare Sub BringWindowToTop Lib "User" (ByVal hWnd As Integer)


Const GWL_STYLE = (-16)
Const WS_DLGFRAME = &H400000
Const WS_SYSMENU = &H80000
Const WS_MINIMIZEBOX = &H20000
Const WS_MAXIMIZEBOX = &H10000
Const SWP_NOSIZE = &H1
Const SWP_NOZORDER = &H4
Const SWP_NOMOVE = &H2
Const SWP_DRAWFRAME = &H20
Const WS_THICKFRAME = &H40000
Const SND_ASYN = 1
Const SND_NODEFAULT = 2
Const SND_NOSTOP = &H10
Const WM_USER = &H400
Const EM_SETREADONLY = (WM_USER + 31)

Global Const BLACK = &H0&
Global Const RED = &HFF&
Global Const GREEN = &HFF00&
Global Const YELLOW = &HFFFF&
Global Const BLUE = &HFF0000
Global Const MAGENTA = &HFF00FF
Global Const CYAN = &HFFFF00
Global Const WHITE = &HFFFFFF
Global Const GREY = &HC0C0C0


Type ProgRec
    Prog As String
    folder As String
    Text As String
    pic As String
End Type

Global ProgArray() As ProgRec

Global Msgx As String
Global XFlag As Integer
Global Root As String
Global CrLf  As String
Global LookFor As String
Global CancelSearch As Integer
Global Go_Down_On_Run As Integer
Global sGAuto_Version As String

Global IsAVI As Integer
Global Const WS_CHILD = &H40000000
Global h%

Sub CenterForm (frm As Form)
    frm.Left = (screen.Width - frm.Width) / 2
    frm.Top = (screen.Height - frm.Height) / 2
End Sub

Sub CloseAVI ()
Dim ret
Dim retstr

ret = mciSendString("close Animation", retstr, 0, 0)
End Sub

Sub Encrypt (Secret$, PassWord$)
' Secret$ = the string you wish to encrypt or decrypt.
' PassWord$ = the password with which to encrypt the string.
' Calling the sub on a string Encrypts it
' Calling the sub on an encrypted  string Decrypts it

Dim L As Long
Dim X As Long
Dim iChar As Integer

    L = Len(PassWord$)
    For X = 1 To Len(Secret$)
        iChar = Asc(Mid$(PassWord$, (X Mod L) - L * ((X Mod L) = 0), 1))
        Mid$(Secret$, X, 1) = Chr$(Asc(Mid$(Secret$, X, 1)) Xor iChar)
    Next X

End Sub

Sub GetBitmapInfo (bmpFileName, vRes As Long, hRes As Long, bits As Integer)
Dim i As Integer
i = FreeFile
Open bmpFileName For Binary As i
Get i, 19, hRes  'vertical resolution
Get i, 23, vRes  'horizontal resolution
Get i, 29, bits  'bits-per-pixel
Close i

End Sub

Function GetFromINI (Section$, KeyName$, FileName$) As String

' [SECTION]
' KEYNAME = 'function return'
' FILENAME.INI
 
   Dim retstr As String
   retstr = String(255, Chr(0))
   GetFromINI = Left(retstr, GetPrivateProfileString(Section$, ByVal KeyName$, "", retstr, Len(retstr), FileName$))
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

Sub MakeTextBoxReadOnly (txt As TextBox)
Dim RetVal As Integer
RetVal = SendMessage(txt.hWnd, EM_SETREADONLY, True, ByVal 0&)

End Sub

Sub PlayAVI (sAviFile, vHeight%, pix As Form)
 
    Dim sRet$
    Dim ErrorStr
    Dim retstr
    Dim CmdStr
    Dim ret
    Dim strr$

   '*** This will open the AVIVideo and create a child window on the
   '*** form where the video will display. Animation is the device_id.
 
    ErrorStr = Space(255)
    retstr = Space(255)
 
    On Error GoTo CloseAVI
 

 
    CmdStr = ("open " & sAviFile & " type AVIVideo alias Animation parent " + LTrim$(Str$(pix.hWnd)) + " style " + Trim$(Str$(WS_CHILD)))
    ret = mciSendString(CmdStr, retstr, 0, 0)
    strr$ = "put Animation window at 0 0 " + Format(0) + " " + Format(vHeight%)
    ret = mciSendString(strr$, retstr, 0, 0)
    ret = mciSendString("play Animation ", retstr, 0, 0)
  
Exit Sub

CloseAVI:  'Error Handling

CloseAVI
Exit Sub

End Sub

Sub PlayAVIFullScreen (vFileName)
    Dim CmdStr$, ReturnVal&
    CmdStr$ = "play " + vFileName + " fullscreen "
    ReturnVal& = mciSendString(CmdStr$, "", 0, 0&)
    ReturnVal& = mciSendString("close " + vFileName, "", 0, 0)

End Sub

Sub PlayWave16 (FileName As String)
' this function is no longer used
Dim iResult As Integer

    iResult = sndPlaySound(FileName, SND_ASYN Or SND_NODEFAULT)

End Sub

Sub RemoveTitleBar (frm As Form)

Static OriginalStyle As Long
Dim CurrentStyle As Long
Dim X As Long
    OriginalStyle = 0
    CurrentStyle = GetWindowLong(frm.hWnd, GWL_STYLE)
    OriginalStyle = OriginalStyle Or (CurrentStyle And WS_DLGFRAME)
    OriginalStyle = OriginalStyle Or (CurrentStyle And WS_SYSMENU)
    OriginalStyle = OriginalStyle Or (CurrentStyle And WS_MINIMIZEBOX)
    OriginalStyle = OriginalStyle Or (CurrentStyle And WS_MAXIMIZEBOX)
    CurrentStyle = CurrentStyle And Not WS_DLGFRAME
    CurrentStyle = CurrentStyle And Not WS_SYSMENU
    CurrentStyle = CurrentStyle And Not WS_MINIMIZEBOX
    CurrentStyle = CurrentStyle And Not WS_MAXIMIZEBOX
    X = SetWindowLong(frm.hWnd, GWL_STYLE, CurrentStyle)
    frm.Width = frm.Width  'to refresh (Refresh didn't work)


End Sub

Sub ResizeControl (ControlName As Control, FormName As Form)
' Works with the following controls:
' TextBox, Picture, List, Command Button...
Dim NewStyle As Long
    
    NewStyle = GetWindowLong(ControlName.hWnd, GWL_STYLE)
    NewStyle = NewStyle Or WS_THICKFRAME
    NewStyle = SetWindowLong(ControlName.hWnd, GWL_STYLE, NewStyle)
    SetWindowPos ControlName.hWnd, FormName.hWnd, 0, 0, 0, 0, SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_DRAWFRAME

End Sub

Sub StopWave16 ()
Dim iResult As Integer

   'iResult = sndPlaySound(ByVal 0&, 0)
   CloseAVI
End Sub

Function WriteINI (Section As String, Key As String, Value As String, inipath As String) As Integer
'-This function takes four variables and writes to the ini file the information
'-Section is section name
'-Key is key name
'-Value is what is on the right side of the equal sign
'-inipath is file name
 
  Dim r%
   r% = WritePrivateProfileString(Section, Key, Value, inipath)
   WriteINI = r%
End Function

