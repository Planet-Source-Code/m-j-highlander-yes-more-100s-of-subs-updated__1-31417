Global gCancelSearch As Integer

Declare Function GetPrivateProfileString Lib "Kernel" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
 
Declare Function WritePrivateProfileString Lib "Kernel" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lplFileName As String) As Integer

Declare Function GetWindowsDirectory Lib "Kernel" (ByVal lpBuffer As String, ByVal nSize As Integer) As Integer

Declare Function GetSystemDirectory Lib "Kernel" (ByVal lpBuffer As String, ByVal nSize As Integer) As Integer

Declare Function RegCreateKey& Lib "SHELL.DLL" (ByVal hKey&, ByVal lpszSubKey$, lphKey&)

Declare Function RegSetValue& Lib "SHELL.DLL" (ByVal hKey&, ByVal lpszSubKey$, ByVal fdwType&, ByVal lpszValue$, ByVal dwLength&)

Const ERROR_SUCCESS = 0&
Const ERROR_BADDB = 1&
Const ERROR_BADKEY = 2&
Const ERROR_CANTOPEN = 3&
Const ERROR_CANTREAD = 4&
Const ERROR_CANTWRITE = 5&
Const ERROR_OUTOFMEMORY = 6&
Const ERROR_INVALID_PARAMETER = 7&
Const ERROR_ACCESS_DENIED = 8&

Const HKEY_CLASSES_ROOT = 1
Const MAX_PATH = 256&
Const REG_SZ = 1

Global gFullFileName     As String

Sub CenterForm (frm As Form)
    frm.Left = (screen.Width - frm.Width) / 2
    frm.Top = (screen.Height - frm.Height) / 2
End Sub

Function DeleteKey (Section As String, Key As String, inipath As String) As Integer
 
  Dim r%
   r% = WritePrivateProfileString(Section, Key, ByVal 0&, inipath)
   DeleteKey = r%
End Function

Function extractfilename (FileName As String) As String
    
'Extract the File name from a full file name


    Dim pos As Integer
    Dim PrevPos As Integer

    pos = InStr(FileName, "\")
    If pos = 0 Then
        extractfilename = ""
        Exit Function
    End If
    
    Do While pos <> 0
        PrevPos = pos
        pos = InStr(pos + 1, FileName, "\")
    Loop

    extractfilename = Right(FileName, Len(FileName) - PrevPos)

End Function

Function GetFromINI (Section$, KeyName$, FileName$) As String
' Section is the information in the brackets
' Keyname is the information under the Section
' [SECTION]
' KEYNAME = ???
' Filename is the path and name of the ini file
 
   Dim RetStr As String
   RetStr = String(255, Chr(0))
   GetFromINI = Left(RetStr, GetPrivateProfileString(Section$, ByVal KeyName$, "", RetStr, Len(RetStr), FileName$))
End Function

Sub GetSettings ()
Dim wd$, FontName$, FontColor$, FontSize$, BackColor$
wd = GetWinDir()
If Dir(wd + "tv.ini") = "" Then Exit Sub
FontName = GetFromINI("Options", "FontName", wd + "tv.ini")
FontColor = GetFromINI("Options", "FontColor", wd + "tv.ini")
FontSize = GetFromINI("Options", "FontSize", wd + "tv.ini")
BackColor = GetFromINI("Options", "BackColor", wd + "tv.ini")
SelStyle = GetFromINI("Options", "SelectionStyle", wd + "tv.ini")

form1.Editor1.FontName = FontName
form1.Editor1.ForeColor = Val(FontColor)
form1.Editor1.FontSize = Val(FontSize)
form1.Editor1.BackColor = Val(BackColor)

On Error Resume Next
form1.Editor1.SelDefaultType = Val(SelStyle)
If Err Then
        form1.ns.Checked = True
        form1.sl.Checked = False
        form1.sb.Checked = False
        form1.Editor1.SelDefaultType = 1
        Exit Sub
End If
On Error GoTo 0
Select Case Val(SelStyle)
    Case 1 'normal
        form1.ns.Checked = True
        form1.sl.Checked = False
        form1.sb.Checked = False
    Case 2 'line
        form1.sl.Checked = True
        form1.ns.Checked = False
        form1.sb.Checked = False
    Case 3 'block
        form1.sl.Checked = False
        form1.ns.Checked = False
        form1.sb.Checked = True
        
End Select

End Sub

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

Sub RegAssociateApp (RootKey$, RootVal$, ext$, ExePath$)

' sample call:
' RegAssociateApp "TestApp", "VB3 Test Assoc App", ".foo","c:\txt.exe %1"

' This will create a root key called "TestApp"
' with a value of "VB3 Test Assoc App"
' then create a root key called ".foo"
' with a value of "TestApp"
' then add a sub-key to "TestApp" providing the command line

' As a result, files with an extension of ".foo" will have the description
' "VB3 Test Assoc App" and will be associated with "c:\txt.exe"


     Dim sKeyName As String   'Holds Key Name in registry.
     Dim sKeyValue As String  'Holds Key Value in registry.
     Dim ret&                 'Holds error status if any from API calls.
     Dim lphKey&              'Holds created key handle from RegCreateKey.
     
     
     sKeyName = RootKey$
     sKeyValue = RootVal$
     ret& = RegCreateKey&(HKEY_CLASSES_ROOT, sKeyName, lphKey&)
     ret& = RegSetValue&(lphKey&, "", REG_SZ, sKeyValue, 0&)
     
     
     sKeyName = ext$
     sKeyValue = RootKey$
     ret& = RegCreateKey&(HKEY_CLASSES_ROOT, sKeyName, lphKey&)
     ret& = RegSetValue&(lphKey&, "", REG_SZ, sKeyValue, 0&)
     
     

     sKeyName = RootKey$
     sKeyValue = ExePath$
     ret& = RegCreateKey&(HKEY_CLASSES_ROOT, sKeyName, lphKey&)
     ret& = RegSetValue&(lphKey&, "shell\open\command", REG_SZ, sKeyValue, MAX_PATH)


End Sub

Sub SaveSettings ()
Dim wd$, FontName$, FontColor$, FontSize$, BackColor$, SelStyl$
wd = GetWinDir()
FontName = form1.Editor1.FontName
FontColor = form1.Editor1.ForeColor
FontSize = CStr(form1.Editor1.FontSize)
BackColor = CStr(form1.Editor1.BackColor)
SelStyl = CStr(form1.Editor1.SelDefaultType)
w% = WriteINI("Options", "FontName", FontName, wd + "tv.ini")
w% = WriteINI("Options", "FontColor", FontColor, wd + "tv.ini")
w% = WriteINI("Options", "FontSize", FontSize, wd + "tv.ini")
w% = WriteINI("Options", "BackColor", BackColor, wd + "tv.ini")
w% = WriteINI("Options", "SelectionStyle", SelStyl, wd + "tv.ini")
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

