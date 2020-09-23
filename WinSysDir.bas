Declare Function GetWindowsDirectory Lib "Kernel" (ByVal lpBuffer As String, ByVal nSize As Integer) As Integer

Declare Function GetSystemDirectory Lib "Kernel" (ByVal lpBuffer As String, ByVal nSize As Integer) As Integer

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

