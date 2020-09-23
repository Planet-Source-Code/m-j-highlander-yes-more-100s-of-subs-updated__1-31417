Option Explicit
Declare Function WinExec Lib "Kernel" (ByVal lpCmdLine As String, ByVal nCmdShow As Integer) As Integer
Declare Function GetModuleUsage% Lib "Kernel" (ByVal hModule%)



Global Const SW_SHOWMINIMIZED = 2
Global Const SW_SHOWNORMAL = 1
Global Const SW_SHOWMINNOACTIVE = 7
Global Const SW_SHOWNOACTIVATE = 4




'Global Const ATTR_ALL_FILES = ATTR_NORMAL Or ATTR_HIDDEN Or ATTR_SYSTEM


Declare Function WinExec% Lib "Kernel" (ByVal lpCmdLine$, ByVal nCmdShow%)
Global Const SW_HIDE = 0



Sub ExecWait (CmdLine As String, ShowState As Integer)


Dim hMod As Integer
Dim iResult As Integer

hMod = WinExec(CmdLine, ShowState)
If hMod < 32 Then
    'couldn't start program so...
    Exit Sub
End If

Do
    DoEvents
    iResult = GetModuleUsage(hMod)
Loop Until iResult = 0

End Sub

Function ExtractDirName (FileName As String) As String

'Extract the Directory name from a full file name
'The return will have "\" appended to the end

    Dim Pos As Integer
    Dim PrevPos As Integer

    Pos = InStr(FileName, "\")
    If Pos = 0 Then
        ExtractDirName = ""
        Exit Function
    End If
    
    Do While Pos <> 0
        PrevPos = Pos
        Pos = InStr(Pos + 1, FileName, "\")
    Loop

    ExtractDirName = Left(FileName, PrevPos)

End Function

Function ExtractFileBaseName (FileName As String)
'Extract the File base name (without extension) from a file name


    Dim Pos As Integer
    Dim PrevPos As Integer

    Pos = InStr(FileName, ".")
    If Pos = 0 Then
        ExtractFileBaseName = FileName
    Else
        ExtractFileBaseName = Left(FileName, Pos - 1)
    End If

End Function

Function ExtractFileName (FileName As String) As String
    
'Extract the File name from a full file name


    Dim Pos As Integer
    Dim PrevPos As Integer

    Pos = InStr(FileName, "\")
    If Pos = 0 Then
        ExtractFileName = ""
        Exit Function
    End If
    
    Do While Pos <> 0
        PrevPos = Pos
        Pos = InStr(Pos + 1, FileName, "\")
    Loop

    ExtractFileName = Right(FileName, Len(FileName) - PrevPos)

End Function

Sub GetCmdLineArgs (CmdLine, Args())
Dim ai, ch, i
ReDim Args(1 To 10) 'more than enough

ai = 1
For i = 1 To Len(CmdLine)
    ch = Mid(CmdLine, i, 1)
    Select Case ch
        Case ","
        If Args(ai) <> "" Then ai = ai + 1
        Case Else
        'do nothing
    End Select
    If ch = "," Then ch = ""
    Args(ai) = Args(ai) + ch
Next i

ReDim Preserve Args(1 To ai)

'optional
For i = 1 To ai
    Args(i) = LCase(Args(i))
Next i

End Sub

Function SmartConvert (BmpName As String) As String
Dim vRes&, hRes&, bits%
Dim xnCmdLine As String
Dim InFile$, OutFile$

InFile = GetAppPath() & BmpName
GetBitmapInfo InFile, vRes, hRes, bits

If bits <= 8 Then
    OutFile = GetAppPath() & ExtractFileBaseName(BmpName) + ".gif "
    xnCmdLine = GetAppPath() & "NConvert -quiet -out 13 -o " & OutFile & InFile
Else
    OutFile = GetAppPath() & ExtractFileBaseName(BmpName) + ".jpg "
    xnCmdLine = GetAppPath() & "NConvert -quiet -out 0 -o " & OutFile & InFile
End If

ExecWait xnCmdLine, SW_HIDE

SmartConvert = OutFile
End Function

