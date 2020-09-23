Option Explicit
'Global gTgtPath
Global gCancel



Global Const ATTR_READONLY = 1    'Read-only file
Global Const ATTR_VOLUME = 8  'Volume label
Global Const ATTR_ARCHIVE = 32    'File has changed since last back-up
Global Const ATTR_NORMAL = 0  'Normal files
Global Const ATTR_HIDDEN = 2  'Hidden files
Global Const ATTR_SYSTEM = 4  'System files
Global Const ATTR_DIRECTORY = 16  'Directory

Global Const ATTR_DIR_ALL = ATTR_DIRECTORY + ATTR_READONLY + ATTR_ARCHIVE + ATTR_HIDDEN + ATTR_SYSTEM
Global Const ATTR_ALL_FILES = ATTR_NORMAL Or ATTR_HIDDEN Or ATTR_SYSTEM Or ATTR_READONLY Or ATTR_ARCHIVE

Function AddSlash (ztring As String) As String

If Right(ztring, 1) <> "\" Then
    AddSlash = ztring & "\"
Else
    AddSlash = ztring
End If


End Function

Function DirExists (sDir As String) As Integer
Dim tmp As String
Dim iResult As Integer

iResult = 0
If Dir$(sDir, ATTR_DIR_ALL) <> "" Then
    iResult = GetAttr(sDir) And ATTR_DIRECTORY
End If

If iResult = 0 Then   'Directory not found, or the passed argument is a filename not a directory
    DirExists = False
Else
    DirExists = True
End If


End Function

Function ExtractDirName (FileName As String) As String

'Extract the Directory name from a full file name
    Dim tmp$
    Dim pos As Integer
    Dim PrevPos As Integer

    pos = InStr(FileName, "\")
    If pos = 0 Then
        ExtractDirName = ""
        Exit Function
    End If
    
    Do While pos <> 0
        PrevPos = pos
        pos = InStr(pos + 1, FileName, "\")
    Loop

    tmp = Left(FileName, PrevPos)
    If Right(tmp, 1) = "\" Then tmp = Left(tmp, Len(tmp) - 1)
    ExtractDirName = tmp
    
End Function

Function FileExists (sFile As String) As Integer

If Dir$(sFile, ATTR_ALL_FILES) = "" Then
    FileExists = False
Else
    FileExists = True
End If

End Function

Function GetAppPath ()

If Len(app.Path) = 3 Then
    GetAppPath = app.Path
Else
    GetAppPath = app.Path + "\"
End If

End Function

Sub SplitIntoThree (sLine$, sVar1$, sVar2$, sVar3$)
On Error GoTo ErrA
Dim pos1%, pos2%

pos1 = InStr(sLine, "^")
pos2 = InStr(pos1 + 1, sLine, "^")
    
sVar1 = Trim(Left(sLine, pos1% - 1))
If pos2 <> 0 Then
    sVar2 = Trim(Mid(sLine, pos1 + 1, pos2 - pos1 - 1))
    sVar3 = Trim(Right(sLine, Len(sLine) - pos2 - 1))
Else
    sVar2 = Trim(Mid(sLine, pos1 + 1, Len(sLine) - pos1))
    sVar3 = ""
End If
'MsgBox sVar1 + "   " + sVar2 + "   " + sVar3
Exit Sub
ErrA:
'MsgBox "Invalid entries exist in 'DOFRAG.INI', edit the file to correct them", 48, "Warning"
sVar1 = "#Invalid Entry in DOFRAG.INI#"
sVar2 = ""
sVar3 = ""
Exit Sub
End Sub

