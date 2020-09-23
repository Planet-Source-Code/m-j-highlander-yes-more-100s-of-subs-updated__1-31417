Attribute VB_Name = "FUNCTIONS_FOR_KILL"
Option Explicit

Public Const vbQout = """"    ' cannot use chr$() in CONST !!

Sub RegisterShellEx(dll_name As String)
Dim sVal As String
Dim ap As String


ap = App.Path
If Right(ap, 1) <> "\" Then ap = ap & "\"

    'Register DLL
    Shell "RegSvr32.exe" & " " & q & ap & dll_name & q & " /s"

End Sub

Function UpFirst(str As String) As String


        Dim FirstLetter As String, OtherLetters As String
        FirstLetter = UCase$(Left(str, 1))
        OtherLetters = LCase(Right(str, Len(str) - 1))
        UpFirst = FirstLetter + OtherLetters

End Function

StrConv
Public Function XRename(sSrcFile As String, sTgtFile As String) As String
Dim idx As Integer
Dim sNewTgt As String, sExt As String

sNewTgt = sTgtFile

If sSrcFile = sTgtFile Then
            Exit Function
End If

On Error Resume Next

Name sSrcFile As sTgtFile

If Err = 58 Then 'File Already Exists
    
        idx = 0
        Do
                Err = 0
                idx = idx + 1
                sExt = ExtractFileExt(sTgtFile)
                sNewTgt = Left(sTgtFile, Len(sTgtFile) - Len(sExt) - 1) + "-" + Format$(idx) + "." + sExt
                Name sSrcFile As sNewTgt
            
        Loop While Err = 58
    
End If
    
    
If Err <> 0 Then      '--------------Other errors occured
            sErrors = sErrors & "Error   :" & Error & "      FileName:  " & sSrcFile & vbCrLf
            Err = 0
            Exit Function       ' to skip adding to UndoStr and vbsUndoStr
End If

'Return the real new name
XRename = sNewTgt

End Function
Function ReplaceChars(ByVal astr As String, ByVal ReplaceWith As String, ByVal UnwantedChars As String) As String
Dim tmpStr As String
Dim ch As String
Dim i As Integer
Dim bReplaceExclamation As Boolean
Dim bReplaceLeftBracket As Boolean
Dim bReplaceRightBracket As Boolean

tmpStr = ""

For i = 1 To Len(UnwantedChars)
    ch = Mid$(UnwantedChars, i, 1)
    '  "!" and "[" and "]" have special meaning to LIKE, they will be handeled manullay
    If ch = "!" Then ch = "": bReplaceExclamation = True
    If ch = "[" Then ch = "": bReplaceLeftBracket = True
    If ch = "]" Then ch = "": bReplaceRightBracket = True
    tmpStr = tmpStr + ch
Next i
UnwantedChars = tmpStr

tmpStr = ""
ch = ""

'If Left(UnwantedChars, 1) <> "[" Then UnwantedChars = "[" + UnwantedChars
'If Right(UnwantedChars, 1) <> "]" Then UnwantedChars = UnwantedChars + "]"

UnwantedChars = "[" & UnwantedChars & "]"

For i = 1 To Len(astr)
    ch = Mid$(astr, i, 1)
    '  "!" and "[" and "]" have special meaning to LIKE
    If (ch = "!" And bReplaceExclamation) Then ch = ReplaceWith
    If (ch = "[" And bReplaceLeftBracket) Then ch = ReplaceWith
    If (ch = "]" And bReplaceRightBracket) Then ch = ReplaceWith
    If ch Like UnwantedChars Then
        ch = ReplaceWith
        If Right$(tmpStr, 1) = ReplaceWith Then ch = ""
    End If
    
    tmpStr = tmpStr + ch
Next i
ReplaceChars = tmpStr

End Function



Function SaveFile(FileName As String, FileContent As String) As Boolean
On Error GoTo Save_Error
Dim FileNum As Integer

FileNum = FreeFile

Open FileName For Output As #FileNum

Print #FileNum, FileContent

Close FileNum
SaveFile = True
Exit Function

Save_Error:
SaveFile = False
Exit Function
End Function


Function LoadFile(FileName As String) As String
'Loads the contents of a file into a string variable

On Error GoTo LoadFile_Error
Dim ff As Integer
Dim FileContents As String

ff = FreeFile
Open FileName For Input As #ff
FileContents = Input(LOF(ff), ff)
Close #ff
LoadFile = FileContents
Exit Function
LoadFile_Error:
    LoadFile = "#ERROR#"
    Exit Function

End Function


