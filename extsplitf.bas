

Function ExtractDirName (FileName As String) As String

'Extract the Directory name from a full file name
'The return will have "\" appended to the end

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

    ExtractDirName = Left(FileName, PrevPos)

End Function

Function ExtractFileName (FileName As String) As String
    
'Extract the File name from a full file name


    Dim pos As Integer
    Dim PrevPos As Integer

    pos = InStr(FileName, "\")
    If pos = 0 Then
        ExtractFileName = ""
        Exit Function
    End If
    
    Do While pos <> 0
        PrevPos = pos
        pos = InStr(pos + 1, FileName, "\")
    Loop

    ExtractFileName = Right(FileName, Len(FileName) - PrevPos)

End Function

Sub SplitFileName (FName As String, TheName As String, TheExt As String)
Dim cntr As Integer
Dim ch As String * 1
Dim ThePos As Integer
ThePos = 0
'FName may contain more than one "." we want the last one
For cntr = Len(FName) To 1 Step -1
    ch = Mid$(FName, cntr, 1)
    If ch = "." Then
        ThePos = cntr
        Exit For
    End If
Next cntr


If ThePos = 0 Then
    TheName = FName
    TheExt = ""
Else
    TheName = Left$(FName, ThePos - 1)
    TheExt = Right(FName, Len(FName) - ThePos)
End If

End Sub

