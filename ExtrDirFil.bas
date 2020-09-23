
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

