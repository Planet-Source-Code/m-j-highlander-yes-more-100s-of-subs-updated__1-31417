Sub SaveFile (FileName As String, FileContent As String)
Dim FileNum As Integer

FileNum = FreeFile

Open FileName For Output As #FileNum

Print #FileNum, FileContent

Close FileNum

End Sub



Sub SplitIntoFour (sLine$, sVar1$, sVar2$, sVar3$, sVar4$)
' Sample Call:
' SplitIntoFour "ProgName^\Folder1\Foldern^Text^BITMAP", a, b, c, d

Dim pos1%, pos2%, pos3%

    pos1 = InStr(sLine, "^")
    pos2 = InStr(pos1 + 1, sLine, "^")
    pos3 = InStr(pos2 + 1, sLine, "^")
    sVar1 = Trim(Left(sLine, pos1% - 1))
    sVar2 = Trim(Mid(sLine, pos1 + 1, pos2 - pos1 - 1))
    sVar3 = Trim(Mid(sLine, pos2 + 1, pos3 - pos2 - 1))
    sVar4 = Trim(Right(sLine, Len(sLine) - pos3))

End Sub
