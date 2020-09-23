Function StripComment (sLine)
Dim Pos As Integer

    Pos = InStr(sLine, ";")
    If Pos <> 0 Then
        StripComment = Trim(Left(sLine, Pos - 1))
    Else
        StripComment = sLine
    End If

End Function
 