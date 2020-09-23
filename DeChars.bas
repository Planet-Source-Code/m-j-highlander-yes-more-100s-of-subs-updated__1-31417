Function ReplaceChars (ByVal astr As String, ByVal ReplaceWith As String, ByVal UnwantedChars As String) As String
Dim tmpStr As String
Dim ch As String
Dim i As Integer

tmpStr = ""

For i = 1 To Len(UnwantedChars)
    ch = Mid$(UnwantedChars, i, 1)
    If ch = "!" Then ch = ""
    tmpStr = tmpStr + ch
Next i
UnwantedChars = tmpStr

tmpStr = ""
ch = ""

If Left(UnwantedChars, 1) <> "[" Then UnwantedChars = "[" + UnwantedChars
If Right(UnwantedChars, 1) <> "]" Then UnwantedChars = UnwantedChars + "]"

For i = 1 To Len(astr)
    ch = Mid$(astr, i, 1)
    If ch = "!" Then ch = ReplaceWith   '  "!" has special meaning to LIKE
    If ch Like UnwantedChars Then
        ch = ReplaceWith
        If Right$(tmpStr, 1) = ReplaceWith Then ch = ""
    End If
    
    tmpStr = tmpStr + ch
Next i
ReplaceChars = tmpStr

End Function
