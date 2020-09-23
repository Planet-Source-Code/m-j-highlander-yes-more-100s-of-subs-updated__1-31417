Function UpEachFirst(strz As String) As String
OutStr = UCase(Left(strz, 1))

For i = 1 To Len(strz) - 1
    ch = Mid(str, i, 1)
        If ch = " " Then
             Char = UCase(Mid(strz, i + 1, 1))
        Else
             Char = LCase(Mid(strz, i + 1, 1))
        End If
    
    OutStr = OutStr + Char

Next i
UpEachFirst = OutStr
End Function

Function UpFirst(strz As String) As String
FirstLetter = UCase$(Left(strz, 1))
OtherLetters = LCase(Right(strz, Len(strz) - 1))
UpFirst = FirstLetter + OtherLetters
End Function

