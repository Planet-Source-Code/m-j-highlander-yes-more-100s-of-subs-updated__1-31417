Function Slasher (Strng As String, flag As String) As String
' Flag could be:
' "\?" to add a slash to the left if it doesn't already exist
' "?\" to add a slash to the right if it doesn't already exist
' "\?\" to enclose the string in slashes
' any other string to strip left and right slashes
' "?" can be any single character.

Dim AString As String

AString = Strng
If flag Like "\?" Then
    'left slash
    If Left(AString, 1) <> "\" Then AString = "\" + AString
ElseIf flag Like "?\" Then
    'right slash
    If Right(AString, 1) <> "\" Then AString = AString + "\"
ElseIf flag Like "\?\" Then
    'right & left slashes
    If Left(AString, 1) <> "\" Then AString = "\" + AString
    If Right(AString, 1) <> "\" Then AString = AString + "\"
Else
    'strip slashes if existing
    If Left(AString, 1) = "\" Then AString = Right(AString, Len(AString) - 1)
    If Right(AString, 1) = "\" Then AString = Left(AString, Len(AString) - 1)
End If

Slasher = AString
End Function

