option explicit
Function RightTo (sString As String, sChar As String) As String

' Returns a sub-string from the rightmost of the first argument
' until a charachter (the second argumennt) is found.
' The charachter is not included in the returned string.
' If the charachter is not found, the entire string will be returned.


Dim iCntr As Integer
Dim iLastPos As Integer
Dim iCurrPos As Integer

Dim ch As String * 1

ch = Left$(sChar, 1)

iCurrPos = InStr(sString, ch)
iLastPos = iCurrPos

Do
    DoEvents
    iCurrPos = InStr(iCurrPos + 1, sString, ch)
    If iCurrPos = 0 Then
        Exit Do
    Else
        iLastPos = iCurrPos
    End If
    
Loop

RightTo = Right$(sString, Len(sString) - iLastPos)

End Function

Function LeftTo (sString As String, sChar As String) As String

' Returns a sub-string from the leftmost of the first argument
' until a charachter (the second argumennt) is found.
' The charachter is not included in the returned string.
' If the charachter is not found, an empty string will be returned.



Dim iPos As Integer
Dim ch As String * 1

ch = Left$(sChar, 1)
iPos = InStr(sString, ch)
If iPos <> 0 Then
    LeftTo = Left$(sString, iPos - 1)
Else
    LeftTo = ""
End If

End Function

