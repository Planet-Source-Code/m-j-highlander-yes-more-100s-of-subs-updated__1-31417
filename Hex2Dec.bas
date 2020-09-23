Function DoHex (txt As String) As String
Dim ch As String
DH = ""
txt = Trim(txt)

For i = 1 To Len(txt) Step 2
ch = Mid(txt, i, 2)
If Left(ch, 1) <> " " Then
    dch = Hex2Dec(ch)
    DH = DH + Chr(dch)
End If
Next i
DoHex = DH
End Function


Function Hex2Dec (ch As String)
ch = UCase(ch)
l = Left(ch, 1)
r = Right(ch, 1)
Select Case r
    Case "0" To "9"
    rNum = Val(r)
    Case "A" To "F"
    rNum = Asc(r) - 55
End Select

Select Case l
    Case "0" To "9"
    lNum = Val(l)
    Case "A" To "F"
    lNum = Asc(l) - 55
End Select

Hex2Dec = rNum + 16 * lNum

End Function
