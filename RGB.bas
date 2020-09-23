Function Hex2Int (ByVal HexNum As String) As Integer
HexNum = UCase$(HexNum)
ch = Right$(HexNum, 1)
Select Case ch
    Case "0" To "9"
        d = Val(ch)
    Case "A"
        d = 10
    Case "B"
        d = 11
    Case "C"
        d = 12
    Case "D"
        d = 13
    Case "E"
        d = 14
    Case "F"
        d = 15
End Select

ch = Left$(HexNum, 1)
Select Case ch
    Case "0" To "9"
        dd = Val(ch)
    Case "A"
        dd = 10
    Case "B"
        dd = 11
    Case "C"
        dd = 12
    Case "D"
        dd = 13
    Case "E"
        dd = 14
    Case "F"
        dd = 15
End Select

Hex2Int = d + 16 * dd
End Function

Sub RGBSplit (RGB_Color As Long, R%, G%, B%)
Dim HexRGB As String

HexRGB = Hex$(RGB_Color)
If Len(HexRGB) < 6 Then HexRGB = String(6 - Len(HexRGB), "0") + HexRGB

HexR = Right(HexRGB, 2)
HexG = Mid(HexRGB, 3, 2)
HexB = Left(HexRGB, 2)
R% = Hex2Int(HexR)
G% = Hex2Int(HexG)
B% = Hex2Int(HexB)

End Sub

