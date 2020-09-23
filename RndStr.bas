Option Explicit

Function RndInt (Lower, Upper) As Integer
'Returns a random integer greater than or equal to the Lower parameter
'and less than or equal to the Upper parameter.
Randomize Timer
RndInt = Int(Rnd * (Upper - Lower + 1)) + Lower

End Function

Function RndStr (StrLen As Integer) As String
' This function generates random strings
' The length is sprcified by the only parameter.
' Frankly, I can't think of any use for this function ;-)

Dim idx As Integer
Dim ch As String * 1
Dim tmp As String

For idx = 1 To StrLen
	ch = Chr$(RndInt(32, 126))
	tmp = tmp + ch
Next idx

RndStr = tmp

End Function

