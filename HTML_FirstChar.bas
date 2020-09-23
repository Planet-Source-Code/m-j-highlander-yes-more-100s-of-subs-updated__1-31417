Attribute VB_Name = "Module1"
Option Explicit
Public Function HTML_FirstChar(sText As String, sLeftTag As String, sRightTag As String) As String

' sample call:
' Text = StrConv(Text, vbProperCase)
' Text = HTML_FirstChar(Text, "<Font color=red SIZE=+2>", "</FONT>")


Dim idx As Integer
Dim a As Variant
Dim ch As String

a = Split(sText, " ")

For idx = LBound(a) To UBound(a)
        ch = Left(a(idx), 1)
        a(idx) = Replace(a(idx), ch, sLeftTag & ch & sRightTag, , 1)
Next idx

HTML_FirstChar = Join(a, " ")

End Function
