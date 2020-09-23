Attribute VB_Name = "TextLinesToArray"
Option Explicit

Function Text2Lines(Text As String) As Variant

        Text2Lines = Split(Text, vbCrLf)
        
End Function

