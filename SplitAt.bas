Attribute VB_Name = "SplitAtSubStr"
Option Explicit

Public Sub SplitAt(ByVal Text As String, ByVal LookFor As String, ByRef LeftSplit As String, ByRef RightSplit As String)
' Splits a string into 2 parts,
' to the left and right of a specified search string
' SAMPLE CALL (using named args):
' SplitAt Text:="this is it!", LookFor:="is", LeftSplit:=le$, RightSplit:=ri$

Dim pos As Long

pos = InStr(sText, sLookFor)
If pos <> 0 Then
        LeftSplit = Left(Text, pos - 1)
        RightSplit = Right(Text, Len(Text) - pos - Len(LookFor) + 1)
Else
        LeftSplit = ""
        RightSplit = ""
End If

End Sub


