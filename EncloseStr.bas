Attribute VB_Name = "EncloseString"
Option Explicit

Function EncloseStr(ByVal Text As String, ByVal LeftStr As String, ByVal RightStr As String) As String
' SAMPLE CALL (using named args):
' EncloseStr(LeftStr:="<B>", RightStr:="</B>", Text:="this is it!")

        EncloseStr = LeftStr & Text & RightStr

End Function


