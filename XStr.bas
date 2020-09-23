Attribute VB_Name = "Replace_Special_Chars"
Option Explicit

Function XStr(Text As String) As String

Dim sTemp As String
Dim QOUT As String

QOUT = Chr$(34)
sTemp = Replace(Text, "\n", vbCrLf)
sTemp = Replace(sTemp, "\q", QOUT)
sTemp = Replace(sTemp, "[nl]", vbCrLf)
sTemp = Replace(sTemp, "[q]", QOUT)
XStr = sTemp

'---------------------------------------------
'SAMPLE CALL:
' Print XStr("this is [q]a[q] test.[nl]This is a new line.")
' //OR
' Print XStr("this is \qa\q test.\nThis is a new line.")
End Function


