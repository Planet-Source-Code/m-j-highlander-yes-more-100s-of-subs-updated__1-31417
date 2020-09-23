Attribute VB_Name = "Module1"
Option Explicit
Function isFileBitmap(FileName As String) As Boolean
Dim iFF As Integer
Dim ch2 As String * 2

iFF = FreeFile
Open FileName For Binary Access Read As #iFF
Get #iFF, , ch2
Close #iFF
Select Case ch2
    Case "BM"
        isFileBitmap = True
    Case Else
         isFileBitmap = False
End Select

End Function
