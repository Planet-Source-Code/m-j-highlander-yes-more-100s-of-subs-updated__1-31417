Option Explicit

Sub GetAllEnvirons (EnvArray() As String)
' This Function retrieves all Environment strings
' and stores them into an array.
' Sample Call:
'   ReDim ea(1 To 1) As String
'   GetAllEnvirons ea()

Dim EnvString, Indx

Indx = 1
Do
    EnvString = Environ(Indx)
    If EnvString = "" Then Exit Do
    ReDim Preserve EnvArray(1 To Indx)
    EnvArray(Indx) = EnvString
    Indx = Indx + 1
    
Loop

End Sub

