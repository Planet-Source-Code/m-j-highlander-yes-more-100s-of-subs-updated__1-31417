Option Explicit

Sub DeSpace (sSourceFile As String, sDestFile As String)
Dim sLine As String, sWLine As String
Dim iSrc As Integer, iDst As Integer
Dim ch As String * 1
Dim idx As Integer
Dim bFlag As Integer 'Boolean
Dim bAster As Integer 'Boolean

iSrc = FreeFile
Open sSourceFile For Input Access Read As #iSrc
iDst = FreeFile
Open sDestFile For Output Access Write As #iDst

Do While Not EOF(iSrc)
    sWLine = ""
    Line Input #iSrc, sLine
    bAster = 0
    For idx = 1 To Len(sLine) '- 1
        ch = Mid$(sLine, idx, 1)
        If ch = " " And Mid$(sLine, idx + 1, 1) = " " Then
            bAster = bAster + 1
            bFlag = False
        Else
            bFlag = True
        End If
        If bAster = 1 Then sWLine = sWLine & "*"
        If bFlag Then
            'If Right$(sWLine, 1) = "*" And ch = " " Then ch = ""
            sWLine = sWLine & ch
        End If
    Next idx
    Print #iDst, sWLine
Loop
Close

End Sub

