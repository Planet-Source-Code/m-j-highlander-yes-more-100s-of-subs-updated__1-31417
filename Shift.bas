Option Explicit

Function ShiftChars (Text As String, iShift As Integer, CrLfTabSkip As Integer) As String
' Synatx:
' Text:  Source text
' iShift: number to add to each char's ascii value
' CrLfTabSkip: if true then CR,LF and TAB will be skipped

Dim cntr As Integer
Dim ch As String * 1
Dim TmpStr As String

Const vbLf = 10
Const vbCr = 13
Const vbTab = 9

If Text = "" Then
    ShiftChars = ""
    Exit Function
End If

TmpStr = ""

For cntr = 1 To Len(Text)
    ch = Mid(Text, cntr, 1)
    'If ((Asc(ch) + iShift) > 255 Or (Asc(ch) + iShift < 1)) Then
    '   ShiftChars = ""
    '   Exit Function
    'End If


    If CrLfTabSkip = True Then
        Select Case Asc(ch)
            Case vbLf, vbCr, vbTab
            'do nothing
            Case Else
            ch = Chr$(Asc(ch) + iShift)
        End Select
    Else
    ch = Chr$(Asc(ch) + iShift)
    End If

    TmpStr = TmpStr + ch
Next cntr

ShiftChars = TmpStr

End Function

