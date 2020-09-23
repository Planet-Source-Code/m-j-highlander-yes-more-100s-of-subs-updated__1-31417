Option Explicit

Sub ShellSort (SortArray() As String)
'The fastets sort algorithm!

Dim Row As Integer
Dim MaxRow As Integer
Dim MinRow As Integer
Dim Swtch As Integer
Dim Limit As Integer
Dim Offset As Integer

MaxRow = UBound(SortArray)
MinRow = LBound(SortArray)
Offset = MaxRow \ 2

Do While Offset > 0
      Limit = MaxRow - Offset
      Do
         Swtch = False         ' Assume no switches at this offset.

         ' Compare elements and switch ones out of order:
         For Row = MinRow To Limit
            If LCase(SortArray(Row)) > LCase(SortArray(Row + Offset)) Then
               Swap SortArray(Row), SortArray(Row + Offset)
               Swtch = Row
            End If
         Next Row

         ' Sort on next pass only to where last switch was made:
         Limit = Swtch - Offset
      Loop While Swtch

      ' No switches at last offset, try one half as big:
      Offset = Offset \ 2
   Loop
End Sub

Sub SortLines (Text As String)
Dim ch As String * 1
Dim Cntr As Long
Dim Index As Integer
Dim MaxIndex As Integer
Dim NewLine As String * 2

'Text = XTrim(Text)
NewLine = Chr(13) + Chr(10)

ReDim Lines(1 To 1000) As String

Index = 1
For Cntr = 1 To Len(Text)
    ch = Mid$(Text, Cntr, 1)
    Select Case Asc(ch)
        Case 13
            'do nothing
        Case 10     'always after the 13
            Index = Index + 1
        Case Else

            Lines(Index) = Lines(Index) + ch
    End Select
Next Cntr

MaxIndex = Index

ReDim Preserve Lines(1 To MaxIndex)

'For Cntr = 1 To MaxIndex
'    Lines(Cntr) = XTrim(Lines(Cntr))
'Next Cntr


ShellSort Lines()
Text = ""
For Cntr = 1 To MaxIndex
    Lines(Cntr) = XTrim(Lines(Cntr))
    If Lines(Cntr) <> "" Then
        Text = Text + Lines(Cntr) + NewLine
    End If

Next Cntr

End Sub

Sub Swap (Var1, Var2)
Dim tmp As Variant
    tmp = Var1
    Var1 = Var2
    Var2 = tmp

End Sub

Function XTrim (sLine As String) As String
Dim ch As String * 1
sLine = Trim$(sLine)
If Right(sLine, 1) = Chr$(13) Or Right(sLine, 1) = Chr$(13) Then
    sLine = Left(sLine, Len(sLine) - 1)
End If
XTrim = sLine
End Function

