Option Explicit


Sub CenterForm (x As Form)
  
x.Top = (Screen.Height * .85) / 2 - x.Height / 2
x.Left = Screen.Width / 2 - x.Width / 2


End Sub

Function Hex2Int (ByVal HexNum As String) As Integer
Dim ch As String
Dim d As Integer, dd As Integer

HexNum = UCase$(HexNum)
ch = Right$(HexNum, 1)
Select Case ch
    Case "0" To "9"
        d = Val(ch)
    Case "A"
        d = 10
    Case "B"
        d = 11
    Case "C"
        d = 12
    Case "D"
        d = 13
    Case "E"
        d = 14
    Case "F"
        d = 15
End Select

ch = Left$(HexNum, 1)
Select Case ch
    Case "0" To "9"
        dd = Val(ch)
    Case "A"
        dd = 10
    Case "B"
        dd = 11
    Case "C"
        dd = 12
    Case "D"
        dd = 13
    Case "E"
        dd = 14
    Case "F"
        dd = 15
End Select

Hex2Int = d + 16 * dd
End Function

Sub RGBSplit (RGB_Color As Long, R%, G%, B%)
Dim HexRGB As String

HexRGB = Hex$(RGB_Color)
If Len(HexRGB) < 6 Then HexRGB = String(6 - Len(HexRGB), "0") + HexRGB
Dim HexR, HexG, HexB
HexR = Right(HexRGB, 2)
HexG = Mid(HexRGB, 3, 2)
HexB = Left(HexRGB, 2)
R% = Hex2Int(HexR)
G% = Hex2Int(HexG)
B% = Hex2Int(HexB)

End Sub

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

Sub Text2Lines (Text As String, Lines() As String)
Dim ch As String * 1
Dim Cntr As Long
Dim Index As Integer
Dim MaxIndex As Integer
Dim NewLine As String * 2

NewLine = Chr(13) + Chr(10)

ReDim Lines(1 To 1000)

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

End Sub

Function UpEachFirst (strz As String) As String
Dim OutStr As String
Dim i As Integer
Dim ch As String, Char As String
OutStr = UCase(Left(strz, 1))

For i = 1 To Len(strz) - 1
    ch = Mid(strz, i, 1)
        If ch = " " Or ch = Chr$(10) Then
             Char = UCase(Mid(strz, i + 1, 1))
        Else
             Char = LCase(Mid(strz, i + 1, 1))
        End If
    
    OutStr = OutStr + Char

Next i
UpEachFirst = OutStr
End Function

Function UpFirst (strz As String) As String
Dim FirstLetter$, OtherLetters$
FirstLetter = UCase$(Left(strz, 1))
OtherLetters = LCase(Right(strz, Len(strz) - 1))
UpFirst = FirstLetter + OtherLetters
End Function

Function XTrim (sLine As String) As String
Dim ch As String * 1
'sLine = Trim$(sLine)
If Right(sLine, 1) = Chr$(13) Or Right(sLine, 1) = Chr$(13) Then
    sLine = Left(sLine, Len(sLine) - 1)
End If
XTrim = sLine
End Function

