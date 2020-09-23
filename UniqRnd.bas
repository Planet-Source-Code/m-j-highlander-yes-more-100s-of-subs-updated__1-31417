Global RndArray() As Integer

Sub BubbleSort (Array())
Dim Swtch As Integer
Dim Limit As Integer
Dim Row As Integer

Limit = UBound(Array)
   Do
      Swtch = False
      For Row = LBound(Array) To (Limit - 1)
        ' Two adjacent elements are out of order, so swap their values
        If Array(Row) > Array(Row + 1) Then
            temp = Array(Row)
            Array(Row) = Array(Row + 1)
            Array(Row + 1) = temp
            Swtch = Row
         End If
      Next Row

      ' Sort on next pass only to where the last switch was made:
      Limit = Swtch
   Loop While Swtch

End Sub

Function GenerateUniqueRnds (NumRnds As Integer, Lower%, Upper%) As Integer
Dim Cntr As Integer
Dim Unique As Integer
Dim tmp As Integer

If (Upper% - Lower% + 1) < NumRnds Or (Lower% = 0) Then
    GenerateUniqueRnds = False
    Exit Function
End If

Cntr = 1
ReDim RndArray(1 To NumRnds)

Do Until Cntr > NumRnds
    tmp = RndInt(Lower, Upper)
    Unique = True
    For i = 1 To Cntr
        If RndArray(i) = tmp Then
            'value already exists
            Unique = False
            Exit For
        End If
    Next
    If Unique = True Then
        RndArray(Cntr) = tmp
        Cntr = Cntr + 1
    End If
Loop
GenerateUniqueRnds = True
End Function

Function RndInt (Lower, Upper) As Integer
'Returns a random integer greater than or equal to the Lower parameter
'and less than or equal to the Upper parameter.
Randomize Timer
RndInt = Int(Rnd * (Upper - Lower + 1)) + Lower

End Function

