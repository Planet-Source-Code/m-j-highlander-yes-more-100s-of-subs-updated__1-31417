Global gSortArray()

Sub BubbleSort (Array())
Dim Swtch As Integer
Dim Limit As Integer
Dim Row As Integer

Limit = UBound(Array)'MaxRow
   Do
      Swtch = False
      For Row = LBound(Array) To (Limit - 1)
        ' Two adjacent elements are out of order, so swap their values
        If Array(Row) > Array(Row + 1) Then
            Temp = Array(Row)
            Array(Row) = Array(Row + 1)
            Array(Row + 1) = Temp
            Swtch = Row
         End If
      Next Row

      ' Sort on next pass only to where the last switch was made:
      Limit = Swtch
   Loop While Swtch

End Sub

Sub ExchangeSort (SortArray())
Dim SmallestRow As Integer
Dim Row As Integer

MaxRow = UBound(SortArray)
MinRow = LBound(SortArray)

For Row = MinRow To MaxRow
    SmallestRow = Row
    For J = Row + 1 To MaxRow
        If SortArray(J) < SortArray(SmallestRow) Then
        SmallestRow = J
        End If
    Next J

    If SmallestRow > Row Then
        Swap SortArray(Row), SortArray(SmallestRow)
    
    End If
Next Row
End Sub

Sub InsertionSort (SortArray())
Dim TempVal
Dim Row As Integer
Dim J As Integer


MaxRow = UBound(SortArray)
MinRow = LBound(SortArray)

For Row = MinRow + 1 To MaxRow
    TempVal = SortArray(Row)
    Temp = TempVal
    For J = Row To MinRow + 1 Step -1

        ' As long as the length of the J-1st element is greater than the
        ' length of the original element in SortArray(Row), keep shifting
        ' the array elements down:
        If SortArray(J - 1) > Temp Then
        SortArray(J) = SortArray(J - 1)

        ' Otherwise, exit the FOR...NEXT loop:
        Else
        Exit For
        End If
    Next J

    ' Insert the original value of SortArray(Row) in SortArray(J):
    SortArray(J) = TempVal
Next Row
End Sub

Sub QuickSort (Low, High)
Dim I As Integer
Dim J As Integer


' SortArray() is declared at module level

   If Low < High Then

      ' Only two elements in this subdivision; swap them if they are out of
      ' order, then end recursive calls:
      If High - Low = 1 Then
         If gSortArray(Low) > gSortArray(High) Then
            Swap gSortArray(Low), gSortArray(High)
            
         End If
      Else

         ' Pick a pivot element at random, then move it to the end:
         RandIndex = RndInt%(Low, High)
         Swap gSortArray(High), gSortArray(RandIndex)
         Partitionx = gSortArray(High)
         Do

            ' Move in from both sides towards the pivot element:
            I = Low: J = High
            Do While (I < J) And (gSortArray(I) <= Partitionx)
               I = I + 1
            Loop
            Do While (J > I) And (gSortArray(J) >= Partitionx)
               J = J - 1
            Loop

            ' If we haven't reached the pivot element, it means that two
            ' elements on either side are out of order, so swap them:
            If I < J Then
               Swap gSortArray(I), gSortArray(J)
               
            End If
         Loop While I < J

         ' Move the pivot element back to its proper place in the array:
         Swap gSortArray(I), gSortArray(High)
         

         ' Recursively call the QuickSort procedure (pass the smaller
         ' subdivision first to use less stack space):
         If (I - Low) < (High - I) Then
            QuickSort Low, I - 1
            QuickSort I + 1, High
         Else
            QuickSort I + 1, High
            QuickSort Low, I - 1
         End If
      End If
   End If
End Sub

Function RndInt (Lower, Upper) As Integer
'Returns a random integer greater than or equal to the Lower parameter
'and less than or equal to the Upper parameter.
Randomize Timer
RndInt = Int(Rnd * (Upper - Lower + 1)) + Lower

End Function

Sub SelectionSort (vArray())

Dim lLoop1 As Long
Dim lLoop2 As Long
Dim lMin As Long
Dim lTemp As Long
    
    For lLoop1 = LBound(vArray) To UBound(vArray) - 1
    lMin = lLoop1
        For lLoop2 = lLoop1 + 1 To UBound(vArray)
        If vArray(lLoop2) < vArray(lMin) Then lMin = lLoop2
        Next lLoop2
        lTemp = vArray(lMin)
        vArray(lMin) = vArray(lLoop1)
        vArray(lLoop1) = lTemp
    Next lLoop1

End Sub

Sub ShellSort (SortArray())
'The fastets sort algorithm!

Dim Row As Integer
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
            If SortArray(Row) > SortArray(Row + Offset) Then
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

Sub Swap (Var1, Var2)
Dim tmp As Variant
    tmp = Var1
    Var1 = Var2
    Var2 = tmp

End Sub

Sub TrivialSort (SortArray())
'TrivialSort is painfuly slow....
'don't even dream of testing it with the 2000 element array

Cntr = LBound(SortArray)
Max = UBound(SortArray)

Do Until Cntr >= Max
    If SortArray(Cntr) > SortArray(Cntr + 1) Then
        Swap SortArray(Cntr), SortArray(Cntr + 1)
        Cntr = LBound(SortArray) - 1
    End If
    Cntr = Cntr + 1
Loop

End Sub

