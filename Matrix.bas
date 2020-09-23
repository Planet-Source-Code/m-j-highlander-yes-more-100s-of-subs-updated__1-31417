Sub MatrixProd (a(), b(), c())

c1 = UBound(a, 2)
r2 = UBound(b, 1)
If c1 <> r2 Then Exit Sub
r1 = UBound(a, 1)
c2 = UBound(b, 2)
lb1 = LBound(a, 1)
lb2 = LBound(b, 2)
lb = LBound(a, 2)
ReDim c(lb1 To r1, lb2 To c2)
For i = lb1 To r1
    For j = lb2 To c2
            c(i, j) = 0
            For k = lb To c1   'or    "R2"
            c(i, j) = c(i, j) + a(i, k) * b(k, j)
            Next k
    Next j
Next i

End Sub

Sub MatrixAdd (a(), b(), c())
lb = LBound(a)
ub = UBound(a)
ReDim c(lb To ub, lb To ub)
For i = lb To ub
    For j = lb To ub
        c(i, j) = a(i, j) + b(i, j)
    Next j
Next i
End Sub

Function MatrixDet (a())

lb = LBound(a)
ub = UBound(a)
MDet = 1
For i = lb To ub - 1
If a(i, i) = 0 Then
FGOS = 0
    For L = i + 1 To ub
    If a(L, i) <> 0 Then
            For M = lb To ub
            Q = a(L, M)
            a(L, M) = a(i, M)
            a(i, M) = Q
            Next M
            MDet = -MDet
            FGOS = -1
            Exit For   'L'
    End If
    Next L
    If FGOS <> -1 Then MDet = 0
End If
If MDet = 0 Then Exit Function
For j = i + 1 To ub
t = a(i, j) / a(i, i)
For k = lb To ub
a(k, j) = a(k, j) - a(k, i) * t
Next k
Next j
Next i
For i = lb To ub
MDet = MDet * a(i, i)
Next i
MatrixDet = MDet
End Function

Sub MatrixInv (a(), IA())
s = MatrixDet(a())

If s = 0 Then Exit Sub

lb = LBound(a)
ub = UBound(a)
ReDim IA(lb To ub, lb To ub) As Variant
ReDim sm(lb To ub - 1, lb To ub - 1) As Variant

For v = lb To ub
    For w = lb To ub
    le = 0
            For x = lb To ub - 1
                ce = 0
                For y = lb To ub - 1
                    If v = x Then le = 1
                    If y = w Then ce = 1
                    sm(x, y) = a(x + le, y + ce)
                Next y
            Next x
        IA(w, v) = (-1) ^ (w + v) * MatrixDet(sm()) / s
    Next w
Next v


End Sub

Sub MatrixMul (M(), b)

For i = LBound(M) To UBound(M)
    For j = LBound(M) To UBound(M)
            M(i, j) = M(i, j) * b
    Next j
Next i


End Sub

Sub MatrixSubtratct (a(), b(), c())
lb = LBound(a)
ub = UBound(a)
ReDim c(lb To ub, lb To ub)
For i = lb To ub
    For j = lb To ub
        c(i, j) = a(i, j) - b(i, j)
    Next j
Next i

End Sub

