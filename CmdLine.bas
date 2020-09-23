Sub GetCmdLineArgs (CmdLine, Args())
ReDim Args(10) 'more than enough
ai = 1
For i = 1 To Len(CmdLine)
    ch = Mid(CmdLine, i, 1)
    Select Case ch
        Case " ", "-", "/", ","
        If Args(ai) <> "" Then ai = ai + 1
        Case Else
        'do nothing
    End Select
    If ch = " " Or ch = "," Then ch = ""
    Args(ai) = Args(ai) + ch
Next i

ReDim Preserve Args(ai)

'optional
For i = 1 To ai
    Args(i) = LCase(Args(i))
Next i

End Sub

