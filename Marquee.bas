Sub ShowMessage (MyText As String, MyPic As PictureBox)
' use this sub in a timer event
Static MsgPtr As Integer

    If MsgPtr = 0 Then MsgPtr = 1
    If Len(MyText) = 0 Then
        MsgPtr = 1
    End If
    MyPic.Cls
    MyPic.Print Mid$(MyText, MsgPtr); MyText;
    MsgPtr = MsgPtr + 1
    
    If MsgPtr > Len(MyText) Then
        MsgPtr = 1
    End If
End Sub
