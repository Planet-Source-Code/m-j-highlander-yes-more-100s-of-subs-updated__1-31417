Function GetAppPath ()
If Len(app.Path) = 3 Then
    GetAppPath = app.Path
Else
    GetAppPath = app.Path + "\"
End If

End Function

