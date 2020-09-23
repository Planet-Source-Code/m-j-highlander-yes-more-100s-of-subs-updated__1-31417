Sub SaveFile (FileName As String, FileContent As String)
Dim FileNum As Integer

FileNum = FreeFile

Open FileName For Output As #FileNum

Print #FileNum, FileContent

Close FileNum

End Sub
