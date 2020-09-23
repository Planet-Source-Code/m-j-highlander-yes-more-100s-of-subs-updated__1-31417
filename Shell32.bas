Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long


Public Function FindExec(DocName) As String
' DocName contains the full path of the document
Dim l As Long
Dim buff As String
buff = Space(255)
l = FindExecutable(DocName, "", buff)
FindExec = Trim(buff)
End Function

Public Sub RunDoc(frm As Form, doc As String, WindowState)
' WindowState:  1 normal window, 3 maximized
Dim l As Long
l = ShellExecute(frm.hwnd, "open", doc, ByVal Chr$(0), "", WindowState)

End Sub
