Declare Function ShellExecute Lib "shell.dll" (ByVal hWnd As Integer, ByVal szOp As String, ByVal szFile As String, ByVal szParams As String, ByVal szDir As String, ByVal cmd As Integer) As Integer

Sub ShellEx (FullName As String, frm As Form)
	a% = ShellExecute(frm.hWnd, "open", FullName, "", "", 5)
End Sub

