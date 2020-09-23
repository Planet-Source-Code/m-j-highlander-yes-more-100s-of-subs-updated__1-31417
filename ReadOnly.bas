Declare Function SendMessage Lib "User" (ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As Any) As Long
Const WM_USER = &H400
Const EM_SETREADONLY = (WM_USER + 31)

Sub MakeTextBoxReadOnly (txt As TextBox)

Dim RetVal As Integer
RetVal = SendMessage(txt.hWnd, EM_SETREADONLY, True, ByVal 0&)

End Sub

Sub UnMakeTextBoxReadOnly (txt As TextBox)

Dim RetVal As Integer
RetVal = SendMessage(txt.hWnd, EM_SETREADONLY, False, ByVal 0&)

End Sub

