Declare Function FindWindow Lib "User" (ByVal lpClassName As Any, ByVal lpWindowName As Any) As Integer
Declare Sub SetWindowPos Lib "User" (ByVal hWnd As Integer, ByVal hWndInsertAfter As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal wFlags As Integer)

Global Const HWND_TOPMOST = -1
Global Const HWND_NOTOPMOST = -2
Global Const SWP_NOMOVE = &H2
Global Const SWP_NOSIZE = &H1

Declare Function FindWindow Lib "User" (ByVal lpClassName As Any, ByVal lpWindowName As Any) As Integer


Sub CenterForm (frm As Form)
On Error Resume Next
    frm.Left = (screen.Width - frm.Width) / 2
    frm.Top = (screen.Height - frm.Height) / 2
End Sub

Sub TahmoaX ()
On Error Resume Next
form1.Command1.FontName = "Tahoma"
form1.Command2.FontName = "Tahoma"
form1.Label1.FontName = "Tahoma"
End Sub

Sub TopMost (frm As Form)
On Error Resume Next
Dim WindowHandle As Integer

WindowHandle = frm.hWnd
SetWindowPos WindowHandle, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE

End Sub

