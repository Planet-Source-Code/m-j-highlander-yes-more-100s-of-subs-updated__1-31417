Option Explicit
Declare Function FindWindow Lib "User" (ByVal lpClassName As Any, ByVal lpWindowName As Any) As Integer
Declare Sub SetWindowPos Lib "User" (ByVal hWnd As Integer, ByVal hWndInsertAfter As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal wFlags As Integer)

Global Const HWND_TOPMOST = -1
Global Const HWND_NOTOPMOST = -2
Global Const SWP_NOMOVE = &H2
Global Const SWP_NOSIZE = &H1

Sub Main ()
Dim WindowCaption As String
Dim WindowHandle As Integer
Dim iResult As Integer

iResult = Shell("calc.exe", 1)
WindowCaption = "Calculator"
WindowHandle = FindWindow(ByVal 0&, WindowCaption)
If WindowHandle = 0 Then
    'Arabic Interface:
    WindowCaption = "ÇáÍÇÓÈÉ"
    WindowHandle = FindWindow(ByVal 0&, WindowCaption)
End If

SetWindowPos WindowHandle, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE

End Sub

