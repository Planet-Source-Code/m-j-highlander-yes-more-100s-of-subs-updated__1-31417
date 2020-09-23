Option Explicit
Declare Function GetWindowLong Lib "User" (ByVal hWnd As Integer, ByVal nIndex As Integer) As Long
Declare Function SetWindowLong Lib "User" (ByVal hWnd As Integer, ByVal nIndex As Integer, ByVal dwNewLong As Long) As Long
Declare Sub SetWindowPos Lib "User" (ByVal hWnd As Integer, ByVal hWndInsertAfter As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal CX As Integer, ByVal CY As Integer, ByVal wFlags As Integer)

Const SWP_NOSIZE = &H1
Const SWP_NOZORDER = &H4
Const SWP_NOMOVE = &H2
Const SWP_DRAWFRAME = &H20
Const GWL_STYLE = (-16)
Const WS_THICKFRAME = &H40000

Sub ResizeControl (ControlName As Control, FormName As Form)
' Works with the following controls:
' TextBox, Picture, List, Command Button...
Dim NewStyle As Long
    
    NewStyle = GetWindowLong(ControlName.hWnd, GWL_STYLE)
    NewStyle = NewStyle Or WS_THICKFRAME
    NewStyle = SetWindowLong(ControlName.hWnd, GWL_STYLE, NewStyle)
    SetWindowPos ControlName.hWnd, FormName.hWnd, 0, 0, 0, 0, SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_DRAWFRAME

End Sub

