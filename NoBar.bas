Option Explicit
Declare Function GetWindowLong Lib "User" (ByVal hWnd As Integer, ByVal nIndex As Integer) As Long
Declare Function SetWindowLong Lib "User" (ByVal hWnd As Integer, ByVal nIndex As Integer, ByVal dwNewLong As Long) As Long
Const GWL_STYLE = (-16)
Const WS_DLGFRAME = &H400000
Const WS_SYSMENU = &H80000
Const WS_MINIMIZEBOX = &H20000
Const WS_MAXIMIZEBOX = &H10000

Sub RemoveTitleBar (frm As Form)

Static OriginalStyle As Long
Dim CurrentStyle As Long
Dim X As Long
    OriginalStyle = 0
    CurrentStyle = GetWindowLong(frm.hWnd, GWL_STYLE)
    OriginalStyle = OriginalStyle Or (CurrentStyle And WS_DLGFRAME)
    OriginalStyle = OriginalStyle Or (CurrentStyle And WS_SYSMENU)
    OriginalStyle = OriginalStyle Or (CurrentStyle And WS_MINIMIZEBOX)
    OriginalStyle = OriginalStyle Or (CurrentStyle And WS_MAXIMIZEBOX)
    CurrentStyle = CurrentStyle And Not WS_DLGFRAME
    CurrentStyle = CurrentStyle And Not WS_SYSMENU
    CurrentStyle = CurrentStyle And Not WS_MINIMIZEBOX
    CurrentStyle = CurrentStyle And Not WS_MAXIMIZEBOX
    X = SetWindowLong(frm.hWnd, GWL_STYLE, CurrentStyle)
    frm.Width = frm.Width  'to refresh (Refresh didn't work)


End Sub

Sub RestoreTitleBar (frm As Form)
'  DIDN'T WORK!!!!!

Static OriginalStyle As Long
Dim CurrentStyle As Long
Dim X As Long
    
    CurrentStyle = GetWindowLong(frm.hWnd, GWL_STYLE)
    CurrentStyle = CurrentStyle Or OriginalStyle
    X = SetWindowLong(frm.hWnd, GWL_STYLE, CurrentStyle)


End Sub

