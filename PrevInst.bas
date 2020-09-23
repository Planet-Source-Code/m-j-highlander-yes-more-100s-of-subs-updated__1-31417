Declare Function FindWindow Lib "User" (ByVal lpClassName As Any, ByVal lpWindowName As Any) As Integer
Declare Function ShowWindow Lib "User" (ByVal hWnd As Integer, ByVal nCmdShow As Integer) As Integer
Declare Sub SetWindowPos Lib "User" (ByVal hWnd As Integer, ByVal hWndInsertAfter As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal wFlags As Integer)

Global Const SWP_NOSIZE = &H1
Global Const SWP_NOMOVE = &H2
Global Const SW_SHOWMAXIMIZED = 3
Global Const SW_SHOWNORMAL = 1

Sub CheckPrevInst (frm As Form)

On Error GoTo PrevInstError

Dim h%, s$, sh%
If app.PrevInstance Then
    s$ = frm.Caption
    
    frm.Caption = "foo"
    h% = FindWindow(ByVal 0&, ByVal s$)
    SetWindowPos h%, 0, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
    sh% = ShowWindow(h%, SW_SHOWNORMAL)
    End
End If

Exit Sub

PrevInstError:
MsgBox "The Error '" + UCase$(Error$) + "' occured while checking for a previous instance of the program.", 16, "Generic AutoRun"
Exit Sub


End Sub

