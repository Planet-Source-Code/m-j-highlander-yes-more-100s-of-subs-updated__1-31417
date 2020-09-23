Declare Function SendMessage Lib "User" (ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As Any) As Long
Declare Function GetTickCount Lib "User" () As Long

Sub CenterForm2 (X As Form)
    X.Top = (Screen.Height * .85) / 2 - X.Height / 2
    X.Left = Screen.Width / 2 - X.Width / 2
End Sub

Sub CheckIfRunning ()
    If App.PrevInstance = True Then
    X = MsgBox("You already have one copy of this program running in your system. If you want to run a new copy of it, firstable quit existing copy.", MB_OK + MB_ISTOP, AppName$)
    End
    End If
End Sub

Function CurDrive$ ()
    CurDrive$ = Left$(CurDir$, 1)
End Function

Function GetLinesCount (C As Control) As Integer
    h = C.hWnd
    GetLinesCount = SendMessage(h, &H40A, 0, 0)
End Function

Sub LimitInput (C As Control, N As Integer)
    X = SendMessage(C.hWnd, &H400 + 21, N, 0)
End Sub

Function MakePath (ByVal strDirName As String) As Integer
'This function can be used when creating new directories.
'it is better to use it intead of MkDir, because it
'it creates each level of the path separatly,
'not like the MkDir, which can create only one level

    Dim strPath As String
    Dim intOffset As Integer
    Dim intAnchor As Integer
    
    On Error Resume Next

    '
    'Remove any trailing backslash
    '
    If Right$(strDirName, 1) = "\" Then
        strDirName = Left$(strDirName, Len(strDirName) - 1)
    End If

    intAnchor = 0

    '
    'Loop and make each subdir of the path separately.  After the loop,
    'MkDir again because strDirName doesn't end with a dir separator
    'char.  At the end, try to change into the dir we just create to
    'determine whether the creation was successful.
    '
    Do
        intOffset = InStr(intAnchor + 1, strDirName, "\")
        intAnchor = intOffset

        If intAnchor > 0 Then
            strPath = Left$(strDirName, intOffset - 1)
            MkDir strPath
          Else
            Exit Do
        End If
    Loop Until intAnchor = 0

    MkDir strDirName

    strPath = CurDir$
    Err = 0
    ChDir strDirName
    MakePath = IIf(Err, False, True)
    ChDir strPath
    
    Err = 0
End Function

Sub Pause (a!)
'Holds the program for A seconds

X# = GetTickCount() / 1000
Do
    DoEvents
Loop Until GetTickCount() / 1000 - X# >= a!

End Sub

Function Ran (A1 As Variant, A2 As Variant) As Long
'Return a random number between A1 and A2

    Randomize Timer
    Ran = Int((A2 - A1 + 1) * Rnd + A1)
End Function

