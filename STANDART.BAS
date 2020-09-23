Declare Function SendMessage Lib "User" (ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As Any) As Long

Declare Function GetTickCount Lib "User" () As Long


Declare Function GetMenu Lib "User" (ByVal hWnd As Integer) As Integer
Declare Function GetMenuItemId Lib "User" (ByVal hMenu As Integer, ByVal nPos As Integer) As Integer
Declare Function GetSubMenu Lib "User" (ByVal hMenu As Integer, ByVal nPos As Integer) As Integer
Declare Function ModifyMenu Lib "User" (ByVal hMenu As Integer, ByVal nPosition As Integer, ByVal wFlags As Integer, ByVal wIDNewItem As Integer, ByVal lpString As Any) As Integer

'Centers form
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

'This function if the filename is valid.
'IT CHECKS ONLY THE NAME, WITHOUT EXTENTION
Function IsValidFileName (File As String) As Integer
   On Error GoTo L1:
      
     'Checking if there is '.' and if there is, end
      If InStr(File, ".") Then GoTo L1:
     'Opening file. If the name is not valid, error occures
      Open File For Binary As #17
     'Closing file
      Close #17
     'If this file wasn't existing before, delete it
      If FileLen(File) = 0 Then Kill File
     'The name is valid
      IsValidFileName = True
     'Exit
      Exit Function

'If an error occured - the file name is not valid
L1:
     IsValidFileName = False '
     Exit Function
     Resume Next

End Function

'This funtion limits the user's text input to N
Sub LimitInput (C As Control, N As Integer)
    X = SendMessage(C.hWnd, &H400 + 21, N, 0)
End Sub

'This function can be used when creating new directories.
'it is better to use it intead of MkDir, because it
'it creates each level of the path separatly,
'not like the MkDir, which can create only one level
Function MakePath (ByVal strDirName As String) As Integer
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

'Holds the program for A seconds
Sub pause (a!)
X# = GetTickCount() / 1000
Do
DoEvents
Loop Until GetTickCount() / 1000 - X# >= a!
End Sub

'This function displays a messagebox that asks the user if
'he really want to quit this application.
'It returns true if answer is 'YES' , otherwise,
'it returns False
Function Quit ()
    X = MsgBox("This will end your " & AppName$ & " Session", MB_YESNO + MB_IQUEST, AppName$)
    If X = IDYES Then Quit = True Else Quit = False
End Function

'Return a random number between A1 and A2
Function Ran (A1 As Variant, A2 As Variant) As Long
    Randomize Timer
    Ran = Int((A2 - A1 + 1) * Rnd + A1)
End Function

'This functions changes the bitmap of specified menu item
'Hwnd - window handle
'PopInd - the popup menu index(1, 2, 3...)
'ItmInd - the item menu index(1, 2, 3...)
'bmp - the bitmap to be set
Sub SetMenuBitmap (Hwnd1 As Integer, PopInd As Integer, ItmInd As Integer, Bmp As Long)
 Dim hMenu As Integer
 Dim hPopup As Integer
 Dim hItem As Integer
 Dim Temp As Integer

     hMenu = GetMenu(Hwnd1)
     hPopup = GetSubMenu(hMenu, PopInd)
     hItem = GetMenuItemId(hPopup, ItmInd)
     Temp = ModifyMenu(hMenu, hItem, &H4, hItem, Bmp)

End Sub

'Return the number in words
Function Snum$ (I)
    Snum$ = Choose(I, "One", "Two", "Three", "For", "Five", "Six", "Seven", "Eight", "Nine", "Ten")
End Function

