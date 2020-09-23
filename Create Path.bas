'---------------------------------------------
' Create the path contained in DestPath$
' First char must be drive letter, followed by
' a ":\" followed by the path, if any.
'---------------------------------------------
Function CreatePath (ByVal DestPath$) As Integer
    Screen.MousePointer = 11

    '---------------------------------------------
    ' Add slash to end of path if not there already
    '---------------------------------------------
    If Right$(DestPath$, 1) <> "\" Then
        DestPath$ = DestPath$ + "\"
    End If
          

    '-----------------------------------
    ' Change to the root dir of the drive
    '-----------------------------------
    On Error Resume Next
    ChDrive DestPath$
    If Err <> 0 Then GoTo errorOut
    ChDir "\"

    '-------------------------------------------------
    ' Attempt to make each directory, then change to it
    '-------------------------------------------------
    BackPos = 3
    forePos = InStr(4, DestPath$, "\")
    Do While forePos <> 0
        temp$ = Mid$(DestPath$, BackPos + 1, forePos - BackPos - 1)

        Err = 0
        MkDir temp$
        If Err <> 0 And Err <> 75 Then GoTo errorOut

        Err = 0
        ChDir temp$
        If Err <> 0 Then GoTo errorOut

        BackPos = forePos
        forePos = InStr(BackPos + 1, DestPath$, "\")
    Loop
                 
    CreatePath = True
    Screen.MousePointer = 0
    Exit Function
                 
errorOut:
    MsgBox "Error While Attempting to Create Directories on Destination Drive.", 48, "SETUP"
    CreatePath = False
    Screen.MousePointer = 0

End Function

