Function GetDirName (ScanString$) As String
    
    ' Sub/Function Name       : GetDirName
    ' Purpose                 : Gets a full directory from a string (with filename)
    ' Parameters              : String to scan for directory name
    ' Return                  : Directory Path
    ' Created by              : Paul Treffers
    ' Date Created            : 19/11/94
    
    Dim ExitWhile As Integer
    Dim Pos%, PosSave%
    ExitWhile = True
    Pos% = 1
    Do While ExitWhile = True
        Pos% = InStr(Pos%, ScanString$, "\")
        If Pos% = 0 Then
            Exit Do
        Else
            Pos% = Pos% + 1
            PosSave% = Pos% - 1
        End If
    Loop
    GetDirName = Left$(ScanString$, PosSave%)
End Function

Function GetFileName (ScanString$) As String
    
    ' Sub/Function Name       : GetFileName
    ' Purpose                 : Gets a filename from string that contains directory also
    ' Parameters              : String to scan for filename
    ' Return                  : Directory Path
    ' Created by              : Paul Treffers
    ' Date Created            : 19/11/94
    
    ExitWhile = True
    Pos% = 1
    Do While ExitWhile = True
        Pos% = InStr(Pos%, ScanString$, "\")
        If Pos% = 0 Then
            Exit Do
        Else
            Pos% = Pos% + 1
            PosSave% = Pos% - 1
        End If
    Loop
    GetFileName = Trim$(Mid$(ScanString$, PosSave% + 1, Len(ScanString$)))

End Function
