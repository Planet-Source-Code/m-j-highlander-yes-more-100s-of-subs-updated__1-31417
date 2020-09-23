Global SearchResult() As String
Global index As Long

Global Total As Long
Global StartingDir As String

Global OriginalDrive As String
Global OriginalDir As String

Const HOURGLASS = 11
Const DEFAULT = 0

Sub AddFileNames ()
'On Error Resume Next
' Add slash in filename unless we are at the root directory
    Path$ = ScanDir.Dir1.Path
    If Right$(Path$, 1) <> "\" Then Path$ = Path$ + "\"

    
    Total = Total + ScanDir.File1.ListCount
    
    
    
    If Total > 0 Then ReDim Preserve SearchResult(1 To Total)

    For j = 1 To ScanDir.File1.ListCount
        fname$ = Path$ + ScanDir.File1.List(j - 1)
        
        SearchResult(index) = fname$
        index = index + 1
    
    Next j
    
    
End Sub

Sub InitiateSearch (Pattern As String, Directory As String)
'Pattern can be one or more extensions and/or one or more filenames.
'example: pattern=*.exe;*.txt;autoexec.bat

ScanDir.File1.Pattern = Pattern
Total = 0
index = 1         'asuming the matrix lower bound is 1
ReDim SearchResult(1 To 1)   'something to start with

Screen.MousePointer = HOURGLASS
ChDrive Left(Directory, 2)
ChDir Directory

StartingDir = CurDir

ScanDir.Dir1.Path = StartingDir
ScanDir.Drive1.Drive = StartingDir

Call Scan
Screen.MousePointer = DEFAULT
'restore original drive & directory...
ChDrive OriginalDrive
ChDir OriginalDir


End Sub

Sub Main ()

'use this code either here or in a form
'in the project sample.mak this sub is not implemented and can be removed
OriginalDrive = Left(CurDir, 2)
OriginalDir = CurDir
Load ScanDir

Call InitiateSearch("*.exe", "d:\")


End Sub

Sub MoveParentDir ()
Rem DO NOT ATTEMPT TO MOVE UP FROM ORIGINAL DIRECTORY

    If UCase$(ScanDir.Dir1.List(-1)) <> StartingDir Then
        ChDir ScanDir.Dir1.List(-2)
        ScanDir.Dir1.Path = ScanDir.Dir1.List(-2)
    End If

End Sub

Sub Scan ()
Dim md As Integer

    md = ScanDir.Dir1.ListCount                 ' dirs to scan ???
    If md > 0 Then                      ' if yes,
        For i = 0 To md - 1             '   for each one of them
            ChDir ScanDir.Dir1.List(i)          '    move to subdir
            ScanDir.Dir1.Path = ScanDir.Dir1.List(i)    '    display its name in Dir List
            ScanDir.File1.Path = ScanDir.Dir1.List(i)   '    display file names in current path
           
            Call Scan                   '   repeat for each subdir
        Next
        
        ScanDir.File1.Path = ScanDir.Dir1.Path          ' display file names in current path
        Call AddFileNames               '  (process files)
        Call MoveParentDir              ' and move to parent dir
    Else                                ' No subdirs to scan
        ScanDir.File1.Path = ScanDir.Dir1.Path
        Call AddFileNames               '  (process files)
        Call MoveParentDir              ' and move to parent dir
    End If

End Sub

