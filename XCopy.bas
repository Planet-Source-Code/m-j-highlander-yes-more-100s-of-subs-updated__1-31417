Option Explicit

Global Const ATTR_DIRECTORY = 16

Function XCopy (srcPath As String, dstPath As String, IncludeSubDirs As Integer, FilePat As String) As Integer

' This routine copies all files matching FilePat from scrPath to dstPath.
' If IncludeSubDirs is set to True, all files in subdirs will be incuded (and
' the subdirs themselves of course), like XCOPY /S
' NOTE: dstPath should be created using MkDir in the calling procedure, before calling this function!
'Sample Call:
'  MkDir "d:\cocoz"
'  r = XCopy("d:\q", "d:\cocoz", True, "*.*")

Dim DirOK As Integer, i As Integer
Dim DirReturn As String
ReDim d(100) As String
Dim dCount As Integer
Dim CurrFile$
Dim CurrDir$
Dim dstPathBackup As String
Dim f%

   On Error GoTo DirErr

   CurrDir$ = CurDir$
   
   ' If Path lacks a "\", add one to the end
   If Right$(srcPath, 1) <> "\" Then srcPath = srcPath & "\"
   srcPath = UCase$(srcPath)
   If Right$(dstPath, 1) <> "\" Then dstPath = dstPath & "\"
   dstPath = UCase$(dstPath)

   dstPathBackup = dstPath
   
   ' Initialize var to hold filenames
   DirReturn = Dir(srcPath & "*.*", ATTR_DIRECTORY)
   
   ' Find all subdirs
   Do While DirReturn <> ""
      ' Make sure we don't do anything with "." and "..", they aren't real files
      If DirReturn <> "." And DirReturn <> ".." Then
         
         If (GetAttr(srcPath & DirReturn) And ATTR_DIRECTORY) = ATTR_DIRECTORY Then
            
            ' It's a dir. Add it to dirlist
            dCount = dCount + 1
            d(dCount) = srcPath & DirReturn

         End If
      End If
      DirReturn = Dir
   Loop
   
   ' Now do all the files matching FilePath (and make sure we don't do the dirs)
   DirReturn = Dir(srcPath & FilePat, 0)

   ' Find all files
   Do While DirReturn <> ""
      ' Make sure we don't get a dir
      If Not ((GetAttr(srcPath & DirReturn) And ATTR_DIRECTORY) = ATTR_DIRECTORY) Then
         ' It's a file. Copy it
''''''''' Frm_Copy.Lbl_CopyInfo.Caption = "Copying " & srcPath & DirReturn & " to " & dstPath & DirReturn
''''''''' Frm_Copy.Lbl_CopyInfo.Refresh
         ' Make sure the file doesn't already exist. If it exists, prompt the user
         ' to overwrite it.
         On Error Resume Next
         f% = FreeFile
         Open dstPath & DirReturn For Input As #f%
         Close #f%
         If Err = 0 Then
            ' Prompt the user
            f% = MsgBox("The file " & dstPath & DirReturn & " already exists. Do you wish to overwrite it?", 4 + 32 + 256)
            If f% = 6 Then FileCopy srcPath & DirReturn, dstPath & DirReturn
         Else
            FileCopy srcPath & DirReturn, dstPath & DirReturn
         End If
      End If
      DirReturn = Dir
   Loop

   ' Now do all subs
   For i = 1 To dCount
      
      ' Check the 'IncludeSubDirs' value. If it's true, we have to make
      ' a dir called 'd(i)' in dstPath, and then assign dstPath & d(i) as
      ' dstPath
      If IncludeSubDirs Then

         On Error GoTo PathErr
         
         dstPath = dstPath & Right$(d(i), Len(d(i)) - Len(srcPath))
         
         ' If the Path exists, then this will work out, if not, an error
         ' will be generated and trapped, and the dir will be made
         ChDir dstPath

         On Error GoTo DirErr

      Else

         ' Since we aren't recoursing, we're done
         XCopy = True
         GoTo ExitFunc
         
      End If

      DirOK = XCopy(d(i), dstPath, IncludeSubDirs, FilePat)

      ' Reset dstPath to the value assigned at the argument-line
      dstPath = dstPathBackup

   Next

   XCopy = True

ExitFunc:

   ChDir CurrDir$

   Exit Function

DirErr:

'   Frm_Copy!Lbl_CopyInfo = "Error: " & Error$(Err)
   
   XCopy = False
   Resume ExitFunc

PathErr:
   ' Didn't find the Dir'ed path
   If Err = 75 Or Err = 76 Then
''''''''' Frm_Copy.Lbl_CopyInfo.Caption = "Making directory " & dstPath
''''''''' Frm_Copy.Lbl_CopyInfo.Refresh
      MkDir dstPath
      Resume Next
   End If

   GoTo DirErr
   
End Function

