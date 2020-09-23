Attribute VB_Name = "GeneralSubs"
Option Explicit

Global Const ATTR_READONLY = 1    'Read-only file
Global Const ATTR_VOLUME = 8  'Volume label
Global Const ATTR_ARCHIVE = 32    'File has changed since last back-up
Global Const ATTR_NORMAL = 0  'Normal files
Global Const ATTR_HIDDEN = 2  'Hidden files
Global Const ATTR_SYSTEM = 4  'System files
Global Const ATTR_DIRECTORY = 16  'Directory

Global Const ATTR_DIR_ALL = ATTR_DIRECTORY + ATTR_READONLY + ATTR_ARCHIVE + ATTR_HIDDEN + ATTR_SYSTEM
Global Const ATTR_ALL_FILES = ATTR_NORMAL Or ATTR_HIDDEN Or ATTR_SYSTEM Or ATTR_READONLY Or ATTR_ARCHIVE


Type BMOffset
    Offset As Long
    Size As Long
End Type

Dim AllOffsets() As Long
Dim ValidOffsets()  As BMOffset


Function ParseString(SubStrs() As String, ByVal SrcStr As String, ByVal Delimiter As String) As Integer
'This routine returns an empty string if it finds two consecutive delimiters

'Sample Call:
'SubStrCount = ParseString(SubStr(), DataString, ",")


ReDim SubStrs(0 To 0) As String
Dim CurPos As Long
Dim NextPos As Long
Dim DelLen As Integer
Dim nCount As Integer
Dim TStr As String



    ' Add delimiters to start and end of string to make loop simpler:
    SrcStr = Delimiter & Trim$(SrcStr) & Delimiter
    ' Calculate the delimiter length only once:
    DelLen = Len(Delimiter)
    ' Initialize the count and position:
    nCount = 0
    CurPos = 1
    NextPos = InStr(CurPos + DelLen, SrcStr, Delimiter)

    ' Loop searching for delimiters:
    Do Until NextPos = 0
        ' Extract a sub-string:
        TStr = Mid$(SrcStr, CurPos + DelLen, NextPos - CurPos - DelLen)
        ' Increment the sub string counter:
        nCount = nCount + 1
        ' Add room for the new sub-string in the array:
        ReDim Preserve SubStrs(nCount) As String
        ' Put the sub-string in the array:
        SubStrs(nCount) = TStr
        ' Position to the last found delimiter:
        CurPos = NextPos
        ' Find the next delimiter:
        NextPos = InStr(CurPos + DelLen, SrcStr, Delimiter)
    Loop


    ' Return the number of sub-strings found:
    ParseString = nCount


End Function


Function DirExists(sDir As String) As Integer
Dim tmp As String
Dim iResult As Integer

iResult = 0
If Dir$(sDir, ATTR_DIR_ALL) <> "" Then
    iResult = GetAttr(sDir) And ATTR_DIRECTORY
End If

If iResult = 0 Then   'Directory not found, or the passed argument is a filename not a directory
    DirExists = False
Else
    DirExists = True
End If


End Function


Sub RWChuncks(srcFile$, DestFile$, Start&, Size&)
Dim src%, tgt%, NumChunks&, BytesLeft&, cntr&
Const CHUNKSIZE = 1000

Dim Chunck As String * CHUNKSIZE
Dim ch As String * 1
Dim LastChunck As String

src = FreeFile
Open srcFile$ For Binary As #src
NumChunks& = Int(Size& / CHUNKSIZE)
BytesLeft& = Size& - NumChunks& * CHUNKSIZE

tgt = FreeFile
Open DestFile$ For Binary As #tgt
Seek #src, Start&

For cntr& = 1 To NumChunks&
    Get #src, , Chunck$
    Put #tgt, , Chunck$
    
Next cntr&

On Error Resume Next
LastChunck = Input(BytesLeft&, src)
Put #tgt, , LastChunck

Close src, tgt

End Sub


Sub CenterForm(frm As Form)
    frm.Left = (Screen.Width - frm.Width) / 2
    frm.Top = (Screen.Height - frm.Height) / 2
End Sub
Function ChangeFileExtension(FileName As String, NewExtension As String) As String
Dim OldExt As String
OldExt = ExtractFileExtension(FileName)
ChangeFileExtension = Left$(FileName, Len(FileName) - Len(OldExt)) & NewExtension

End Function

Function ExtractFileExtension(FileName As String) As String

    Dim pos As Integer
    Dim PrevPos As Integer

    pos = InStr(FileName, ".")
    If pos = 0 Then
    ExtractFileExtension = ""
    Exit Function
    End If
    
    Do While pos <> 0
    PrevPos = pos
    pos = InStr(pos + 1, FileName, ".")
    Loop

    ExtractFileExtension = Right(FileName, Len(FileName) - PrevPos)

End Function
Function ExtractDirName(FileName As String) As String

'Extract the Directory name from a full file name
    Dim tmp$
    Dim pos As Integer
    Dim PrevPos As Integer

    pos = InStr(FileName, "\")
    If pos = 0 Then
        ExtractDirName = ""
        Exit Function
    End If
    
    Do While pos <> 0
        PrevPos = pos
        pos = InStr(pos + 1, FileName, "\")
    Loop

    tmp = Left(FileName, PrevPos)
    If Right(tmp, 1) = "\" Then tmp = Left(tmp, Len(tmp) - 1)
    ExtractDirName = tmp
    
End Function
Function ExtractFileName(FileName As String) As String
    
'Extract the File name from a full file name


    Dim pos As Integer
    Dim PrevPos As Integer

    pos = InStr(FileName, "\")
    If pos = 0 Then
    ExtractFileName = ""
    Exit Function
    End If
    
    Do While pos <> 0
    PrevPos = pos
    pos = InStr(pos + 1, FileName, "\")
    Loop

    ExtractFileName = Right(FileName, Len(FileName) - PrevPos)

End Function




Sub TextBox(Txt As String)
frmHelp.txtMain.Text = Txt
frmHelp.Show
End Sub



Public Sub SelectAll(txtControl As TextBox)
txtControl.SelStart = 0
txtControl.SelLength = Len(txtControl.Text)

End Sub
