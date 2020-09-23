Option Explicit

Sub GetFiles (sSrcFile As String, sDestDir As String, sDstFileBaseName As String, sDstFileExt As String, LookFor As String, LookForOffset As Integer, DstFileSize As Long)
'SAMPLE CALL:
'GetFiles "c:\x\x.bin", "c:\x", "coco", ".jpg", "JFIF", 50, 100000

Dim SrcFileNum As Integer
Dim SrcFileLen As Long
Dim sDstFile As String
Dim indx As Integer
Dim CurrentPos As Long
Dim InPos As Integer
Dim Chunck As String * 1024
Dim ChunckSize As Integer
ChunckSize = Len(Chunck)

SrcFileNum = OpenBinaryFile(sSrcFile)
SrcFileLen = LOF(SrcFileNum)

Do ' exit loop inside
    ReadChunck Chunck, SrcFileNum
    CurrentPos = Seek(SrcFileNum)
    InPos = InStr(Chunck, LookFor)
    If InPos <> 0 Then
        indx = indx + 1
        'save file with a specified length, but restore SEEK...//TODO
        sDstFile = sDestDir & "\" & sDstFileBaseName & Format(indx, "0000") & sDstFileExt
        '***************
        Seek #SrcFileNum, CurrentPos + InPos - LookForOffset
        SaveChunck SrcFileNum, sDstFile, DstFileSize
        Seek #SrcFileNum, CurrentPos + InPos  ' +1  ???
    Else
        'chunck is clean so SEEK
        Seek #SrcFileNum, CurrentPos + ChunckSize
    End If
Debug.Print CurrentPos
If Seek(SrcFileNum) > LOF(SrcFileNum) Then Exit Do
Loop

End Sub

Function OpenBinaryFile (sFileName As String) As Integer
'Opens a file in binary mode and returns the file number

Dim fn As Integer
fn = FreeFile
Open sFileName For Binary As #fn
OpenBinaryFile = fn

End Function

Sub ReadChunck (sChunckVar As String, iFileNum As Integer)
' Gets a chunck of bytes from a binary file
' BUT, does not move the position pointer in the file
' Returns
Dim lOldSeekPos As Long

lOldSeekPos = Seek(iFileNum)
Get #iFileNum, , sChunckVar
Seek #iFileNum, lOldSeekPos


End Sub

Sub SaveChunck (iFileNum As Integer, sDestFile As String, lSize As Long)
Dim lOldSeekPos As Long
Dim Chunk As String * 3000
Dim NumLoops As Integer
Dim cntr As Integer
Dim iTgtFileNum

lOldSeekPos = Seek(iFileNum)
iTgtFileNum = OpenBinaryFile(sDestFile)
NumLoops = lSize / Len(Chunk) + 1
For cntr = 1 To NumLoops
    Get #iFileNum, , Chunk
    Put #iTgtFileNum, , Chunk
Next cntr
Seek #iFileNum, lOldSeekPos

End Sub

