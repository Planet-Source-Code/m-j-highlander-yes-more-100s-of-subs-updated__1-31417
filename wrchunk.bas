Attribute VB_Name = "WRCHUNK"
Option Explicit

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

