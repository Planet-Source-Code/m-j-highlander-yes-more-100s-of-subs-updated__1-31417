Attribute VB_Name = "Module1"
Option Explicit



Sub main()
'USAGE SAMPLE
Dim src_off&, src_size&, tgt_off&, src$, tgt$, b As Boolean
src_off& = 0
src_size = 70000 'use 0 for ALL
tgt_off& = 367326
src$ = "E:\Bitmap_010.bmp"
tgt$ = "E:\bEcard.exe"
'change the values to something real then uncomment the following line:
'b = ReadWriteChunks(src$, tgt$, src_off&, src_size&, tgt_off&)
MsgBox Str(b)
End Sub


Function ReadWriteChunks(sSourceFile As String, sTargetFile As String, lSourceStart As Long, ByVal lSourceSize As Long, lTargetStart As Long) As Boolean
On Error GoTo ReadWriteChunksError
'********* VARIABLE DECLARATION
Const CHUNKSIZE = 1000  ' maybe this should be a Parameter,
                        ' and use Input() for reading ?
'Dim lSourceSize As Long
Dim iSrc, iTgt, lNumChunks, lBytesLeft, lCntr
Dim sChunck As String * CHUNKSIZE
Dim ch As String * 1
Dim sLastChunck As String

'********* OPEN FILES
iSrc = FreeFile
Open sSourceFile For Binary Access Read As #iSrc
iTgt = FreeFile
Open sTargetFile For Binary Access Write As #iTgt

If lSourceSize = 0 Then
    ' a zero means that we want to take the entire file
    lSourceSize = LOF(iSrc)   ' this is a ByVal parameter, so it's ok to change its value!
End If
lNumChunks = Int(lSourceSize / CHUNKSIZE)
lBytesLeft = lSourceSize - lNumChunks * CHUNKSIZE

Seek #iSrc, lSourceStart + 1    ' The given offsets are 0-based
Seek #iTgt, lTargetStart + 1    ' since all the world is 0-based (except VB!!!)

For lCntr = 1 To lNumChunks
    Get #iSrc, , sChunck
    Put #iTgt, , sChunck
Next lCntr

sLastChunck = Input(lBytesLeft, iSrc)
Put #iTgt, , sLastChunck
Close #iSrc, #iTgt

ReadWriteChunks = True
Exit Function

ReadWriteChunksError:
    ReadWriteChunks = False
    Exit Function
End Function
