
Sub RWChuncksEx (SrcFile$, DestFile$, Start&, Size&)
'if DestFile$ exists the chunks will be appended to it rather than replacing it.

Const CHUNKSIZE = 1000

Dim Chunck As String * CHUNKSIZE
Dim ch As String * 1
Dim LastChunck As String

src = FreeFile
Open SrcFile$ For Binary As #src
NumChunks& = Int(Size& / CHUNKSIZE)
BytesLeft& = Size& - NumChunks& * CHUNKSIZE

tgt = FreeFile
Open DestFile$ For Binary As #tgt
If LOF(tgt) <> 0 Then Seek #tgt, LOF(tgt) + 1
Seek #src, Start&

For cntr& = 1 To NumChunks&
    Get #src, , Chunck$
    Put #tgt, , Chunck$
    
Next cntr&

LastChunck = Input(BytesLeft&, src)
Put #tgt, , LastChunck

Close src, tgt

End Sub

