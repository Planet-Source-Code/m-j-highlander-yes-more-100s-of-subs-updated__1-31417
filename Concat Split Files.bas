Sub ConcatSplitFiles (firstfile$, cSplit%)
    Dim x%, fh1%, fh2%, outfile$, outfileLen&, CopyLeftOver&, CopyChunk#, filevar$
    Dim iFileMax%, iFile%, y%

    For x% = 2 To cSplit%
    
        fh1% = FreeFile
        Open Left$(firstfile$, Len(firstfile$) - 1) + Format$(1) For Binary As fh1%
                
        fh2% = FreeFile
        outfile$ = Left$(firstfile$, Len(firstfile$) - 1) + Format$(x%)
        Open outfile$ For Binary As fh2%
            
        ' Goto the end of file (plus one bytes) to start writing data
        Seek #fh1%, LOF(fh1%) + 1

        outfileLen& = LOF(fh2%)
        CopyLeftOver& = outfileLen& Mod 10
        CopyChunk# = (outfileLen& - CopyLeftOver&) / 10
        filevar$ = String$(CopyLeftOver&, 32)
        Get #fh2%, , filevar$
        Put #fh1%, , filevar$
        filevar$ = String$(CopyChunk#, 32)
        iFileMax% = 10
        For iFile% = 1 To iFileMax%
            Get #fh2%, , filevar$
            Put #fh1%, , filevar$
        Next iFile%

        Close fh1%, fh2%
        y% = SetTime(outfile$, firstfile$)
        Kill outfile$

    Next x%
    
    FileCopy Left$(firstfile$, Len(firstfile$) - 1) + Format$(1), firstfile$
    Kill Left$(firstfile$, Len(firstfile$) - 1) + Format$(1)
End Sub
