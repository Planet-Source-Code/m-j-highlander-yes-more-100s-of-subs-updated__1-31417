Function ParseString (SubStrs() As String, ByVal SrcStr As String, ByVal Delimiter As String) As Integer
'This routine returns an empty string if it finds two consecutive delimiters
' Dimension variables:
ReDim SubStrs(0) As String
Dim CurPos As Long
Dim NextPos As Long
Dim DelLen As Integer
Dim nCount As Integer
Dim TStr As String



    ' Add delimiters to start and end of string to make loop simpler:
    SrcStr = Delimiter & SrcStr & Delimiter
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

