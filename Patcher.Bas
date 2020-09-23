Attribute VB_Name = "Patchers"
Option Explicit

Type FileOffsetRec
    FileName As String
    PatchOffset As Long
End Type

Public FileOffsetRecArray() As FileOffsetRec

Sub ParsePatchList(PatchListFile As String)
Dim iFF As Integer
Dim sTemp As String
Dim idx As Integer
iFF = FreeFile
Open PatchListFile For Input Access Read As #iFF
idx = 0
ReDim FileOffsetRecArray(1 To 1)  'create array
Do While Not EOF(iFF)
    Line Input #iFF, sTemp
    If Left(sTemp, 1) = ";" Then
        'a comment, ignore it.
    ElseIf Trim(sTemp) = "" Then
        'an empty line, ignore it
    Else
        idx = idx + 1
        'MsgBox sTemp & "       " & Str(idx)

        ReDim Preserve FileOffsetRecArray(1 To idx)
        SplitString sTemp, FileOffsetRecArray(idx).FileName, FileOffsetRecArray(idx).PatchOffset
        'MsgBox "left " & FileOffsetRecArray(idx).FileName
        'MsgBox "right " & FileOffsetRecArray(idx).PatchOffset
    End If
Loop
Close #iFF

End Sub
Function PatchFile(srcFile As String, DestFile As String, PatchOffset As Long) As Boolean
'MsgBox srcFile
'MsgBox tgtFile
'MsgBox Str(PatchOffset)
On Error GoTo PatchError
Dim srcSize As Long
Dim src%, tgt%, NumChunks&, BytesLeft&, cntr&
Const CHUNKSIZE = 1000

Dim Chunck As String * CHUNKSIZE
Dim ch As String * 1
Dim LastChunck As String

srcSize = FileLen(srcFile)
src = FreeFile
Open srcFile$ For Binary As #src
NumChunks& = Int(srcSize / CHUNKSIZE)
BytesLeft& = srcSize - NumChunks& * CHUNKSIZE

tgt = FreeFile
Open DestFile$ For Binary As #tgt
Seek #tgt, PatchOffset + 1
For cntr& = 1 To NumChunks&
    Get #src, , Chunck$
    Put #tgt, , Chunck$
Next cntr&

LastChunck = Input(BytesLeft&, src)
Put #tgt, , LastChunck
Close src, tgt
PatchFile = True
Exit Function

''''''Error Handler
PatchError:
    PatchFile = False
    Close #src, #tgt
    Exit Function

End Function


Sub ProccessCommandLine()

ReDim Args(1 To 1) As String
Dim ArgCount As Byte
ArgCount = ParseString(Args(), Command$, ",")

''''''''''''X
'Dim i
'MsgBox Str(ArgCount)
'For i = 1 To ArgCount
'MsgBox Args(i)
'Next
''''''''''''X

Dim srcFile As String
Dim tgtFile As String
Dim PatchOffset As Long
Dim bResult As Boolean
Dim idx As Integer

srcFile = Trim$(Args(1))
tgtFile = Trim$(Args(2))

If (srcFile = "" Or tgtFile = "") Then
    MsgBox "Please specify Input and Output files", 16, "Error"
    Exit Sub
End If
If Dir(srcFile, ATTR_ALL_FILES) = "" Then
    MsgBox "File Not Found  '" + srcFile + "'", 16, "Error"
    Exit Sub
End If
If Dir(tgtFile, ATTR_ALL_FILES) = "" Then
    MsgBox "File Not Found  '" + tgtFile + "'", 16, "Error"
    Exit Sub
End If
    'ALWAYS create a BackUp?
        FileCopy tgtFile, ChangeFileExtension(tgtFile, "bak")

'''''''''''''''is the input file a Patch List File (*.ptl)?
Select Case LCase(ExtractFileExtension(srcFile))
    Case "ptl"
    'batch patch ;-)***********************************************
    ParsePatchList srcFile
    For idx = LBound(FileOffsetRecArray) To UBound(FileOffsetRecArray)
        bResult = PatchFile(FileOffsetRecArray(idx).FileName, tgtFile, FileOffsetRecArray(idx).PatchOffset)
    Next idx
    MsgBox Str(idx - 1) & " Patches complete!", 64
    '**************************************************************
    Case Else
    'single patch
    If ArgCount < 3 Then
        MsgBox "Patch Offset required for a single patch operation", 16, "Error"
        Exit Sub
    End If
    PatchOffset = Val(Args(3))
    bResult = PatchFile(srcFile, tgtFile, PatchOffset)
    If bResult Then MsgBox "Patch operation complete!", 64 Else MsgBox "Error patching, aborting.", 16
End Select

End Sub

Sub SplitString(sString As String, sLeft As String, lRight As Long, Optional sDelemeter As Variant)
If IsMissing(sDelemeter) Then sDelemeter = ","
Dim iPos As Integer
Dim sRight As String
iPos = InStr(sString, sDelemeter)
If iPos = 0 Then
    sLeft = ""
    sRight = ""
    Exit Sub
End If
sLeft = Trim$(Left$(sString, iPos - 1))
sRight = Trim$(Right$(sString, Len(sString) - iPos))
lRight = Val(sRight)

End Sub


