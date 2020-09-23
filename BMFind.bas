Attribute VB_Name = "BMFIND1"
Type BMOffset
    Offset As Long
    Size As Long
End Type

Dim AllOffsets() As Long
Dim ValidOffsets()  As BMOffset

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


Function GetBitmapInfoEx(ff As Integer, Offset As Long) As Long

Dim vRes As Long
Dim hRes As Long
Dim bits As Integer
Dim Num_Colors As Long
Dim bytes As Single

'FF = FreeFile
'Open bmpFileName For Binary As #FF
Get #ff, Offset + 19, hRes   'horizontal resolution
Get #ff, Offset + 23, vRes   'vertical resolution
Get #ff, Offset + 29, bits   'bits-per-pixel
'MsgBox Str(bits)
'Close #FF

Select Case bits
    Case 1, 4, 8, 16, 24
    'valid bitmap
    Num_Colors = (2 ^ bits) / 8
    bytes = bits / 8
    'the 3000 is just compensation and it is arbitrary
    GetBitmapInfoEx = hRes * vRes * bytes + 3000
Case Else
    GetBitmapInfoEx = 0
End Select

End Function

Sub GetBitmaps(FName As String, tgtFolder As String, BName As String)
Dim Ostr
Dim oi As Integer
Dim vi As Integer

ReDim AllOffsets(1 To 1000)
ReDim ValidOffsets(1 To 1000)

Form1.image1.Visible = False
Form1.lblStatus.Caption = "Searching..."
oi = 0
vi = 0

Const CHUNKSIZE = 1024   'Bytes

Dim ff As Integer
Dim Chunk As String * CHUNKSIZE
Dim pos As Integer

ff = FreeFile

Open FName For Binary As #ff
Do While Not EOF(ff)
    Get #ff, , Chunk
    pos = 0
    Do
         pos = InStr(pos + 1, Chunk, "BM")
         If pos <> 0 Then
                oi = oi + 1
                AllOffsets(oi) = Seek(ff) - CHUNKSIZE + pos - 2
         End If
        
    Loop Until pos = 0
    

Loop

MaxIndex = oi
If MaxIndex = 0 Then
    Form1.lblStatus.Caption = ""
    MsgBox "No bitmaps found", 48, "Finished"
    Form1.image1.Visible = True
    Exit Sub
End If

ReDim Preserve AllOffsets(1 To MaxIndex)


For i = 1 To MaxIndex
    z = GetBitmapInfoEx(ff, AllOffsets(i))
    If z <> 0 Then
        vi = vi + 1
        
        ValidOffsets(vi).Offset = AllOffsets(i) + 1
        ValidOffsets(vi).Size = z
    End If
Next


If vi = 0 Then
    Form1.lblStatus.Caption = ""
    MsgBox "No bitmaps found", 48, "Finished"
    Form1.image1.Visible = True
    Exit Sub
End If

ReDim Preserve ValidOffsets(1 To vi)

If Right(tgtFolder, 1) <> "\" Then tgtFolder = tgtFolder + "\"

For i = 1 To vi
    tgtFile$ = tgtFolder + BName + Format(i, "000") + ".bmp"
    Ostr = Ostr & ExtractFileName(tgtFile$) & "  ,  " & Format(ValidOffsets(i).Offset - 1) & vbCrLf
    RWChuncks FName, tgtFile$, ValidOffsets(i).Offset, ValidOffsets(i).Size
    
    Form1.lblStatus.Caption = Format(i) + " of " + Format(vi) + " Bitmaps Saved"
    Form1.lblStatus.Refresh
Next

'MsgBox Format(vi) + " Bitmaps retrieved", 64, "Done"
Form1.lblStatus.Caption = ""
Form1.lblStatus.Refresh
Form1.image1.Visible = True
Close #ff
Erase AllOffsets, ValidOffsets    ' optional
Ostr = ";" & Format(vi) + " Bitmaps Retrieved:" & vbCrLf & vbCrLf & ";File Name     ,   Offset" & vbCrLf & Ostr & vbCrLf & vbCrLf
Ostr = Ostr & ";(Offsets are calculated considering that the first byte in the file is at the ofsset 0 not 1)" & vbCrLf
TextBox CStr(Ostr)
End Sub

Sub TextBox(Txt As String)
frmTextBox.txtMain.Text = Txt
frmTextBox.Show 1
End Sub


