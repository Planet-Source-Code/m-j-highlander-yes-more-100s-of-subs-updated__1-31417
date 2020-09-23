Type FileInfo
    FileName As String
    FileSize As Long
    Flag As Integer
End Type

Global FInfo As FileInfo

Function FC (File1 As String, File2 As String, CompBytes As Integer) As Integer
Dim FirstOffset As Long
Dim Cntr As Integer
Dim FF1 As Integer, FF2 As Integer
Dim ch1 As String * 1
Dim ch2 As String * 1

If CompBytes = 0 Then
    CompBytes = 10
End If

ReDim Offsets(1 To CompBytes)
FirstOffset = Int(FileLen(File1)) / CompBytes
For Cntr = 1 To CompBytes
    Offsets(Cntr) = FirstOffset * Cntr
Next Cntr

FF1 = FreeFile
FF2 = FF1 + 1
Open File1 For Binary As FF1
Open File2 For Binary As FF2

For Cntr = 1 To CompBytes
    Get #FF1, Offsets(Cntr), ch1
    Get #FF2, Offsets(Cntr), ch2
    If ch1 <> ch2 Then
        FC = False
        Exit Function
    End If
Next Cntr

FC = True
End Function

