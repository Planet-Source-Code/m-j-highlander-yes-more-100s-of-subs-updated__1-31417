Sub GetBitmapInfo (bmpFileName, vRes As Long, hRes As Long, bits As Integer)
Dim i As Integer
i = FreeFile
Open bmpFileName For Binary As i
Get i, 19, hRes  'vertical resolution
Get i, 23, vRes  'horizontal resolution
Get i, 29, bits  'bits-per-pixel
Close i

End Sub

