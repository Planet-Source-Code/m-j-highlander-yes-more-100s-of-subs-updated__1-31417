Declare Function ExtFloodFill Lib "GDI" (ByVal hDC%, ByVal i%, ByVal i%, ByVal w&, ByVal i%) As Integer
 
Const black = &H0&
Const red = &HFF&
Const green = &HFF00&
Const yellow = &HFFFF&
Const blue = &HFF0000
Const magenta = &HFF00FF
Const cyan = &HFFFF00
Const white = &HFFFFFF

Sub Paint (PicCtl As control, X As Single, Y As Single, crColor As Long, wFillType As Integer)
'wFillType:
'FLOODFILLBORDER  = 0   Fill until crColor& color encountered
'FLOODFILLSURFACE = 1   Fill surface until crColor& color not
'crColor& = RGB()  Color to look for
Dim Result As Integer

Result = ExtFloodFill(PicCtl.hDC, X, Y, crColor&, wFillType%)

End Sub

