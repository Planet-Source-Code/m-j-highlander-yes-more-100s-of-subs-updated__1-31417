Declare Function StretchBlt% Lib "GDI" (ByVal hDC%, ByVal X%, ByVal Y%, ByVal nWidth%, ByVal nHeight%, ByVal hSrcDC%, ByVal XSrc%, ByVal YSrc%, ByVal nSrcWidth%, ByVal nSrcHeight%, ByVal dwRop&)

Declare Function bitBlt Lib "GDI" (ByVal hDestDC As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal hSrcDC As Integer, ByVal XSrc As Integer, ByVal YSrc As Integer, ByVal dwRop As Long) As Integer

Global Const BLACKNESS = &H42
Global Const DSINVERT = &H550009
Global Const MERGECOPY = &HC000CA
Global Const MERGEPAINT = &HBB0226
Global Const NOTSRCCOPY = &H330008
Global Const NOTSRCERASE = &H1100A6
Global Const PATCOPY = &HF00021
Global Const PATINVERT = &H5A0049
Global Const PATPAINT = &HFB0A09
Global Const SRCAND = &H8800C6
Global Const SRCCOPY = &HCC0020
Global Const SRCERASE = &H4400328
Global Const SRCINVERT = &H660046
Global Const SRCPAINT = &HEE0086
Global Const WHITENESS = &HFF0062

Sub Strech (tgt As PictureBox, src As PictureBox)

F% = StretchBlt(tgt.hDC, 0, 0, tgt.Width, tgt.Height, src.hDC, 0, 0, src.Width, src.Height, &HCC0020)
tgt.Refresh

End Sub

