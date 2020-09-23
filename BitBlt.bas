Declare Function BitBlt Lib "GDI" (ByVal hDestDC As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal hSrcDC As Integer, ByVal XSrc As Integer, ByVal YSrc As Integer, ByVal dwRop As Long) As Integer

Const BLACKNESS = &H42
Const DSINVERT = &H550009
Const MERGECOPY = &HC000CA
Const MERGEPAINT = &HBB0226
Const NOTSRCCOPY = &H330008
Const NOTSRCERASE = &H1100A6
Const PATCOPY = &HF00021
Const PATINVERT = &H5A0049
Const PATPAINT = &HFB0A09
Const SRCAND = &H8800C6
Const SRCCOPY = &HCC0020
Const SRCERASE = &H4400328
Const SRCINVERT = &H660046
Const SRCPAINT = &HEE0086
Const WHITENESS = &HFF0062

