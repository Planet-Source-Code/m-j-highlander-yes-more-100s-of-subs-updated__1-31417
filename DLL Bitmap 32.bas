Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function LoadBitmap Lib "user32" Alias "LoadBitmapA" (ByVal hInstance As Long, ByVal lpBitmapName As String) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long

Sub LoadDLLBitmap32 (Dll as string, BMP as string , Pic as Control)
d& = LoadLibrary(DLL)
x& = LoadBitmap(d&, BMP)
dc& = GetDC(Pic.hwnd)
cdc& = CreateCompatibleDC(dc&)
o& = SelectObject(cdc&, x&)
bb& = BitBlt(dc&, 0, 0, 640, 480, cdc&, 0, 0, &HCC0020)
z& = SelectObject(cdc&, o&)
z& = DeleteDC(cdc&)
z& = ReleaseDC(Pic.hwnd, dc&)

FreeLibrary d&
End Sub


