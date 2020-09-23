Option Explicit

Declare Function LoadLibrary Lib "Kernel" (ByVal lpLibFileName As String) As Integer
Declare Sub FreeLibrary Lib "Kernel" (ByVal hLibModule As Integer)

Declare Function CreateCompatibleDC Lib "GDI" (ByVal hDC As Integer) As Integer
Declare Function DeleteDC Lib "GDI" (ByVal hDC As Integer) As Integer

Declare Function SelectObject Lib "GDI" (ByVal hDC As Integer, ByVal hObject As Integer) As Integer

Declare Function LoadBitmap Lib "User" (ByVal hInstance As Integer, ByVal lpBitmapName As Any) As Integer

Declare Function BitBlt Lib "GDI" (ByVal hDestDC As Integer, ByVal x As Integer, ByVal Y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal hSrcDC As Integer, ByVal XSrc As Integer, ByVal YSrc As Integer, ByVal dwRop As Long) As Integer


Global Const SRCCOPY = &HCC0020  'DWORD

Sub LoadDLLBitmapRedraw (DLLName As String, BMPName As String, pic As Control)

Dim hDLL%, hBitmap%, hOld%, Result%, hDCCompat%
    
    'this Sub works -even if pic is hidden- only if:
    pic.AutoRedraw = True
    
    hDLL% = LoadLibrary(DLLName)
    hBitmap% = LoadBitmap(hDLL%, BMPName)
    
    'Create a device context compatible with the picture box
    hDCCompat% = CreateCompatibleDC(pic.hDC)
    
    'Put the DLL bitmap in this device context
    hOld% = SelectObject(hDCCompat%, hBitmap%)
    
    'Copy the bitmap from the device context to the picture
    Result% = BitBlt(pic.hDC, 0, 0, 640, 480, hDCCompat%, 0, 0, SRCCOPY)
    
    'Bitmap woun't show without refreshing
    pic.Refresh
    
    'Select the bitmap out of the device context
    Result% = SelectObject(hDCCompat%, hOld%)
    
    'Delete the device context
    Result% = DeleteDC(hDCCompat%)
    
    FreeLibrary hDLL%

End Sub

