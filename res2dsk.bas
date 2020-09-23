Declare Function GlobalLock Lib "Kernel" (ByVal hMem As Integer) As Long
Declare Function GlobalUnlock Lib "Kernel" (ByVal hMem As Integer) As Integer
Declare Function GlobalFree Lib "Kernel" (ByVal hMem As Integer) As Integer
Declare Function FreeResource Lib "Kernel" (ByVal hResData As Integer) As Integer
Declare Function SizeOfResource Lib "Kernel" (ByVal hInstance As Integer, ByVal hResInfo As Integer) As Integer
Declare Function FindResource% Lib "Kernel" (ByVal hInstance%, ByVal lpName$, ByVal lpType As Any)
Declare Function LoadResource% Lib "Kernel" (ByVal hInstance%, ByVal hResInfo%)
Declare Function LockResource& Lib "Kernel" (ByVal hResData%)
Declare Function sndplaysound Lib "mmsystem" (ByVal filenae As Any, ByVal snd_async As Any) As Integer
Declare Function LoadLibrary Lib "Kernel" (ByVal lpLibFileName As String) As Integer
Declare Function LoadBitmap Lib "User" (ByVal hInstance As Integer, ByVal lpBitmapName As Any) As Integer
Declare Function GetObj Lib "GDI" Alias "GetObject" (ByVal hObject As Integer, ByVal nCount As Integer, lpObject As Any) As Integer
Declare Function CreateCompatibleDC Lib "GDI" (ByVal hDC As Integer) As Integer
Declare Function SelectObject Lib "GDI" (ByVal hDC As Integer, ByVal hObject As Integer) As Integer
Declare Function BitBlt Lib "GDI" (ByVal hDestDC As Integer, ByVal x As Integer, ByVal y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal hSrcDC As Integer, ByVal XSrc As Integer, ByVal YSrc As Integer, ByVal dwRop As Long) As Integer
Declare Function DeleteDC Lib "GDI" (ByVal hDC As Integer) As Integer
Declare Sub FreeLibrary Lib "Kernel" (ByVal hLibModule As Integer)
Declare Function DeleteObject Lib "GDI" (ByVal hObject As Integer) As Integer
Declare Function LoadIcon% Lib "User" (ByVal hInstance%, ByVal lpIconName As Any)
Declare Function DrawIcon% Lib "User" (ByVal hDC%, ByVal x%, ByVal y%, ByVal hicon%)



' OpenFile() Structure
Type OFSTRUCT

      cBytes As String * 1
      fFixedDisk As String * 1
      nErrCode As Integer
      reserved As String * 4
      szPathName As String * 128

End Type


' OpenFile() Flags
Global Const OF_READ = &H0
Global Const OF_WRITE = &H1
Global Const OF_READWRITE = &H2
Global Const OF_SHARE_COMPAT = &H0
Global Const OF_SHARE_EXCLUSIVE = &H10
Global Const OF_SHARE_DENY_WRITE = &H20
Global Const OF_SHARE_DENY_READ = &H30
Global Const OF_SHARE_DENY_NONE = &H40
Global Const OF_PARSE = &H100
Global Const OF_DELETE = &H200
Global Const OF_VERIFY = &H400
Global Const OF_CANCEL = &H800
Global Const OF_CREATE = &H1000
Global Const OF_PROMPT = &H2000
Global Const OF_EXIST = &H4000
Global Const OF_REOPEN = &H8000

Declare Function OpenFile Lib "Kernel" (ByVal lpFilename As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Integer) As Integer
Declare Function hRead Lib "kernel" Alias "_hread" (ByVal hFile As Integer, lpMem As Any, ByVal lSize As Long) As Long
Declare Function hWrite Lib "Kernel" Alias "_hwrite" (ByVal hFile As Integer, lpMem As Any, ByVal lSize As Long) As Long
Declare Function lClose Lib "kernel" Alias "_lclose" (ByVal hFile As Integer) As Integer

' Global Memory Flags
Global Const GMEM_FIXED = &H0
Global Const GMEM_MOVEABLE = &H2
Global Const GMEM_NOCOMPACT = &H10
Global Const GMEM_NODISCARD = &H20
Global Const GMEM_ZEROINIT = &H40
Global Const GMEM_MODIFY = &H80
Global Const GMEM_DISCARDABLE = &H100
Global Const GMEM_NOT_BANKED = &H1000
Global Const GMEM_SHARE = &H2000
Global Const GMEM_DDESHARE = &H2000
Global Const GMEM_NOTIFY = &H4000
Global Const GMEM_LOWER = GMEM_NOT_BANKED

Global Const GHND = (GMEM_MOVEABLE Or GMEM_ZEROINIT)
Global Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)

Function SaveRes (DLLName As String, ResName As String, ResType As String, XResSize As Long, OutFile As String) As Long

'Sample Call:
'Result& = SaveRes("test16.dll", "JPEG_1", "JPEG",0, "D:\pic1.jpg")

' The Function SIZEOFRESOURCE() fails if the resource is larger than 32 KB
' however, if the resource was less than 64 KB specify 0 for the XResSize
' argument and it will be automatically corrected.
' But if the resource was larger than 64KB then you should supply it's size
' or else the results are not guarenteed.

Dim hLoadWave As Integer
Dim hwaveres As Integer
Dim hSound As Long
Dim hRelease As Integer
Dim res As Integer
Dim hLibInst As Integer
Dim InpFile As String
Dim hFile As Integer
Dim fileStruct As OFSTRUCT
Dim FSize As Long
Dim BytesRead As Long
Dim BytesWritten As Long
Dim hMem As Integer
Dim lpMem As Long
Dim r As Integer

hLibInst = LoadLibrary(DLLName)
hwaveres = FindResource(hLibInst, ResName, ResType)

FSize = SizeOfResource(hLibInst, hwaveres)

If (XResSize = 0 And FSize < 0) Then
    FSize = 64& * 1024& + FSize 'add 64 KB
ElseIf XResSize <> 0 Then
    FSize = XResSize
Else
    'XResSize=0 & FSize >0
    'do nothing
End If


hLoadWave = LoadResource(hLibInst, hwaveres)
hSound = LockResource(hLoadWave)

hFile = OpenFile(OutFile, fileStruct, OF_CREATE Or OF_WRITE Or OF_SHARE_DENY_NONE)

BytesWritten = hWrite(hFile, ByVal hSound, FSize)

r = lClose(hFile)

'hrelease = GlobalUnlock(hloadwave)
hRelease = FreeResource(hLoadWave)

FreeLibrary (hLibInst)

SaveRes = BytesWritten
End Function

