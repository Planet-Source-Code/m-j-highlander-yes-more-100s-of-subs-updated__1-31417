Declare Function sndPlaySound Lib "mmsystem.dll" (ByVal lpRes As Any, ByVal wflags As Integer) As Integer

Declare Function LoadLibrary Lib "kernel" (ByVal lpLibFileName As String) As Integer

Declare Function FindResource Lib "kernel" (ByVal hInstance As Integer, ByVal lpname As String, ByVal lpType As Any) As Integer

Declare Function LoadResource Lib "kernel" (ByVal hInstance As Integer, ByVal hResInfo As Integer) As Integer

Declare Function LockResource Lib "kernel" (ByVal hResData As Integer) As Long

Declare Function FreeResource Lib "kernel" (ByVal hResData As Integer) As Integer

Declare Sub FreeLibrary Lib "kernel" (ByVal hInstance As Integer)

Const SND_MEMORY = 4
Const SND_NODEFAULT = &H2
Const SND_ASYNC = &H1

Sub PlayDllWaveMem (DLLName As String, WaveName As String)

Dim hInstance As Integer
Dim hResInfo As Integer
Dim hRes As Integer
Dim lpRes As Long
Dim iReturn As Integer

    hInstance = LoadLibrary(DLLName)
    hResInfo = FindResource(hInstance, WaveName, "WAVE")
    hRes = LoadResource(hInstance, hResInfo)
    lpRes = LockResource(hRes)
    iReturnVal = sndPlaySound(lpRes, SND_MEMORY Or SND_NODEFAULT Or SND_ASYNC)
    iReturnVal = FreeResource(hRes)
    FreeLibrary (hInstance)

End Sub

