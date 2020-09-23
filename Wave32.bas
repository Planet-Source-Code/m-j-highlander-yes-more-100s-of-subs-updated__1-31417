Attribute VB_Name = "Module1"
Declare Function FindResource Lib "kernel32" Alias "FindResourceA" (ByVal hInstance As Long, ByVal lpName As String, ByVal lpType As String) As Long
Declare Function LoadResource Lib "kernel32" (ByVal hInstance As Long, ByVal hResInfo As Long) As Long
Declare Function LockResource Lib "kernel32" (ByVal hResData As Long) As Long
Declare Function FreeResource Lib "kernel32" (ByVal hResData As Long) As Long

Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long


Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long




Public Const SND_NODEFAULT = &H2         '  silence not default, if sound not found
Public Const SND_ASYNC = &H1         '  play asynchronously



Public Const SND_RESOURCE = &H40004

Public Const SND_PURGE = &H40

Sub PlayWave(DLLName As String, WaveRes As String)
    Dim bRtn As Long
    Dim hInst As Long
     
    hInst = LoadLibrary(DLLName)
   
    bRtn = PlaySound(WaveRes, hInst, SND_RESOURCE Or SND_ASYNC Or SND_NODEFAULT)

    FreeLibrary (hInst)
   
End Sub

Sub StopPlay()
  Dim bRtn As Long
  bRtn = PlaySound("", 0, SND_PURGE)
  
End Sub


