Attribute VB_Name = "Module1"
' WAVE Playing Sub for 32 bit VB
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long




Public Sub Play(FileName)
'Play Sound Asynchronously (SND_ASYNC = &H1)
'To stop playing a sound call PLAY ""

sndPlaySound FileName, &H1

End Sub


