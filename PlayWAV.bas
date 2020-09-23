Declare Function sndPlaySound Lib "MMSYSTEM" (ByVal lpszSoundName As String, ByVal uFlags As Integer) As Integer

Sub Play (FileName As Variant)
'Plays a WAV file
'To call use code like: Play "sound.wav"

    X = sndPlaySound(FileName, &H1)
End Sub

