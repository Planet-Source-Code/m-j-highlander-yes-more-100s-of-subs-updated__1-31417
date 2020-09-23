Declare Function mciSendString Lib "mmsystem" (ByVal lpstrCommand$, ByVal lpstrReturnStr As String, ByVal wReturnLen%, ByVal hCallBack%) As Long
Declare Function mciGetErrorString Lib "mmsystem" (ByVal dwError&, ByVal lpstrReturnStr As Any, ByVal wReturnLen%) As Long
 
Global Const WS_CHILD = &H40000000

Global h%

Sub CloseAVI ()
ret = mciSendString("close Animation", retstr, 0, 0)
End Sub

Sub main ()
PlayAVIFullScreen "x.avi"
End Sub

Sub PlayAVI (sAviFile, vHeight%, pix As Form)
 
    Dim sRet$
 
   '*** This will open the AVIVideo and create a child window on the
   '*** form where the video will display. Animation is the device_id.
 
    ErrorStr = Space(255)
    retstr = Space(255)
 
    On Error GoTo CloseAVI
 

 
        CmdStr = ("open " & sAviFile & " type AVIVideo alias Animation parent " + LTrim$(Str$(pix.hWnd)) + " style " + Trim$(Str$(WS_CHILD)))
    ret = mciSendString(CmdStr, retstr, 0, 0)
    If ret > 0 Then
        ret = mciGetErrorString(ret, ErrorStr, 255)

    End If
    strr$ = "put Animation window at 0 0 " + Format(0) + " " + Format(vHeight%)
    ret = mciSendString(strr$, retstr, 0, 0)
    If ret > 0 Then
        ret = mciGetErrorString(ret, ErrorStr, 255)

    End If
 
    ret = mciSendString("play Animation ", retstr, 0, 0)
    If ret > 0 Then
        ret = mciGetErrorString(ret, ErrorStr, 255)

    End If
  
CloseAVI:
 
    
    'ret = mciSendString("close Animation", retstr, 0, 0)
    If ret > 0 Then
        ret = mciGetErrorString(ret, ErrorStr, 255)

    End If
 
End Sub

Sub PlayAVIFullScreen (vFileName)

    Dim CmdStr$, ReturnVal&
    CmdStr$ = "play " + vFileName + " fullscreen "
    ReturnVal& = mciSendString(CmdStr$, "", 0, 0&)
    ReturnVal& = mciSendString("close " + vFileName, "", 0, 0)

End Sub

