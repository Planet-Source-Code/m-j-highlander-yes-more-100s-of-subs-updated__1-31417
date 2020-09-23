
Sub BackBMP ()
On Error GoTo ErrHandler2

If Dir(GetAppPath() + "gauto.bmp") <> "" Then
    frmMain.Picture = LoadPicture(GetAppPath() + "gauto.bmp")
End If

Exit Sub
ErrHandler2:
MsgBox "The Error '" + UCase$(Error$) + "' occured while loading background bitmap.", 16, "Generic AutoRun"
Exit Sub

End Sub

Sub Check_For_Prev_Instance ()
On Error GoTo ErrHandler3

If app.PrevInstance Then
    MsgBox "Generic AutoRun is Already Running, Quitting.", 48, "Generic AutoRun"
    Xit
End If

Exit Sub
ErrHandler3:
MsgBox "The Error '" + UCase$(Error$) + "' occured while checking for a previous instance of the program.", 16, "Generic AutoRun"
Exit Sub

End Sub

Sub DispX ()

frmMain.picX.Left = 0
frmMain.picX.Top = 0
frmMain.picX.Width = frmMain.ScaleWidth
frmMain.picX.Height = frmMain.ScaleHeight
frmMain.picX.Visible = True

End Sub

Sub Exec (CmdLine As String)
On Error GoTo ErrHandler4

Dim Result As Integer
Result = Shell(GetAppPath() + "exec32.exe " + Chr$(34) + CmdLine + Chr$(34), 1)

Exit Sub
ErrHandler4:
MsgBox "The Error '" + UCase$(Error$) + "' occured while trying to execute 'EXEC32.EXE'.", 16, "Generic AutoRun"
Exit Sub
End Sub

Sub GAutoINI ()
On Error GoTo IniErrHandler
Dim i As Integer
Dim tmp As String
Dim a$, b$, c$, d$
Dim max As Integer
Dim gff As Integer
Dim GAuto As String

GAuto = GetAppPath() + "gauto.ini"
gff = FreeFile

''''''''''''''Get CD Drive Name
Root = UCase(Left(GetAppPath(), 2))
If Command$ <> "" Then Root = Command$

''''''''''''''Read INI Entires
'tmp = StripComment(GetFromINI("version", "Build", GAuto))
'sGAuto_Version = Trim$(tmp)

'If sGAuto_Version = "" Then sGAuto_Version = "4.x"
sGAuto_Version = "5.1"

tmp = StripComment(GetFromINI("options", "cdtitle", GAuto))
If tmp = "" Then tmp = "Generic AutoRun"
frmMain.lblCDTitle.Caption = tmp
frmMain.Caption = tmp

tmp = StripComment(GetFromINI("options", "titlecolor", GAuto))
If tmp = "" Then tmp = "BLACK"
frmMain.lblCDTitle.ForeColor = StringToColor(tmp)

tmp = StripComment(GetFromINI("options", "labelcolor", GAuto))
If tmp = "" Then tmp = "BLACK"
frmMain.lblPath.ForeColor = StringToColor(tmp)
frmMain.lblInfo.ForeColor = StringToColor(tmp)




Open GAuto For Input As #gff

Do While InStr(LCase(tmp), "[contents]") = 0
    'read and ignore
    Line Input #gff, tmp
Loop

i = 1
ReDim ProgArray(1 To 300)  'more than enough

Do While Not EOF(gff)
    Line Input #gff, tmp
	If InStr(tmp, "^") <> 0 Then
	    SplitIntoFour tmp, a, b, c, d
	    ProgArray(i).Prog = a
	    ProgArray(i).folder = Root + b
	    'type only a name and it will be used for Text & Bitmap
	    If d = "" Then d = c
	    If LCase(Right(c, 4)) <> ".txt" Then c = c + ".txt"
	    ProgArray(i).Text = GetAppPath() + c
	    If LCase(Right(d, 4)) <> ".bmp" Then d = d + ".bmp"
	    ProgArray(i).pic = GetAppPath() + d
	    i = i + 1
	End If
Loop

Close #gff

max = i - 1
ReDim Preserve ProgArray(1 To max)
ShellSort ProgArray()

For i = 1 To max
    frmMain.lstProgs.AddItem ProgArray(i).Prog
    frmMain.lstProgs.ItemData(i - 1) = i
Next i



Exit Sub
IniErrHandler:
MsgBox "The Error '" + UCase$(Error$) + "' occured while initializing.", 16, "Generic AutoRun"
Exit Sub


End Sub

Function GetAppPath () As String
Dim ap As String

    ap = app.Path
    If Len(ap) > 3 Then ap = ap + "\"
    GetAppPath = ap

End Function

Function GetRVal (sLine As String) As String
Dim Pos As Integer

    Pos = InStr(sLine, "=")
    GetRVal = Trim(Right(sLine, Len(sLine) - Pos))
    
End Function

Sub Install_MP3 ()
On Error GoTo MP3ErrHandler

Dim tmp As String
Dim tmp2 As String
Dim MP3 As Integer
Dim GAuto As String
Dim Result As Integer

''''''''does GAUTO.INI instructs us to install MP3 ???
GAuto = GetAppPath() + "gauto.ini"
tmp = StripComment(GetFromINI("options", "MP3", GAuto))
If tmp = "" Then tmp = "0"
MP3 = CInt(Trim$(tmp))
If MP3 = 0 Then Exit Sub

'''''''if already installed, Exit
tmp2 = StripComment(GetFromINI("drivers32", "msacm.l3codec", GetWinDir() + "system.ini"))
If LCase$(tmp2) = "l3codecp.acm" Then Exit Sub

'''''''Now the installation part
FileCopy GetAppPath() + "L3codecp.acm", GetSysDir() + "L3codecp.acm"
Result = WriteINI("drivers32", "msacm.l3codec", "l3codecp.acm", GetWinDir() + "system.ini")


Exit Sub
MP3ErrHandler:
MsgBox "The Error '" + UCase$(Error$) + "' occured while trying to install MPEG Layer3 Driver.", 16, "Generic AutoRun"
Exit Sub

End Sub

Sub PlayMusic ()
On Error GoTo MusicError
IsAVI = False

If Dir(GetAppPath() + "gauto.wav") <> "" Then
     ' a WAVE file, notice that the second parameter is a dummy
     ' PlayAVI GetAppPath() + "gauto.wav", screen.Height / screen.TwipsPerPixelY, frmFoo
     frmMain.mci.FileName = GetAppPath() + "gauto.wav"
     frmMain.mci.Command = "open"
     frmMain.mci.Command = "play"

ElseIf Dir(GetAppPath() + "gauto.avi") <> "" Then
     ' an AVI file
     PlayAVI GetAppPath() + "gauto.avi", screen.Height / screen.TwipsPerPixelY, frmFoo
     'PlayAVIFullScreen GetAppPath() + "gauto.avi"
     '''''IsAVI = True   'what was it for ????
End If

Exit Sub
MusicError:
MsgBox "The Error '" + UCase$(Error$) + "' occured while trying to play the startup wave or video.", 16, "Generic AutoRun"
Exit Sub

End Sub

Sub RandomBackBMP ()
'i'm still considering this...
End Sub

Sub ShellSort (SortArray() As progRec)

Dim Row As Integer
Dim MaxRow As Integer
Dim MinRow As Integer
Dim Swtch As Integer
Dim Limit As Integer
Dim Offset As Integer

MaxRow = UBound(SortArray)
MinRow = LBound(SortArray)
Offset = MaxRow \ 2

Do While Offset > 0
      Limit = MaxRow - Offset
      Do
	 Swtch = False
	 
	 For Row = MinRow To Limit
	    If LCase$(SortArray(Row).Prog) > LCase$(SortArray(Row + Offset).Prog) Then
	       Swap SortArray(Row), SortArray(Row + Offset)
	       Swtch = Row
	    End If
	 Next Row

	 
	 Limit = Swtch - Offset
      Loop While Swtch

      
      Offset = Offset \ 2
   Loop
End Sub

Sub SplitIntoFour (sLine$, sVar1$, sVar2$, sVar3$, sVar4$)
' Sample Call:
' SplitIntoFour "ProgName^\Folder1\Foldern^Text^BITMAP", a, b, c, d

Dim pos1%, pos2%, pos3%

    pos1 = InStr(sLine, "^")
    pos2 = InStr(pos1 + 1, sLine, "^")
    pos3 = InStr(pos2 + 1, sLine, "^")
    
    sVar1 = Trim(Left(sLine, pos1% - 1))
    If pos2 = 0 Then
	sVar2 = Trim(Mid(sLine, pos1 + 1, Len(sLine) - pos1))
	sVar3 = ""
	sVar4 = ""
    Else
	sVar2 = Trim(Mid(sLine, pos1 + 1, pos2 - pos1 - 1))
	If pos3 <> 0 Then
	    sVar3 = Trim(Mid(sLine, pos2 + 1, pos3 - pos2 - 1))
	    sVar4 = Trim(Right(sLine, Len(sLine) - pos3))
	Else
	    sVar3 = Trim(Right(sLine, Len(sLine) - pos2))
	    sVar4 = ""
	End If
    End If

'MsgBox sVar1 + "   " + sVar2 + "   " + sVar3 + "   " + sVar4

End Sub

Function StringToColor (sColor) As Long
' no need to define them here, they have been defined as Global
' Const BLACK = &H0&
' Const RED = &HFF&
' Const GREEN = &HFF00&
' Const YELLOW = &HFFFF&
' Const BLUE = &HFF0000
' Const MAGENTA = &HFF00FF
' Const CYAN = &HFFFF00
' Const WHITE = &HFFFFFF
' Const GREY = &HC0C0C0

    Select Case LCase(sColor)
	Case "black", "k"
	    StringToColor = BLACK
	Case "red", "r"
	    StringToColor = RED
	Case "green", "g"
	    StringToColor = GREEN
	Case "yellow", "y"
	    StringToColor = YELLOW
	Case "blue", "b"
	    StringToColor = BLUE
	Case "magenta", "m"
	    StringToColor = MAGENTA
	Case "cyan", "c"
	    StringToColor = CYAN
	Case "white", "w"
	    StringToColor = WHITE
	Case "grey", "e"
	    StringToColor = GREY
    Case Else
	    StringToColor = BLACK
    End Select
End Function

Function StripComment (sLine)
Dim Pos As Integer

    Pos = InStr(sLine, ";")
    If Pos <> 0 Then
	StripComment = Trim(Left(sLine, Pos - 1))
    Else
	StripComment = sLine
    End If

End Function

Sub Swap (Var1 As progRec, Var2 As progRec)
Dim tmp As progRec

    tmp = Var1
    Var1 = Var2
    Var2 = tmp

End Sub

Sub Xbmp ()
On Error GoTo BadError

Dim iByteValue As Integer

If Dir(GetAppPath() + "x.bmp") <> "" Then

    iByteValue = 480
    FileCopy GetAppPath() + "x.bmp", "c:\~~x.tmp"
	Open "c:\~~x.tmp" For Binary As #1
	Put #1, 23, iByteValue
	Close #1
    frmMain.picX.Picture = LoadPicture("c:\~~x.tmp")
    Kill "c:\~~x.tmp"
End If

Exit Sub
BadError:
    ''''if the bitmap on the root is still there, it better not ;-)
    If Dir("c:\~~x.tmp") <> "" Then Kill "c:\~~x.tmp"
    MsgBox "Worst time for an error, almost caught !!!!", , ";-)"
    Exit Sub

End Sub

Sub Xit ()
Unload frmFind
Unload frmText
Unload frmXRated
Unload frmAbout
Unload frmFoo
Unload frmMain
End
End Sub

