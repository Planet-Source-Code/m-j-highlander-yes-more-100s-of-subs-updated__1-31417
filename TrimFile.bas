Sub TrimFileLines (InFile as string, OutFile as string,sTrim as string)
dim FF1 as integer,FF2 as integer
dim ReadLn as string


FF1=freefile
Open InFile For Input As #FF1
FF1=freefile
Open OutFile For Output As #FF2
Do While Not EOF(FF1)
	Line Input #FF1, ReadLn 
	Select Case lcase(sTrim)
		Case "l","left"
			ReadLn  = LTrim(ReadLn)
			Print #FF2, ReadLn
		Case "r","right"
			ReadLn  = RTrim(ReadLn)
			Print #FF2, ReadLn
		Case "b","both","lr","rl","leftright","rightleft",""
			ReadLn  = Trim(ReadLn)
			Print #FF2, ReadLn
	end select
Loop

Close #FF1,#FF2

End Sub
