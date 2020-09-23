
Sub EnhPrint (text, FontN, FontS, header, Footer)
Printer.FontName = FontN
Printer.FontSize = FontS
htw = Printer.TextWidth(header)
Printer.CurrentX = (Printer.ScaleWidth - htw) / 2
Printer.Print header
Printer.Print
ptw = Printer.TextWidth("page   ")
pnc = (Printer.ScaleWidth - ptw) / 2
th = Printer.TextHeight("A")
tw = Printer.TextWidth("W")
Printer.CurrentX = 570
st = 1
    Do
    en = InStr(st + 1, text, " ")
    If en = 0 Then Exit Do
    word = Mid(text, st, en - st)
    lgth = Printer.TextWidth(word)
	If Printer.CurrentX + lgth > Printer.ScaleWidth - 2 * tw Then
	    Printer.CurrentX = 570
	    Printer.CurrentY = Printer.CurrentY + th
	    word = LTrim(word)
	End If
    For i = 1 To Len(word)
	    ch$ = Mid$(word, i, 1)
	    If Printer.CurrentX > Printer.ScaleWidth - 2 * tw Then
		Printer.CurrentX = 570
		Printer.CurrentY = Printer.CurrentY + th
	    End If
	    If Printer.CurrentY > Printer.ScaleHeight - 5 * th Then
		Printer.Print
		Printer.CurrentX = pnc
		Printer.Print "page   " + Printer.Page
		Printer.NewPage
		Printer.CurrentX = (Printer.ScaleWidth - htw) / 2
		Printer.Print header
		Printer.Print
		Printer.CurrentX = 570
	    End If
	    If ch$ = Chr$(13) Then
		Printer.CurrentX = 570
		Printer.CurrentY = Printer.CurrentY + th
	    ElseIf ch$ = Chr$(10) Then
	    Else
		Printer.Print ch$;
	    End If
	If (i Mod 100) = 0 Then DoEvents
    Next i
    DoEvents
    st = en
    Loop
lastword = Right$(text, Len(text) - st + 1)
For i = 1 To Len(lastword)
ch$ = Mid$(lastword, i, 1)
If Printer.CurrentX > Printer.ScaleWidth - 2 * tw Then
Printer.CurrentX = 570
Printer.CurrentY = Printer.CurrentY + th
End If
If Printer.CurrentY > Printer.ScaleHeight - 5 * th Then
	Printer.Print
	Printer.CurrentX = pnc
	Printer.Print "page   " + Printer.Page
	Printer.NewPage
	Printer.CurrentX = (Printer.ScaleWidth - htw) / 2
	Printer.Print header
	Printer.Print
	Printer.CurrentX = 570
End If
If ch$ = Chr$(13) Then
Printer.CurrentX = 570
Printer.CurrentY = Printer.CurrentY + th
ElseIf ch$ = Chr$(10) Then
Else
Printer.Print ch$;
End If
Next i
If ((Printer.CurrentY < Printer.ScaleHeight - 5 * th) And Printer.Page <> 1) Then
Printer.CurrentX = pnc
Printer.CurrentY = Printer.ScaleHeight - 4 * th
Printer.Print "page  '+printer.page"
End If
Printer.EndDoc
End Sub

