Sub Encrypt (Secret$, PassWord$)
' Secret$ = the string you wish to encrypt or decrypt.
' PassWord$ = the password with which to encrypt the string.
' Calling the sub on a string Encrypts it
' Calling the sub on an encrypted  string Decrypts it

Dim L As Long
Dim X As Long
Dim iChar As Integer

    L = Len(PassWord$)
    For X = 1 To Len(Secret$)
        iChar = Asc(Mid$(PassWord$, (X Mod L) - L * ((X Mod L) = 0), 1))
        Mid$(Secret$, X, 1) = Chr$(Asc(Mid$(Secret$, X, 1)) Xor iChar)
    Next X

End Sub

Sub main ()
EncStr$ = Chr$(18) + Chr$(69) + Chr$(9) + Chr$(26) + Chr$(5) + Chr$(12) + Chr$(12) + Chr$(89) + Chr$(27)
EncStr$ = EncStr$ + Chr$(4) + Chr$(27) + Chr$(4) + Chr$(76) + Chr$(13) + Chr$(30) + Chr$(89) + Chr$(62) + Chr$(1) + Chr$(43)
EncStr$ = EncStr$ + Chr$(22)
Encrypt EncStr$, "sexology"

msg = "This executable is not supposed to be launched directly." + Chr(13)
msg = msg + "Please run CD-EJCL.EXE" + Chr(13) + Chr(13)
msg = msg + EncStr$
MsgBox msg, 48, "Oops!"
End Sub

