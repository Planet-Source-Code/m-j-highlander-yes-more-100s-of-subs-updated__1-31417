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

