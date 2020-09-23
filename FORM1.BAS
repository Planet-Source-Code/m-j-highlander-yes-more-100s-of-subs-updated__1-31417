' FORM1.FRM
Option Explicit

Sub Form_Load ()
Dim l0020 As Integer
Dim l0022 As String
Dim l0024 As String
Dim l0026 As String
Dim l002C As Integer
Dim l002E As String
Dim l0030
Dim l0032 As Variant
Picture1.AutoSize = False
Picture1.AutoSize = True
Me.Width = Picture1.Width + 120
Me.Height = Picture1.Height + 340
sub100 Me
Me.Show
Me.Refresh
l0020 = FreeFile
l0022 = App.Path + "\crack.dat"
Open l0022 For Input As l0020
Line Input #l0020, l0024
Line Input #l0020, l0026
Close #l0020
l002E = Date$
Date$ = l0026
l002C = Shell(l0024, 5)
l0032 = DoEvents()
l0032 = DoEvents()
l0032 = DoEvents()
Date$ = l002E
For l0030 = 1 To 10000
    l0032 = DoEvents()
Next l0030
End
Exit Sub
LC72:
MsgBox "CRACK.DAT not found or EXE File not found. quitting...", 16, Me.Caption
End
End Sub
