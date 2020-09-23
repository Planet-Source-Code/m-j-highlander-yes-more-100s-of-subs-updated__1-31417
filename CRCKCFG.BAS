' CRCKCFG.FRM
Option Explicit

Sub Command1_Click ()
Dim l0032 As Integer
Dim l0034 As String
If  Text1.Text = "" Then Beep: Exit Sub
If  Text2.Text = "" Then Beep: Exit Sub
If  Text3.Text = "" Then Beep: Exit Sub
If  Text4.Text = "" Then Beep: Exit Sub
If  Text4.Tag = "" Then Beep: Exit Sub
l0034 = Text4.Tag
l0032 = FreeFile
Open l0034 For Output As l0032
Print #l0032, Text4.Text
Print #l0032, Text1.Text + "-" + Text2.Text + "-" + Text3.Text
Close #l0032
Me.Hide
l0034 = "Now, set your icon up to run crack.exe from the same directory as the " + Chr$(10) + Text4.Text + "." + Chr$(10) + "Then your program will be Date Cracked!"
MsgBox l0034, 64, Me.Caption
End
End
End Sub

Sub Command2_Click ()
End
End Sub

Sub Command3_Click ()
Dim l0038 As String
Dim l003A As Variant
Dim l003E As String
l0038$ = fnC8(Me, "*.exe;*.com|*.exe;*.com", "")
If  l0038$ = "" Then Exit Sub
Text4.Text = l0038$
For l003A = Len(l0038$) To 1 Step -1
    If  Mid$(l0038$, l003A, 1) = "\" Then Exit For
Next l003A
l003E$ = Left$(l0038$, l003A) + "CRACK.DAT"
Text4.Tag = l003E$
End Sub

Sub Form_Load ()
Dim l0040 As Variant
l0040 = Date$
Text1.Text = Left$(Date$, 2)
Text2.Text = Mid$(Date$, 4, 2)
Text3.Text = Right$(Date$, 4)
End Sub
