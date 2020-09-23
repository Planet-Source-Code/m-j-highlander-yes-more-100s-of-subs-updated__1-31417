' Module1
Option Explicit

Function fnC8 (p00EE As Form, p00F0, p) As String
p00EE.CMDialog1.CancelError = True
p00EE.CMDialog1.Flags = gv0026 Or gv0042
p00EE.CMDialog1.Filter = p00F0
p00EE.CMDialog1.FileName = ""
On Error GoTo L9B0
p00EE.CMDialog1.Action = 1
If  p00EE.CMDialog1.FileName = "" Then Exit Function
fnC8 = p00EE.CMDialog1.FileName
Exit Function
L9B0:
Exit Function
Resume Next
Exit Function
End Function

Function fn100 (p00F8 As Form, p00FA) As String
p00F8.CMDialog1.CancelError = True
p00F8.CMDialog1.Flags = gv0026 Or gv0022
p00F8.CMDialog1.Filter = p00FA
p00F8.CMDialog1.FileName = ""
On Error GoTo LA7E
p00F8.CMDialog1.Action = 2
If  p00F8.CMDialog1.FileName = "" Then Exit Function
fn100 = p00F8.CMDialog1.FileName
Exit Function
LA7E:
Exit Function
Resume Next
Exit Function
End Function
