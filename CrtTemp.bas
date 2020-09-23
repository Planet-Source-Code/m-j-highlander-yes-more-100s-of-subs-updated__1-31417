Declare Function GetTempFileName Lib "Kernel" (ByVal cDriveLetter As Integer, ByVal lpPrefixString As String, ByVal wUnique As Integer, ByVal lpTempFileName As String) As Integer

Function CreateTempFile () As String
Dim iResult As Integer
Dim sTempFileName As String

sTempFileName = String(250, " ")
iResult = GetTempFileName(0, "abc", 0, sTempFileName)

CreateTempFile = Trim(sTempFileName)
End Function

