Option Explicit
Declare Function GetTempFileName Lib "Kernel" (ByVal cDriveLetter As Integer, ByVal lpPrefixString As String, ByVal wUnique As Integer, ByVal lpTempFileName As String) As Integer

Function GetValidTempName () As String

Const PREFIX$ = "fil"
Dim sBuffer As String * 255
Dim iResult As Integer
Dim sTempName As String

    sBuffer = String(255, " ")
    iResult = GetTempFileName(0, PREFIX$, 0, sBuffer)
    sTempName = Trim$(sBuffer)
    Kill sTempName
    
    GetValidTempName = UCase$(sTempName)

End Function

