Declare Function GetSystemDirectory Lib "kernel" (ByVal Buff As String, ByVal nsize As Integer) As Integer

Function GetSysDir () As String
'Returns the System Directory with a trailing "\"
'To call use code like: a = GetSysDir()
    
    Dim Buff As String * 255
    Dim Result As Integer
    Dim SysDir As String
    Result = GetSystemDirectory(Buff, 255)
    SysDir = Left(Buff, Result)
    GetSysDir = SysDir + "\"
    
End Function

