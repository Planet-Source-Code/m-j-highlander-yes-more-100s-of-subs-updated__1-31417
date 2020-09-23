
Function iIsValidFileName (FName As String) As Integer
' Check if FName is a valid file name
' Returns True Or False

On Error Resume Next

Const ATTR_NORMAL = 0
Const ATTR_HIDDEN = 2
Const ATTR_SYSTEM = 4

Dim ff As Integer
Dim Exists As Integer
Dim Attr_Mask As Integer

Attr_Mask = ATTR_NORMAL + ATTR_HIDDEN + ATTR_SYSTEM
If Dir$(FName, Attr_Mask) <> "" Then
    Exists = True
Else
    Exists = False
End If
ff = FreeFile
Open FName For Binary As #ff
If Err Then
    'not a valid name
    iIsValidFileName = False
    Exit Function
Else
    iIsValidFileName = True
    Close #ff
    If Exists = False Then Kill FName
End If

End Function

Function sIsValidFileName (FName As String) As String
' Check if FName is a valid file name
' Returns the Actual file (might differ from FName) name or an empty string

On Error Resume Next

Const ATTR_NORMAL = 0
Const ATTR_HIDDEN = 2
Const ATTR_SYSTEM = 4

Dim ff As Integer
Dim Exists As Integer
Dim Attr_Mask As Integer

Attr_Mask = ATTR_NORMAL + ATTR_HIDDEN + ATTR_SYSTEM
If Dir$(FName, Attr_Mask) <> "" Then
    Exists = True
Else
    Exists = False
End If
ff = FreeFile
Open FName For Binary As #ff
If Err Then
    'not a valid name
    sIsValidFileName = ""
    Exit Function
Else
    sIsValidFileName = Dir(FName)
    Close #ff
    If Exists = False Then Kill FName
End If

End Function

