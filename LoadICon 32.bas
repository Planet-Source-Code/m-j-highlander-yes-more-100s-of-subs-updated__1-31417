Attribute VB_Name = "Module1"
Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Declare Function LoadIcon Lib "user32" Alias "LoadIconA" (ByVal hInstance As Long, ByVal lpIconName As String) As Long
Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long


Sub LoadDLLIcon32 (DLLName As String, IconName As String, pic As Control)
l& = LoadLibrary(DLLName)
i& = LoadIcon(l&, IconName)
d& = DrawIcon(pic.hdc, 0, 0, i&)
FreeLibrary l&

End Sub


