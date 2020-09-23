Option Explicit
Declare Function GetSystemDirectory Lib "kernel" (ByVal Buff As String, ByVal nsize As Integer) As Integer
Global Const ATTR_READONLY = 1    'Read-only file
Global Const ATTR_VOLUME = 8   'Volume label
Global Const ATTR_ARCHIVE = 32     'File has changed since last back-up
Global Const ATTR_NORMAL = 0   'Normal files
Global Const ATTR_HIDDEN = 2   'Hidden files
Global Const ATTR_SYSTEM = 4   'System files
Global Const ATTR_DIRECTORY = 16   'Directory

Global Const ATTR_DIR_ALL = ATTR_DIRECTORY + ATTR_READONLY + ATTR_ARCHIVE + ATTR_HIDDEN + ATTR_SYSTEM
Global Const ATTR_ALL_FILES = ATTR_NORMAL Or ATTR_HIDDEN Or ATTR_SYSTEM Or ATTR_READONLY Or ATTR_ARCHIVE

Sub CopyDLLs ()
Dim f As String
Dim idx As Integer
ReDim DLL_Array(1 To 100) As String
idx = 0
f = Dir$(GetAppPath() + "dll\*.*", ATTR_ALL_FILES)
Do Until f = ""
    idx = idx + 1
    DLL_Array(idx) = f
    'Print DLL_Array(idx), "*"
    f = Dir
Loop

ReDim Preserve DLL_Array(1 To idx)

For idx = LBound(DLL_Array) To UBound(DLL_Array)
    If FileExists(GetSysDir() & DLL_Array(idx)) Then
        'file exists
        If FileLen(GetSysDir() & DLL_Array(idx)) <> FileLen(GetAppPath() & "dll\" & DLL_Array(idx)) Then
            'different size, ///dirty method to check for a different version!
            'copy file
            On Error Resume Next   ' in case destination was read-only or in-use
            FileCopy GetAppPath() & "dll\" & DLL_Array(idx), GetSysDir() & DLL_Array(idx)
            On Error GoTo 0
        End If
    Else
        'file does not exist, so copy it
        FileCopy GetAppPath() & "dll\" & DLL_Array(idx), GetSysDir() & DLL_Array(idx)
        'Print DLL_Array(idx)
    End If
Next idx

End Sub

Function DirExists (sDir As String) As Integer
Dim tmp As String
Dim iResult As Integer

iResult = 0
If Dir$(sDir, ATTR_DIR_ALL) <> "" Then
    iResult = GetAttr(sDir) And ATTR_DIRECTORY
End If

If iResult = 0 Then   'Directory not found, or the passed argument is a filename not a directory
    DirExists = False
Else
    DirExists = True
End If


End Function

Function FileExists (sFile As String) As Integer

If Dir$(sFile, ATTR_ALL_FILES) = "" Then
    FileExists = False
Else
    FileExists = True
End If

End Function

Function GetAppPath ()

If Len(app.Path) = 3 Then
    GetAppPath = app.Path
Else
    GetAppPath = app.Path + "\"
End If

End Function

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

Sub Main ()

CopyDLLs

End Sub

