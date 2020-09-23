Attribute VB_Name = "Registry_API_32"
Option Explicit

Global Const HKEY_CLASSES_ROOT = &H80000000
Global Const HKEY_CURRENT_USER = &H80000001
Global Const HKEY_LOCAL_MACHINE = &H80000002
Global Const HKEY_USERS = &H80000003
Global Const HKEY_DYN_DATA = &H80000004


' Registry API Functions
Declare Function RegEnumKey Lib "advapi32" Alias "RegEnumKeyA" _
   (ByVal hKey As Long, ByVal Index As Long, _
    ByVal RetKey As String, ByVal RetSize As Long) As Long

Declare Function RegOpenKey Lib "advapi32" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal SubKey As String, hOpenKey As Long) As Long

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" _
    (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, _
    ByVal samDesired As Long, ByRef phkResult As Long) As Long

Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" _
    (ByVal hKey As Long, _
     ByVal lpSubKey As String, _
     phkResult As Long) As Long
     
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" _
    (ByVal hKey As Long, _
     ByVal lpSubKey As String) As Long
     
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" _
    (ByVal hKey As Long, _
     ByVal lpSubKey As String) As Long
     
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" _
    (ByVal hKey As Long, _
     ByVal lpValueName As String, _
     ByVal lpReserved As Long, _
     lpType As Long, _
     lpData As Any, _
     lpcbData As Long) As Long
     
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" _
    (ByVal hKey As Long, _
     ByVal lpValueName As String, _
     ByVal Reserved As Long, _
     ByVal dwType As Long, _
     lpData As Any, _
     ByVal cbData As Long) As Long

' Registry error constants
Const ERROR_SUCCESS = 0&
Const ERROR_BADDB = 1009&
Const ERROR_BADKEY = 1010&
Const ERROR_CANTOPEN = 1011&
Const ERROR_CANTREAD = 1012&
Const ERROR_CANTWRITE = 1013&
Const ERROR_REGISTRY_RECOVERED = 1014&
Const ERROR_REGISTRY_CORRUPT = 1015&
Const ERROR_REGISTRY_IO_FAILED = 1016&
Const ERROR_NOT_REGISTRY_FILE = 1017&
Const ERROR_KEY_DELETED = 1018&
Const ERROR_NO_LOG_SPACE = 1019&
Const ERROR_KEY_HAS_CHILDREN = 1020&
Const ERROR_CHILD_MUST_BE_VOLATILE = 1021&
Const ERROR_RXACT_INVALID_STATE = 1369&




''Reg Key Security Options...
Private Const READ_CONTROL = &H20000
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_CREATE_LINK = &H20
Private Const KEY_READ = KEY_QUERY_VALUE + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + READ_CONTROL
Private Const KEY_WRITE = KEY_SET_VALUE + KEY_CREATE_SUB_KEY + READ_CONTROL
Private Const KEY_EXECUTE = KEY_READ
Private Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE _
                            + KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS _
                            + KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL

Private Declare Function RegEnumValue Lib "advapi32" Alias "RegEnumValueA" _
    (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, _
    ByRef lpcbValueName As Long, ByVal lpReserved As Long, ByRef lpType As Long, _
    ByVal lpData As String, ByRef lpcbData As Long) As Long




''Reg Data Types...
Public Const REG_NONE = 0                                          ' No value type
Public Const REG_SZ = 1                                            ' Unicode nul terminated string
Public Const REG_EXPAND_SZ = 2                                     ' Unicode nul terminated string
Public Const REG_BINARY = 3                                        ' Free form binary
Public Const REG_DWORD = 4                                         ' 32-bit number
Public Const REG_DWORD_LITTLE_ENDIAN = 4                           ' 32-bit number (same as REG_DWORD)
Public Const REG_DWORD_BIG_ENDIAN = 5                              ' 32-bit number
Public Const REG_LINK = 6                                          ' Symbolic Link (unicode)
Public Const REG_MULTI_SZ = 7                                      ' Multiple Unicode strings
Public Const REG_RESOURCE_LIST = 8                                 ' Resource list in the resource map
Public Const REG_FULL_RESOURCE_DESCRIPTOR = 9                      ' Resource list in the hardware description
Public Const REG_RESOURCE_REQUIREMENTS_LIST = 10
Sub EnumSubVals(hRootKey, hKey, SubKeys())
' Sample Call:
' EnumSubVals HKEY_CURRENT_USER, "Software\ACD Systems\ACDSee32", x()

ReDim SubKeys(1 To 1)
Dim i As Long, hkRet As Long
Dim sKey As String * 255
Dim iRet As Long

Dim sVal As String * 255


i = 0
iRet = RegOpenKey(hRootKey, hKey, hkRet)
Do
    
    iRet = RegEnumValue(hkRet, i, sKey, Len(sKey), 0, REG_BINARY, sVal, Len(sVal))
    If iRet <> 0 Then Exit Do
    i = i + 1
    ReDim Preserve SubKeys(1 To i)
    SubKeys(i) = Left$(sKey, Len(sKey) - 1)
    
Loop

End Sub


Sub EnumSubKeys(hRootKey, hKey, SubKeys())
ReDim SubKeys(1 To 1)
Dim i As Long, hkRet As Long
Dim sKey As String * 255
Dim iRet As Long

i = 0
iRet = RegOpenKey(hRootKey, hKey, hkRet)
Do
    
    iRet = RegEnumKey(hkRet, i, sKey, Len(sKey))
    If iRet <> 0 Then Exit Do
    i = i + 1
    ReDim Preserve SubKeys(1 To i)
    SubKeys(i) = Left$(sKey, Len(sKey) - 1)
    
Loop

End Sub

Public Sub DeleteValue(hKey_Root As Long, RegistryKey As String, SubVal As String)
Dim lResult As Long
Dim lKeyId As Long


If Len(RegistryKey) = 0 Then
    Exit Sub
End If

If Len(SubVal) = 0 Then
        Exit Sub
End If

' Open the key by attempting to create it. If it
' already exists we get back an ID.
lResult = RegCreateKey(hKey_Root, RegistryKey, lKeyId)
If lResult = 0 Then
    lResult = RegDeleteValue(lKeyId, ByVal SubVal)
End If

End Sub

Public Sub DeleteKey(hKey_Root As Long, SubKey As String)
Dim lResult As Long
Dim lKeyId As Long

If Len(SubKey) = 0 Then
        Exit Sub
End If

    lResult = RegDeleteKey(hKey_Root, ByVal SubKey)


End Sub


Public Sub CreateKey(hKey_Root As Long, RegistryKey As String)
'Sample Call:
'  CreateKey HKEY_LOCAL_MACHINE, "AOCOX\FOO\Z"

Dim lResult As Long
Dim hKey As Long


If Len(RegistryKey) = 0 Then
      Exit Sub
End If

lResult = RegCreateKey(hKey_Root, RegistryKey, hKey)
' RegCloseKey is required only when Opening or Creating a key
lResult = RegCloseKey(hKey)

End Sub


Public Sub SetSubVal(hKey_Root As Long, RegistryKey As String, SubVal As String, KeyValue As String)
'Sample Call: SetSubVal HKEY_CURRENT_USER, "FooKey", "xsub", "xval"
'Use an empty string for SubVal to access the default value of the key
'Use an empty string for KeyValue to clear the contents.

Dim lResult As Long
Dim lKeyId As Long

' Make sure all required properties have been set
If Len(RegistryKey) = 0 Then
        Exit Sub
End If



' Open the key by attempting to create it. If it
' already exists we get back an ID.
lResult = RegCreateKey(hKey_Root, RegistryKey, lKeyId)

If lResult <> ERROR_SUCCESS Then
        Exit Sub
End If

If Len(KeyValue) = 0 Then
    ' No key value, so clear any existing entry
    lResult = RegSetValueEx(lKeyId, _
                SubVal, _
                0&, _
                REG_SZ, _
                0&, _
                0&)
  Else
    ' Set the registry entry to the value
    lResult = RegSetValueEx(lKeyId, _
                SubVal, _
                0&, _
                REG_SZ, _
                ByVal KeyValue, _
                Len(KeyValue) + 1)
End If

End Sub



Public Function GetSubVal(hKey_Root, RegistryKey, SubVal) As String
'Sample Call
'sv$ = GetSubVal(HKEY_CURRENT_USER, "software\adaptec\sessions\selector", "left")
'To get the default value, use an empty string for  SubVal

Dim lResult             As Long
Dim lKeyId              As Long
Dim tKeyValue           As String
Dim lBufferSize         As Long


If Len(RegistryKey) = 0 Then
    ' The key property is not set, so flag an error
     GetSubVal = ""
     Exit Function
End If


lResult = RegOpenKeyEx(hKey_Root, RegistryKey, 0, KEY_ALL_ACCESS, lKeyId)
If lResult <> ERROR_SUCCESS Then
    ' Call failed, can't open the key so exit
    GetSubVal = ""
    Exit Function
End If

' Determine the size of the data in the registry entry
lResult = RegQueryValueEx(lKeyId, SubVal, 0&, REG_SZ, 0&, lBufferSize)
                
If lBufferSize < 2 Then
    ' No data value available
    GetSubVal = ""
    Exit Function
End If

' Allocate the needed space fopr the key data
tKeyValue = String(lBufferSize + 1, " ")

' Get the value of the registry entry
lResult = RegQueryValueEx(lKeyId, SubVal, 0&, REG_SZ, ByVal tKeyValue, lBufferSize)

' Trim the null at the end of the returned value
GetSubVal = Left$(tKeyValue, lBufferSize - 1)

End Function



