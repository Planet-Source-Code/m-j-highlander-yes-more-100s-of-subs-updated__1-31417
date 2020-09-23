Attribute VB_Name = "FileName_Functions"
Option Explicit
Public Function CombinePath(ByVal Drive As String, ByVal Path As String, ByVal FileName As String, ByVal FileExt As String) As String

If Right(Drive, 1) <> ":" Then
        Drive = Drive & ":"
End If

If Right(Path, 1) <> "\" Then
        Path = Path & "\"
End If
If Left(Path, 1) <> "\" Then
        Path = "\" & Path
End If

If Left(FileExt, 1) <> "." Then
        FileExt = "." & FileExt
End If

CombinePath = Drive & Path & FileName & FileExt

End Function

Public Function ExtractFileNameName(FileName As String) As String
' Extracts the name WITHOUT extension parts of a full file name
' uses ExtractFileName()

Dim pos As Integer
Dim sTemp As String

sTemp = ExtractFileName(FileName)
pos = InStrRev(sTemp, ".")
If pos = 0 Then
    ExtractFileNameName = sTemp
Else
    ExtractFileNameName = Left(sTemp, pos - 1)
End If

End Function


Public Function ExtractFilePath(FileName As String) As String
' Extract the full path (with drive) from a full file name
' VB6 spesific

Dim sTemp As String
Dim pos As Integer

pos = InStr(FileName, "\")

If pos = 0 Then            ' no slashes, so
        sTemp = ""
        pos = InStr(FileName, ":")  'maybe there is a colon? (drive name)
        If pos <> 0 Then
                sTemp = Left(FileName, pos)
        End If

Else
        pos = InStrRev(FileName, "\")
        sTemp = Left(FileName, pos)
        If Right(sTemp, 1) = "\" Then sTemp = Left(sTemp, Len(sTemp) - 1)
End If


ExtractFilePath = sTemp

End Function

Public Function ExtractFilePathPath(FileName As String) As String
' Extracts path WITHOUT drive
' uses ExtractFilePath()

Dim sTemp As String
Dim pos As Integer

sTemp = ExtractFilePath(FileName)

pos = InStr(sTemp, ":")
If pos <> 0 Then               'no colons found
        sTemp = Right(sTemp, Len(sTemp) - pos)
End If

If Left(sTemp, 1) = "\" Then
        sTemp = Right(sTemp, Len(sTemp) - 1)
End If

ExtractFilePathPath = sTemp

End Function

Public Function ExtractDrive(Path As String) As String
' Extract drive name (x:) from a full path

Dim sTemp As String
Dim pos As Integer

pos = InStr(Path, ":")
If pos <> 0 Then
        sTemp = Left(Path, pos)
Else
        sTemp = ""
End If

ExtractDrive = sTemp

End Function

Public Function ExtractFileName(ByVal FileName As String) As String
' Extracts the NAME with EXTENSION parts of a full file name
' VB6 spesific

Dim pos As Integer
Dim sTemp As String

pos = InStrRev(FileName, "\")
If pos = 0 Then      ' no slash found
        sTemp = FileName
        pos = InStr(sTemp, ":")  'maybe there is a colon? (drive name)
        If pos <> 0 Then
                sTemp = Right(sTemp, Len(sTemp) - pos)
        End If
Else                 ' slash found (last slash, if more than one)
        sTemp = Right(FileName, Len(FileName) - pos)
End If

ExtractFileName = sTemp

End Function

Public Function ChangeFileExt(ByVal FileName As String, ByVal NewExtension As String) As String
' uses ExtractFileExt()
' If NewExtension starts with a dot or not it's ok.
' If NewExtension is empty the extension will be removed.
' FileName should not contain path info (actually it's ok
'          unless the path contains a dot and the filename
'          has no extension!

Dim OldExt As String
Dim sTemp As String

OldExt = ExtractFileExt(FileName)

If Left(NewExtension, 1) = "." Then
        'if "dot" exists remove it, var is ByVal so it will not change
        NewExtension = Right(NewExtension, Len(NewExtension) - 1)
End If

If OldExt = "" Then  ' if file has no extension
        NewExtension = "." & NewExtension
End If
sTemp = Left(FileName, Len(FileName) - Len(OldExt)) & NewExtension
If Right(sTemp, 1) = "." Then   'extension was empty, so remove dot
        sTemp = Left(sTemp, Len(sTemp) - 1)
End If

ChangeFileExt = sTemp

End Function

Public Function ExtractFileExt(ByVal FileName As Variant) As String
' VB6 specific, since it uses InStrRev()
' Returns:
' File extension without the "dot" (unlike the equiv Delphi function)
' Empty string if no extension exists.
' FileName should not contain path info (actually it's ok
'          unless the path contains a dot and the filename
'          has no extension!

Dim pos As Integer
Dim PrevPos As Integer

        pos = InStrRev(FileName, ".")  ' locate last dot
        If pos = 0 Then                'no dots found
                ExtractFileExt = ""
        Else
                ExtractFileExt = Right(FileName, Len(FileName) - pos)
        End If

End Function



Public Sub ParsePath(ByVal FullFileName As String, ByRef Drive As String, ByRef Path As String, ByRef FileName As String, ByRef FileExt As String)
' Split a full filename to its components:
' Drive with colon (x:)
' Path without drive and without start or trailing slashes
' FileName without extension
' Extension without dot

Drive = ExtractDrive(FullFileName)
Path = ExtractFilePathPath(FullFileName)
FileName = ExtractFileNameName(FullFileName)
FileExt = ExtractFileExt(FullFileName)

End Sub


Function TrailingBackslash(ByVal Path As String, ByVal SlashState As Boolean) As String
' Add or Remove Trailing Backslash.

Dim sTemp As String

sTemp = Path
Select Case SlashState
        Case True              'add Backslash if not already there
                If Right(sTemp, 1) <> "\" Then
                        sTemp = sTemp & "\"
                End If
        Case False              'remove Backslash if it exists
                If Right(sTemp, 1) = "\" Then
                        sTemp = Left(sTemp, Len(sTemp) - 1)
                End If
End Select
        
TrailingBackslash = sTemp
        
End Function


