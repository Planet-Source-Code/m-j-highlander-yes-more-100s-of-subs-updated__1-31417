Attribute VB_Name = "Module1"
Option Explicit
Type sTag
    href As String
    text As String
End Type

Function DeSpace(sText As String) As String
Dim idx As Integer
Dim sTemp As String
Dim sResult As String
Dim ch As String * 1
Dim InSpaces As Boolean
Dim scount As Integer

sTemp = Trim$(sText)
sResult = ""

scount = 0
For idx = 1 To Len(sTemp)
    ch = Mid(sTemp, idx, 1)
        
    Select Case ch
        Case " "
        scount = scount + 1
        Case Else
        InSpaces = False
        scount = 0
    End Select
    
    
    If scount > 1 Then
        InSpaces = True
    End If
    
    If Not (InSpaces) Then
        sResult = sResult & ch
    End If
    
Next idx
DeSpace = sResult
End Function


Function ExtractText(sLine As String) As String
Dim idx As Integer
Dim ch As String * 1
Dim sTemp As String
Dim InTag As Boolean

For idx = 1 To Len(sLine)
    ch = Mid(sLine, idx, 1)
    
    Select Case ch
        Case "<"
          InTag = True
        Case ">"
          InTag = False
        Case Else
        'do nothing
    End Select
    
    
    If Not (InTag) Then
        Select Case ch
            Case ">"
                ch = ""
            Case Chr(13), Chr(10), Chr(9)
                ch = " "
            Case Else
            'do nothing
         End Select
    sTemp = sTemp + ch
    End If

Next idx

ExtractText = DeSpace(sTemp)

End Function


Function ExtractURL(Tag As String) As String
Dim qpos1 As Integer
Dim qpos2 As Integer
Dim hpos As Integer

hpos = InStr(LCase(Tag), "href")
qpos1 = InStr(hpos + 1, Tag, Chr(34))
qpos2 = InStr(qpos1 + 1, Tag, Chr(34))
ExtractURL = LCase(Mid(Tag, qpos1, qpos2 - qpos1 + 1))

End Function

Function FindTags(SourceText As String, LeftTag As String, RightTag As String, TagArray() As sTag)
Dim pos1 As Long
Dim pos2 As Long
Dim CurrentTag As String
Dim idx As Long
Dim lText As String

lText = LCase(SourceText)
LeftTag = "href="
RightTag = "</a>"

    
pos1 = InStr(lText, LeftTag)
idx = 0
ReDim TagArray(1 To 1)
Do While pos1 <> 0

    pos2 = InStr(pos1 + 1, lText, RightTag)
    CurrentTag = Mid(SourceText, pos1, pos2 - pos1 + 4)
    CurrentTag = "<a " + CurrentTag
    pos1 = InStr(pos2 + 1, lText, LeftTag)
    idx = idx + 1
    ReDim Preserve TagArray(1 To idx)
    'TagArray(idx) = CurrentTag
    TagArray(idx).href = ExtractURL(CurrentTag)
    TagArray(idx).text = ExtractText(CurrentTag)
Loop

End Function

Function LoadFile(FileName As String) As String
'Loads the contents of a file into a string variable

On Error GoTo LoadFile_Error
Dim FF As Integer
Dim FileContents As String

FF = FreeFile
Open FileName For Input As #FF
FileContents = Input(LOF(FF), FF)
Close #FF
LoadFile = FileContents

Exit Function
LoadFile_Error:
    LoadFile = "#ERROR#"
    Exit Function

End Function

