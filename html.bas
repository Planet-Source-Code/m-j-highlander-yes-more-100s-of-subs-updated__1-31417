Option Explicit

Global Const ATTR_NORMAL = 0  'Normal files
Global Const ATTR_HIDDEN = 2  'Hidden files
Global Const ATTR_SYSTEM = 4  'System files
Global Const ATTR_DIRECTORY = 16  'Directory

Function Array_To_HTML_Table (TLines() As String, TableAlign As String, TableWidth As String, Border As Integer, CelsPerRow As Integer, Cel_Alignment As String) As String
' Important:
' This function takes PLAIN text as input not HTML!

Dim tmp As String
Dim indx As Integer
Dim inner_index As Integer
Dim CrLf As String
Dim q As String
Dim Table_Header As String
Dim sBorder As String

If Border = True Then sBorder = "BORDER" Else sBorder = ""
Select Case Left(LCase(TableAlign), 1)
    Case "c": TableAlign = "CENTER"
    Case "l": TableAlign = "LEFT"
    Case "r": TableAlign = "RIGHT"
    Case Else: TableAlign = "LEFT"
End Select

Select Case Left(LCase(Cel_Alignment), 1)
    Case "c": Cel_Alignment = "CENTER"
    Case "l": Cel_Alignment = "LEFT"
    Case "r": Cel_Alignment = "RIGHT"
    Case Else: Cel_Alignment = "LEFT"
End Select
Table_Header = "<TABLE " + sBorder + " ALIGN=" + TableAlign + " WIDTH=" + TableWidth + ">"
'MsgBox Table_Header
'CelsPerRow
q = Chr(34)
CrLf = Chr(13) + Chr(10)

Dim Max_Bound As Integer
Max_Bound = ((UBound(TLines) \ CelsPerRow) + (UBound(TLines) Mod CelsPerRow)) * CelsPerRow

ReDim Preserve TLines(1 To Max_Bound)
tmp = Table_Header + CrLf
'For indx = LBound(TLines) To UBound(TLines): MsgBox TLines(indx): Next
Debug.Print LBound(TLines)
Debug.Print UBound(TLines)
'Debug.Print CelsPerRow
'Debug.Print UBound(TLines) \ CelsPerRow
'Dim LastInnerIndex As Integer
'LastInnerIndex = 0

For indx = LBound(TLines) To UBound(TLines) \ CelsPerRow
    tmp = tmp + "<TR ALIGN=" + Cel_Alignment + ">"
        
        For inner_index = indx * CelsPerRow - CelsPerRow + 1 To (indx * CelsPerRow)
            tmp = tmp + "<TD>" + TLines(inner_index) + "</TD>" + CrLf
        Next inner_index
        
    tmp = tmp + "</TR>" + CrLf
Next indx
tmp = tmp + "</TABLE>"
Array_To_HTML_Table = tmp

End Function

Function Body_Colors (TextColor As String, LinkColor As String, ALinkColor As String, VLinkColor As String) As String
Dim sText As String
Dim sLink As String
Dim sALink As String
Dim sVLink As String

If TextColor = "" Then
    sText = ""
Else
    sText = "TEXT=" & TextColor
End If
'''''''''''''''''''''''''''''''''
If LinkColor = "" Then
    sLink = ""
Else
    sLink = "LINK=" & LinkColor
End If
'''''''''''''''''''''''''''''''''
If ALinkColor = "" Then
    sALink = ""
Else
    sALink = "ALINK=" & ALinkColor
End If
'''''''''''''''''''''''''''''''''
If VLinkColor = "" Then
    sVLink = ""
Else
    sVLink = "VLINK=" & VLinkColor
End If

Body_Colors = Trim$(sText & " " & sLink & " " & sALink & " " & sVLink)
End Function

Function HTML_Body (BGColor As String, BGPic As String, BGFixed As Integer, Colors As String) As String
Dim sColor As String
Dim sPic As String
Dim sFixed As String
Dim sColors As String

Dim q As String
Dim CrLf As String

q = Chr(34)
CrLf = Chr(13) + Chr(10)

If BGColor = "" Then sColor = "" Else sColor = "BGCOLOR=" & BGColor
If BGFixed = True Then sFixed = "BGPROPERTIES=FIXED" Else sFixed = ""
If BGPic = "" Then
    sPic = ""
    sFixed = ""
Else
    sPic = "BACKGROUND=" & q & BGPic & q
End If

If Colors = "" Then sColors = "" Else sColors = CrLf + Colors
HTML_Body = Trim$("<BODY " & sColor & " " & sPic & " " & sFixed & " " & sColors) & ">" + CrLf

End Function

Function HTML_Bold (HTML As String) As String

    HTML_Bold = "<B>" + HTML + "</B>"

End Function

Function HTML_Font (HTML As String, Face As String, Size As Integer, Color As String) As String
'//TODO
'  Check if color is a hex value or a color name,
'  valid color names are red,green,blue,cyan,magenta,black,grey,....?
Dim tmp As String
Dim CrLf As String
Dim q As String

CrLf = Chr(13) + Chr(10)
q = Chr(34)

If Color = "" Then Color = "black"
If Face = "" Then Face = "MS Sans Serif"
If Size = 0 Then Size = 3

tmp = "<FONT COLOR=" + Color + " FACE=" + q + Face + q + " SIZE=" + Format$(Size) + ">"
tmp = tmp + CrLf + HTML + CrLf + "</FONT>"

HTML_Font = tmp

End Function

Function HTML_Italic (HTML As String) As String

    HTML_Italic = "<I>" + HTML + "</I>"

End Function

Function HTML_Link (Target As String, Caption As String) As String
Dim tmp As String
Dim q As String * 1
q = ""
tmp = "<A HREF=" + q + Target + q + ">" + Caption + "</A>"

HTML_Link = tmp

End Function

Function HTML_Table (HTMLCells As String, Border As Integer, Alignment As String, TableWidth As String) As String
Dim sBorder As String
Dim sAlign As String

Dim CrLf As String
CrLf = Chr(13) + Chr(10)

If Border <= 0 Then
    sBorder = ""
Else
    sBorder = "BORDER=" & CStr(Border)
End If
Select Case LCase$(Alignment)
    Case "c", "center"
        sAlign = "CENTER"
    Case "l", "left"
        sAlign = "LEFT"
    Case "r", "right"
        sAlign = "ROGHT"
    Case "t", "top"
        sAlign = "TOP"
    End Select

HTML_Table = "<TABLE " + sBorder + " ALIGN=" + sAlign + " WIDTH=" + TableWidth + " >" + CrLf + HTMLCells + CrLf + "</TABLE>" + CrLf

End Function

Function HTML_TableCell (astr As String) As String

HTML_TableCell = "<TD>" + astr + "</TD>"

End Function

Function HTML_TableRow (astr As String) As String

Dim CrLf As String
CrLf = Chr(13) + Chr(10)

HTML_TableRow = "<TR>" + CrLf + astr + CrLf + "</TR>" + CrLf

End Function

Function HTML_Underline (HTML As String) As String

    HTML_Underline = "<U>" + HTML + "</U>"

End Function

Function ReadFile (sFileName As String) As String
On Error GoTo ReadFileError
Dim FF As Integer
Dim TmpStr As String

FF = FreeFile
Open sFileName For Input As #FF
TmpStr = Input$(LOF(FF), FF)
Close #FF
ReadFile = TmpStr
Exit Function

ReadFileError:
    ReadFile = ""
    Exit Function

End Function

Function Text_2_HTML (Text As String) As String
' Important:
' This function takes PLAIN text as input not HTML!

Dim tmp As String
Dim indx As Integer
Dim CrLf As String
Dim q As String


ReDim TLines(1 To 1) As String
q = Chr(34)
CrLf = Chr(13) + Chr(10)
TextToLines Text, TLines()

tmp = ""

For indx = LBound(TLines) To UBound(TLines)
        tmp = tmp + TLines(indx) + "<BR>" + CrLf
Next indx


Text_2_HTML = tmp

End Function

Function Text_2_HTML_List (Text As String, Numbered As Integer) As String
' Important:
' This function takes PLAIN text as input not HTML!

Dim tmp As String
Dim indx As Integer
Dim CrLf As String
Dim q As String

Dim LstO As String, LstC As String


ReDim TLines(1 To 1) As String
q = Chr(34)
CrLf = Chr(13) + Chr(10)
TextToLines Text, TLines()


If Numbered = True Then
    LstO = "<OL>"
    LstC = "</OL>"
Else
    LstO = "<UL>"
    LstC = "</UL>"
End If


tmp = LstO

For indx = LBound(TLines) To UBound(TLines)
        If TLines(indx) <> "" Then
            tmp = tmp + "<LI>" + TLines(indx) + "<BR>" + CrLf
        Else
            tmp = tmp + "<BR>" + CrLf
        End If
Next indx


'If Right(tmp, 6) = "<BR>" + CrLf Then
'    tmp = Left(tmp, Len(tmp) - 6)
'End If


tmp = tmp + LstC

Text_2_HTML_List = tmp

End Function

Function Text_To_HTML_Table (Text As String, TableAlign As String, TableWidth As String, Border As Integer, CelsPerRow As Integer, Cel_Alignment As String) As String
' Important:
' This function takes PLAIN text as input not HTML!

Dim tmp As String
Dim indx As Integer
Dim inner_index As Integer
Dim CrLf As String
Dim q As String
Dim Table_Header As String
Dim sBorder As String

If Border = True Then sBorder = "BORDER" Else sBorder = ""
Select Case Left(LCase(TableAlign), 1)
    Case "c": TableAlign = "CENTER"
    Case "l": TableAlign = "LEFT"
    Case "r": TableAlign = "RIGHT"
    Case Else: TableAlign = "LEFT"
End Select

Select Case Left(LCase(Cel_Alignment), 1)
    Case "c": Cel_Alignment = "CENTER"
    Case "l": Cel_Alignment = "LEFT"
    Case "r": Cel_Alignment = "RIGHT"
    Case Else: Cel_Alignment = "LEFT"
End Select
Table_Header = "<TABLE " + sBorder + " ALIGN=" + TableAlign + " WIDTH=" + TableWidth + ">"
'MsgBox Table_Header
'CelsPerRow
ReDim TLines(1 To 1) As String
q = Chr(34)
CrLf = Chr(13) + Chr(10)
TextToLines Text, TLines()

Dim Max_Bound As Integer
Max_Bound = ((UBound(TLines) \ CelsPerRow) + (UBound(TLines) Mod CelsPerRow)) * CelsPerRow

ReDim Preserve TLines(1 To Max_Bound)
tmp = Table_Header + CrLf
'For indx = LBound(TLines) To UBound(TLines): MsgBox TLines(indx): Next
Debug.Print LBound(TLines)
Debug.Print UBound(TLines)
'Debug.Print CelsPerRow
'Debug.Print UBound(TLines) \ CelsPerRow
'Dim LastInnerIndex As Integer
'LastInnerIndex = 0

For indx = LBound(TLines) To UBound(TLines) \ CelsPerRow
    tmp = tmp + "<TR ALIGN=" + Cel_Alignment + ">"
        
        For inner_index = indx * CelsPerRow - CelsPerRow + 1 To (indx * CelsPerRow)
            tmp = tmp + "<TD>" + TLines(inner_index) + "</TD>" + CrLf
        Next inner_index
        
    tmp = tmp + "</TR>" + CrLf
Next indx
tmp = tmp + "</TABLE>"
Text_To_HTML_Table = tmp

End Function

Sub TextToLines (Text As String, Lines() As String)
'//TODO
'  use temp files in the temp folder...

Dim FF As Integer
Dim index As Integer
FF = FreeFile
Open "c:\~~tmp.tmp" For Output As #FF
Print #FF, Text
Close FF
FF = FreeFile
Open "c:\~~tmp.tmp" For Input As #FF
index = 1
Do While Not EOF(FF)
    ReDim Preserve Lines(1 To index)
    Line Input #FF, Lines(index)
    index = index + 1
Loop
Close #FF
Kill "c:\~~tmp.tmp"
End Sub

Function WriteFile (sFileName As String, sContents As String) As Integer

Const ATTR_ALL_FILES = ATTR_NORMAL Or ATTR_HIDDEN Or ATTR_SYSTEM Or ATTR_DIRECTORY
Dim FF As Integer

On Error GoTo WriteFileError

FF = FreeFile
If Dir(sFileName, ATTR_ALL_FILES) = "" Then
    Open sFileName For Output As #FF
    Print #FF, sContents
    Close #FF
    WriteFile = True
Else       'File Already Exists
    WriteFile = False
End If

Exit Function

WriteFileError:
    WriteFile = False
    Exit Function

End Function

Function WriteHTMLFile (sFileName As String, sContents As String, sTitle As String) As Integer
Dim HTML_Header As String
Dim HTML_Footer As String
Dim HTML As String
Dim iRet As Integer
Dim CrLf As String
CrLf = Chr(13) + Chr(10)

HTML_Header = "<HTML>" + CrLf + "<TITLE>" + CrLf + sTitle + "</TITLE>" + CrLf + "<BODY>"
HTML_Footer = "</BODY>" + CrLf + "</HTML>"
HTML = HTML_Header + sContents + HTML_Footer
iRet = WriteFile(sFileName, HTML)

WriteHTMLFile = iRet

End Function

