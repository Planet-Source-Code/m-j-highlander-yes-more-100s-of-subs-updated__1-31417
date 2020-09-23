Attribute VB_Name = "Module2"
Public z() As sTag
Public TextColor As String
Public LinkColor As String
Public VisitedLinkColor As String
Public BackgroundColor As String

Public TextFontSize As String
Public TextFontName As String

Sub Main()
'Initilize Global Vars
TextFontName = "Tahoma,Arial"
TextFontSize = "2"
TextColor = "black"
LinkColor = "blue"

BackgroundColor = "#FFFFD0" '& Hex$(frmLayoutOptions.picBackground.BackColor)

'Load Main Form
frmMain.Show
End Sub


