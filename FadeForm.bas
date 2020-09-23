Type RECT
    Left As Integer
    Top As Integer
    Right As Integer
    Bottom As Integer
End Type

Declare Function CreateSolidBrush Lib "GDI" (ByVal crColor As Long) As Integer
Declare Function FillRect Lib "User" (ByVal hDC As Integer, lpRect As RECT, ByVal hBrush As Integer) As Integer
Declare Function DeleteObject Lib "GDI" (ByVal hObject As Integer) As Integer

Sub FadeForm (frm As Form, Color As String)
'  add this sub to a form's Activate event
'  You can use  FadeForm Me ,"color"
'  Where color can be : red,green, or blue (you can use r,g,and b)
'  Any other value or an empty string will cause a gray background
    
    Dim iRed%, iGreen%, iBlue%
    Dim hBrush%
    Dim iHeight%, iStepInterval%, iInterval%, iResult%, iOldScaleMode%
    Dim FillArea As RECT
    
    frm.AutoRedraw = True
    iOldScaleMode = frm.ScaleMode
    frm.ScaleMode = 3  'Pixel
    iHeight = frm.ScaleHeight
    ' Divide the form into 63 regions
    iStepInterval = iHeight \ 63
    FillArea.Left = 0
    FillArea.Right = frm.ScaleWidth
    FillArea.Top = 0
    FillArea.Bottom = iStepInterval
 
    Select Case LCase(Color)
    Case "red", "r"
    iRed = 255: iGreen = 0: iBlue = 0
    Case "green", "g"
    iRed = 0: iGreen = 255: iBlue = 0
    Case "blue", "b"
    iRed = 0: iGreen = 0: iBlue = 255
    Case Else
    iRed = 255: iGreen = 255: iBlue = 255
    End Select
    For iInterval = 1 To 63
        hBrush = CreateSolidBrush(RGB(iRed, iGreen, iBlue))
        iResult = FillRect(frm.hDC, FillArea, hBrush)
        iResult = DeleteObject(hBrush)
        iBlue = Abs(iBlue - 4)
        iGreen = Abs(iGreen - 4)
        iRed = Abs(iRed - 4)
        FillArea.Top = FillArea.Bottom
        FillArea.Bottom = FillArea.Bottom + iStepInterval
    Next
 
    ' Fill the remainder of the form with black
    FillArea.Bottom = FillArea.Bottom + 63
    hBrush = CreateSolidBrush(RGB(0, 0, 0))
    iResult = FillRect(frm.hDC, FillArea, hBrush)
    iResult = DeleteObject(hBrush)
    frm.ScaleMode = iOldScaleMode
    frm.Refresh
End Sub

