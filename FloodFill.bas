Declare Function ExtFloodFill Lib "GDI" (ByVal hDC%, ByVal i%, ByVal i%, ByVal w&, ByVal i%) As Integer
Global Const FLOOD_FILL_BORDER = 0  ' Fill until crColor& color encountered
Global Const FLOOD_FILL_SURFACE = 1 ' Fill surface until crColor& color not encountered

Sub Paint (pic As Control, x As Single, y As Single)
'Fill While the color at X,Y  is encountered

Dim PointColor As Long
Dim wFillType As Integer
Dim Result As Integer

wFillType = FLOOD_FILL_SURFACE
pic.FillStyle = 0  'Solid
PointColor = pic.Point(x, y)

Result = ExtFloodFill(pic.hDC, x, y, PointColor, wFillType)

End Sub

Sub PaintUntil (pic As Control, x As Single, y As Single, BorderColor As Long)
'Fill Until BorderColor is encountered

Dim wFillType As Integer
Dim Result As Integer
wFillType = FLOOD_FILL_BORDER
pic.FillStyle = 0  'Solid

Result = ExtFloodFill(pic.hDC, x, y, BorderColor, wFillType)

End Sub

