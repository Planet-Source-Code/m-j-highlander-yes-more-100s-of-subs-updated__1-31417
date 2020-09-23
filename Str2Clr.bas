Function StringToColor (sColor) As Long
Const BLACK = &H0&
Const RED = &HFF&
Const GREEN = &HFF00&
Const YELLOW = &HFFFF&
Const BLUE = &HFF0000
Const MAGENTA = &HFF00FF
Const CYAN = &HFFFF00
Const WHITE = &HFFFFFF
Const GREY = &HC0C0C0

    Select Case LCase(sColor)
        Case "black", "k"
            StringToColor = BLACK
        Case "red", "r"
            StringToColor = RED
        Case "green", "g"
            StringToColor = GREEN
        Case "yellow", "y"
            StringToColor = YELLOW
        Case "blue", "b"
            StringToColor = BLUE
        Case "magenta", "m"
            StringToColor = MAGENTA
        Case "cyan", "c"
            StringToColor = CYAN
        Case "white", "w"
            StringToColor = WHITE
        Case "grey", "e"
            StringToColor = GREY
    Case Else
            StringToColor = BLACK
    End Select
End Function
