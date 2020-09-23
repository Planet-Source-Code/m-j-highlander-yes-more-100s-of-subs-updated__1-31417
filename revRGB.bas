Function RevRGB(red, green, blue)
'  In the RGB function the low-order byte contains the value for red,
'  the middle byte contains the value for green,
'  and the high-order byte contains the value for blue. 
'  For applications that require the byte order to be reversed,
'  the following function will provide the same information with
'  the bytes reversed: 

    RevRGB= CLng(blue + (green * 256) + (red * 65536))
End Function
