'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@           Smooth Fade Bas By WILDeHACK (WILDeHACK@aol.com)
'@  Modified From Module Received from Daniel Appleman
'@  From his Book "Visual Basic Programmer's Guide to the Window's API"
'@
'@
'@ Use At your Own Risk
'@
'@ Steps to do this
'@ 1) Put a pictureBox on your form ans Call it "faded"
'@ 2) Place that picturebox in the upper-left corner of the screen, touching the sides
'@ 3) In the form_Load event, add this code: TheFormLoad me
'@ 4) In the Form_Resize Event, add this code: ResizeTheForm me
'@ 5) In the Picture_Paint Event, add this:  FillPicture Me
'@ 6) You should be set
'@
'@  Check out |
'@            |
'@            \/


'Mess with this number to determine the number of sections to be contructed
'The greater the number, the smoother the fade
Global Const PALENTRIES = 64





Type POINTAPI  '4 Bytes - Synonymous with LONG
	X As Integer
	y As Integer
End Type

Type SIZEAPI  '4 Bytes - Synonymous with LONG
	X As Integer
	y As Integer
End Type

' ParameterBlock description structure for use with LoadModule
Type PARAMETERBLOCK  '14 Bytes
	wEnvSeg As Integer
	lpCmdLine As Long
	lpCmdShow As Long
	dwReserved As Long
End Type


'  GDI Logical Objects:

'  Pel Array
Type PELARRAY  ' 10 Bytes
	paXCount As Integer
	paYCount As Integer
	paXExt As Integer
	paYExt As Integer
	paRGBs As Integer
End Type

'  Logical Brush (or Pattern)
Type LOGBRUSH     '8 Bytes
	lbStyle As Integer
	lbColor As Long
	lbHatch As Integer
End Type

'  Logical Pen
Type LOGPEN    '10 Bytes
	lopnStyle As Integer
	lopnWidth As POINTAPI
	lopnColor As Long
End Type

Type PALETTEENTRY    '4 Bytes
	peRed As String * 1
	peGreen As String * 1
	peBlue As String * 1
	peFlags As String * 1
End Type

'  Logical Palette
Type LOGPALETTE
	palVersion As Integer
	palNumEntries As Integer
	palPalEntry As String * 252 ' Array length is arbitrary; may be changed
End Type
' Project PalTest

' Module containing global contstants and general purpose
' routines.
Declare Function SetClipboardData Lib "User" (ByVal wFormat As Integer, ByVal hMem As Integer) As Integer
Declare Function CloseClipboard Lib "User" () As Integer
Declare Function OpenClipboard Lib "User" (ByVal hWnd As Integer) As Integer
Declare Sub AnimatePalette Lib "GDI" (ByVal hPalette%, ByVal wStartIndex%, ByVal wNumEntries%, lpPaletteColors As PALETTEENTRY)
Declare Function SendMessageByNum& Lib "User" Alias "SendMessage" (ByVal hWnd%, ByVal wMsg%, ByVal wParam%, ByVal lParam&)
Option Explicit
Global Const PC_RESERVED = &H1
Global Const PC_EXPLICIT = &H2
Global Const PC_NOCOLLAPSE = &H4
Global Const DIB_RGB_COLORS = 0
Global Const DIB_PAL_COLORS = 1
Global Const SYSPAL_STATIC = 1
Global Const SYSPAL_NOSTATIC = 2
Global Const CF_TEXT = 1
Global Const CF_BITMAP = 2
Global Const CF_METAFILEPICT = 3
Global Const CF_SYLK = 4
Global Const CF_DIF = 5
Global Const CF_TIFF = 6
Global Const CF_OEMTEXT = 7
Global Const CF_DIB = 8
Global Const CF_PALETTE = 9
Global Const CF_OWNERDISPLAY = &H80
Global Const CF_DSPTEXT = &H81
Global Const CF_DSPBITMAP = &H82
Global Const CF_DSPMETAFILEPICT = &H83
Global Const CF_PRIVATEFIRST = &H200
Global Const CF_PRIVATELAST = &H2FF


'   This is similar to the LOGPALLETTE defined in
'   APIDECS.BAS, however instead of using a buffer, we
'   create a 64 entry palette for our use.

Type LOGPALETTE64
	palVersion As Integer
	palNumEntries As Integer
	palPalEntry(PALENTRIES) As PALETTEENTRY
End Type

' And create a type safe alias to create palette that handles this structure
Declare Function CreatePalette64% Lib "GDI" Alias "CreatePalette" (lpLogPalette As LOGPALETTE64)


' The six palettes that this program will use are defined here
Global UsePalettes%
Global logPalettes As LOGPALETTE64

' This is a message used within Visual Basic to retrieve
' the handle of a palette
Global Const VBM_GETPALETTE% = &H101C

'   This function creates 6 palettes that are used by
'   the PalTest program
'
Sub CreateAllPalettes ()
    Dim entrynum%
    Dim oldmouseptr%
    Dim X%

    oldmouseptr% = Screen.MousePointer
    Screen.MousePointer = 11
    ' Initialize the logical palette
    
	logPalettes.palVersion = &H300
	logPalettes.palNumEntries = PALENTRIES
    
    For entrynum% = 0 To PALENTRIES - 1
	logPalettes.palPalEntry(entrynum%).peRed = Chr$(0)
	logPalettes.palPalEntry(entrynum%).peGreen = Chr$(0)
	logPalettes.palPalEntry(entrynum%).peBlue = Chr$((255 * entrynum%) / PALENTRIES)
	logPalettes.palPalEntry(entrynum%).peFlags = Chr$(PC_RESERVED)
    Next entrynum%


    ' And create the palettes
  
	UsePalettes = CreatePalette64(logPalettes)
    Screen.MousePointer = oldmouseptr%
End Sub

'   FillPicture draws a spectrum in the specified picture
'   control using the appropriate palette for that control
'
Sub FillPicture (asdf As Form)
    Dim totwidth&, startloc&, endloc&
    Dim pic As control
    Dim X&
    'Dim rc As RECT
    'Dim usebrush%
    'Dim t%

    Set pic = asdf.faded

    totwidth& = pic.ScaleHeight
    For X& = 0 To PALENTRIES - 1
	' We're using long arithmetic for speed. Note the
	' ordering of operations to preserve precesion
	startloc& = (totwidth& * X&) / PALENTRIES
	endloc& = (totwidth& * (X& + 1)) / PALENTRIES
	pic.Line (0, startloc&)-(pic.ScaleWidth, endloc&), GetPalColor(X&), BF
    Next X&

End Sub

'
'   Gets the Long RGB color for a palette entry
'
Function GetPalColor& (entry&)
    Dim res&
    Dim pe As PALETTEENTRY
    LSet pe = logPalettes.palPalEntry(entry&)
    ' We build a long value using this rather awkward
    ' shifting technique.
    ' We actually could save time by performing a raw
    ' memory copy from the pe object into a long variable.
    ' since they are the same format.
    res& = Asc(pe.peRed)
    res& = res& Or (Asc(pe.peGreen) * 256&)
    res& = res& Or (Asc(pe.peBlue) * 256& * 256&)
    GetPalColor& = res&
End Function

Sub resizetheform (pop As Form)
pop.faded.Height = pop.Height
pop.faded.Width = pop.Width
End Sub

Sub TheFormLoad (xyz As Form)
    xyz.faded.Height = xyz.Height
    xyz.faded.Width = xyz.Width
    Dim X%, h%

    CreateAllPalettes
    
    
	h% = OpenClipboard(xyz.hWnd)
	If h% = 0 Then
	    MsgBox "Can't open clipboard"
	    End
	End If
	h% = SetClipboardData(CF_PALETTE, UsePalettes%)
	h% = CloseClipboard()
	xyz.faded.Picture = Clipboard.GetData(CF_PALETTE)
    ' don't own them any more, so don't mess with them.

    

End Sub

