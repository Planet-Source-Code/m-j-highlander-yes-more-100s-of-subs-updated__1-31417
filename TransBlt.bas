Type bitmap
      bmType As Integer
      bmWidth As Integer
      bmHeight As Integer
      bmWidthBytes As Integer

      bmPlanes As String * 1
      bmBitsPixel As String * 1
      bmBits As Long
   End Type
   Declare Function BitBlt Lib "GDI" (ByVal srchDC As Integer, ByVal srcX As Integer, ByVal srcY As Integer, ByVal srcW As Integer, ByVal srcH As Integer, ByVal desthDC As Integer, ByVal destX As Integer, ByVal destY As Integer, ByVal op As Long) As Integer
   Declare Function SetBkColor Lib "GDI" (ByVal hDC As Integer, ByVal cColor As Long) As Long
   Declare Function CreateCompatibleDC Lib "GDI" (ByVal hDC As Integer) As Integer
   Declare Function DeleteDC Lib "GDI" (ByVal hDC As Integer) As Integer
   Declare Function CreateBitmap Lib "GDI" (ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal cbPlanes As Integer, ByVal cbBits As Integer, lpvBits As Any) As Integer
   Declare Function CreateCompatibleBitmap Lib "GDI" (ByVal hDC As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer) As Integer
   Declare Function SelectObject Lib "GDI" (ByVal hDC As Integer, ByVal hObject As Integer) As Integer
   Declare Function DeleteObject Lib "GDI" (ByVal hObject As Integer) As Integer
   Declare Function GetObjectEx Lib "GDI" Alias "GetObject" (ByVal hObject As Integer, ByVal nCount As Integer, bmp As Any) As Integer


   Const SRCCOPY = &HCC0020
   Const SRCAND = &H8800C6
   Const SRCPAINT = &HEE0086
   Const NOTSRCCOPY = &H330008
   

Sub TransparentBlt (dest As PictureBox, ByVal srcBmp As Integer, ByVal destX As Integer, ByVal destY As Integer, ByVal TransColor As Long)


      Const PIXEL = 3
      Dim destScale As Integer
      Dim srcDC As Integer  'source bitmap (color)
      Dim saveDC As Integer 'backup copy of source bitmap
      Dim maskDC As Integer 'mask bitmap (monochrome)
      Dim invDC As Integer  'inverse of mask bitmap (monochrome)
      Dim resultDC As Integer 'combination of source bitmap & background
      Dim bmp As bitmap 'description of the source bitmap
      Dim hResultBmp As Integer 'Bitmap combination of source & background

      Dim hSaveBmp As Integer 'Bitmap stores backup copy of source bitmap
      Dim hMaskBmp As Integer 'Bitmap stores mask (monochrome)
      Dim hInvBmp As Integer  'Bitmap holds inverse of mask (monochrome)
      Dim hPrevBmp As Integer 'Bitmap holds previous bitmap selected in DC
      Dim hSrcPrevBmp As Integer  'Holds previous bitmap in source DC
      Dim hSavePrevBmp As Integer 'Holds previous bitmap in saved DC
      Dim hDestPrevBmp As Integer 'Holds previous bitmap in destination DC
      Dim hMaskPrevBmp As Integer 'Holds previous bitmap in the mask DC
      Dim hInvPrevBmp As Integer  'Holds previous bitmap in inverted mask DC
      Dim OrigColor As Long  'Holds original background color from source DC
      Dim Success As Integer 'Stores result of call to Windows API
      If TypeOf dest Is PictureBox Then 'Ensure objects are picture boxes

        destScale = dest.ScaleMode 'Store ScaleMode to restore later
        dest.ScaleMode = PIXEL 'Set ScaleMode to pixels for Windows GDI
        'Retrieve bitmap to get width (bmp.bmWidth) & height (bmp.bmHeight)
        Success = GetObjectEx(srcBmp, Len(bmp), bmp)
        srcDC = CreateCompatibleDC(dest.hDC)    'Create DC to hold stage
        saveDC = CreateCompatibleDC(dest.hDC)   'Create DC to hold stage
        maskDC = CreateCompatibleDC(dest.hDC)   'Create DC to hold stage
        invDC = CreateCompatibleDC(dest.hDC)    'Create DC to hold stage
        resultDC = CreateCompatibleDC(dest.hDC) 'Create DC to hold stage
        'Create monochrome bitmaps for the mask-related bitmaps:
        hMaskBmp = CreateBitmap(bmp.bmWidth, bmp.bmHeight, 1, 1, ByVal 0&)

        hInvBmp = CreateBitmap(bmp.bmWidth, bmp.bmHeight, 1, 1, ByVal 0&)
        'Create color bitmaps for final result & stored copy of source
        hResultBmp = CreateCompatibleBitmap(dest.hDC, bmp.bmWidth, bmp.bmHeight)
        hSaveBmp = CreateCompatibleBitmap(dest.hDC, bmp.bmWidth, bmp.bmHeight)
        hSrcPrevBmp = SelectObject(srcDC, srcBmp)     'Select bitmap in DC
        hSavePrevBmp = SelectObject(saveDC, hSaveBmp) 'Select bitmap in DC
        hMaskPrevBmp = SelectObject(maskDC, hMaskBmp) 'Select bitmap in DC
        hInvPrevBmp = SelectObject(invDC, hInvBmp)    'Select bitmap in DC
        hDestPrevBmp = SelectObject(resultDC, hResultBmp) 'Select bitmap
        Success = BitBlt(saveDC, 0, 0, bmp.bmWidth, bmp.bmHeight, srcDC, 0, 0, SRCCOPY)'Make backup of source bitmap to restore later

        'Create mask: set background color of source to transparent color.
        OrigColor = SetBkColor(srcDC, TransColor)
        Success = BitBlt(maskDC, 0, 0, bmp.bmWidth, bmp.bmHeight, srcDC, 0, 0, SRCCOPY)
        TransColor = SetBkColor(srcDC, OrigColor)
        'Create inverse of mask to AND w/ source & combine w/ background.
        Success = BitBlt(invDC, 0, 0, bmp.bmWidth, bmp.bmHeight, maskDC, 0, 0, NOTSRCCOPY)
        'Copy background bitmap to result & create final transparent bitmap
        Success = BitBlt(resultDC, 0, 0, bmp.bmWidth, bmp.bmHeight, dest.hDC, destX, destY, SRCCOPY)
        'AND mask bitmap w/ result DC to punch hole in the background by
        'painting black area for non-transparent portion of source bitmap.
        Success = BitBlt(resultDC, 0, 0, bmp.bmWidth, bmp.bmHeight, maskDC, 0, 0, SRCAND)
        'AND inverse mask w/ source bitmap to turn off bits associated
        'with transparent area of source bitmap by making it black.
        Success = BitBlt(srcDC, 0, 0, bmp.bmWidth, bmp.bmHeight, invDC, 0, 0, SRCAND)
        'XOR result w/ source bitmap to make background show through.
        Success = BitBlt(resultDC, 0, 0, bmp.bmWidth, bmp.bmHeight, srcDC, 0, 0, SRCPAINT)
        Success = BitBlt(dest.hDC, destX, destY, bmp.bmWidth, bmp.bmHeight, resultDC, 0, 0, SRCCOPY)'Display transparent bitmap on backgrnd
        Success = BitBlt(srcDC, 0, 0, bmp.bmWidth, bmp.bmHeight, saveDC, 0, 0, SRCCOPY)'Restore backup of bitmap.
        hPrevBmp = SelectObject(srcDC, hSrcPrevBmp) 'Select orig object
        hPrevBmp = SelectObject(saveDC, hSavePrevBmp) 'Select orig object

        hPrevBmp = SelectObject(resultDC, hDestPrevBmp) 'Select orig object
        hPrevBmp = SelectObject(maskDC, hMaskPrevBmp) 'Select orig object
        hPrevBmp = SelectObject(invDC, hInvPrevBmp) 'Select orig object
        Success = DeleteObject(hSaveBmp)   'Deallocate system resources.
        Success = DeleteObject(hMaskBmp)   'Deallocate system resources.
        Success = DeleteObject(hInvBmp)    'Deallocate system resources.
        Success = DeleteObject(hResultBmp) 'Deallocate system resources.
        Success = DeleteDC(srcDC)          'Deallocate system resources.
        Success = DeleteDC(saveDC)         'Deallocate system resources.
        Success = DeleteDC(invDC)          'Deallocate system resources.
        Success = DeleteDC(maskDC)         'Deallocate system resources.
        Success = DeleteDC(resultDC)       'Deallocate system resources.
        dest.ScaleMode = destScale 'Restore ScaleMode of destination.

      End If

End Sub

