Attribute VB_Name = "PRINT1"
'   QuickSmith - Smith Chart Program
'  Copyright (C) <2013-2014>  <Nathan Iyer>
'
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with this program.  If not, see <http://www.gnu.org/licenses/>.
'
' minor changes for 32/64 bit porting J.G.Zvonar 6/14/2014, Version 5.0.0

Option Explicit

Type PALETTEENTRY
    peRed As String * 1
    peGreen As String * 1
    peBlue As String * 1
    peFlags As String * 1
End Type

Type LOGPALETTE
    palVersion As Integer
    palNumEntries As Integer
    palPalEntry(255) As PALETTEENTRY ' Enough for 256 colors
End Type

Type BITMAPINFOHEADER
   biSize As Long
   biWidth As Long
   biHeight As Long
   biPlanes As Integer
   biBitCount As Integer
   biCompression As Long
   biSizeImage As Long
   biXPelsPerMeter As Long
   biYPelsPerMeter As Long
   biClrUsed As Long
   biClrImportant As Long
End Type

Type BITMAPINFO
   bmiHeader As BITMAPINFOHEADER
   bmiColors(255) As PALETTEENTRY 'Enough for 256 colors
End Type

Type Rect1
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Global Const PIXELS = 3
Global Const SRCCOPY = &HCC0020
Global Const BI_RGB = 0
Global Const DIB_RGB_COLORS = 0
Global Const GMEM_MOVEABLE = 2
Global Const RASTERCAPS = 38
Global Const RC_STRETCHDIB = &H2000
Global Const RC_PALETTE = &H100
Global Const PLANES = 14
Global Const BITSPIXEL = 12
Global Const SIZEPALETTE = 104

Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Declare Function GetSystemPaletteEntries Lib "gdi32" (ByVal hDC As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
Declare Function CreatePalette Lib "gdi32" (lpLogPalette As LOGPALETTE) As Long
Declare Function SelectObject Lib "user32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwROP As Long) As Long
Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Declare Function StretchDIBits Lib "gdi32" (ByVal hDC As Integer, ByVal DestX As Integer, ByVal DestY As Integer, ByVal wDestWidth As Integer, ByVal wDestHeight As Integer, ByVal SrcX As Integer, ByVal SrcY As Integer, ByVal wSrcWidth As Integer, ByVal wSrcHeight As Integer, ByVal lpBits As Long, BitsInfo As BITMAPINFO, ByVal wUsage As Integer, ByVal dwROP As Long) As Integer
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Declare Function GlobalAlloc Lib "Kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Declare Function GlobalLock Lib "Kernel32" (ByVal hMem As Long) As Long
Declare Function GlobalUnlock Lib "Kernel32" (ByVal hMem As Long) As Long
Declare Function GlobalFree Lib "Kernel32" (ByVal hMem As Long) As Long
Declare Function SelectPalette Lib "gdi32" (ByVal hDC As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
Declare Function RealizePalette Lib "gdi32" (ByVal hDC As Long) As Long
Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long

'Error Constants: ' Device does not support StretchDIBits.
Global Const ERR_DEVSTRETCHDIB = 11105 ' Palette is not 256-color palette.
Global Const ERR_PALSIZE = 11106 ' Unable to create device context.
Global Const ERR_CREATEMEMDC = 11107 ' Unable to create bitmap.
Global Const ERR_CREATEBMP = 11108 ' Unable to retrieve system palette.
Global Const ERR_GETPALETTE = 11109 ' Unable to create a new palette.
Global Const ERR_CREATEPAL = 11120 ' Unable to copy bitmap to memory.
Global Const ERR_BITBLT = 11110 ' Unable to allocate memory for DIB bits.
Global Const ERR_BITMEM = 11111 ' Unable to lock DIB bits memory.
Global Const ERR_LOCKBITMEM = 11112 ' Unable to get DIB bits
Global Const ERR_GETDIB = 11113 ' Unable to copy bitmap to destination.
Global Const ERR_STRETCHDIB = 11114 ' Unable to unlock DIB bits memory.
Global Const ERR_UNLOCKMEM = 11115 ' Unable to free DIB bits memory.
Global Const ERR_FREEMEM = 11116 ' Unable to select palette.
Global Const ERR_SELPAL = 11117 ' Unable to delete palette.
Global Const ERR_DELPAL = 11121 ' Unable to delete bitmap.
Global Const ERR_DELBMP = 11118 ' Unable to select palette.
Global Const ERR_DELMEMDC = 11119


Sub PrintClient256(frmSrc As Form)
'--------------------------------------------------------------------------
' PrintForm256: ' - Prints the entire form.
' - Renders 256-color bitmaps as they appear on the form.
' - Adjusts output to the size and orientation of the printer's page.
' - Calls StretchFormToDC to copy the contents of the form to the printer.
' - Starts and ends a print job. '
' frmSrc: ' - The form object to print. '
' Errors: ' - A message box is displayed for StrechFormToDC errors.
' - Otherwise, ther is no error trapping. '
'--------------------------------------------------------------------------



   Dim hDCWindow As Integer
   Dim WindowWidth As Integer
   Dim WindowHeight As Integer
   Dim WindowRatio As Double
   Dim PrinterWindowWidth As Integer
   Dim PrinterWindowHeight As Integer
   Dim PrinterRatio As Double
   Dim R
   Dim OldScaleMode As Integer

   Screen.MousePointer = 11 ' Hourglass

   ' Setup form.
   OldScaleMode = frmSrc.ScaleMode  ' Save old ScaleMode
   frmSrc.ScaleMode = PIXELS ' All dimensions must be in pixels.
   hDCWindow = frmSrc.hDC    ' hDC of client area
   WindowWidth = frmSrc.ScaleWidth
   WindowHeight = frmSrc.ScaleHeight
   ' The following must be entered on a single line:
   WindowRatio = (WindowWidth * Screen.TwipsPerPixelX) / (WindowHeight * Screen.TwipsPerPixelY)

   ' Setup printer.
   Printer.ScaleMode = PIXELS
   Printer.Print ""; ' Start print job; initialize printer object.
   ' The following must be entered on a single line:
   PrinterRatio = (Printer.ScaleWidth * Printer.TwipsPerPixelX) / (Printer.ScaleHeight * Printer.TwipsPerPixelY)

   ' Scale the output to the page size.
   If WindowRatio >= PrinterRatio Then
      PrinterWindowWidth = Printer.ScaleWidth
      ' The following must be entered on a single line:
      PrinterWindowHeight = (PrinterWindowWidth * Printer.TwipsPerPixelX) / (WindowRatio * Printer.TwipsPerPixelY)
   Else
      PrinterWindowHeight = Printer.ScaleHeight
      ' The following must be entered on a single line:
      PrinterWindowWidth = (PrinterWindowHeight * Printer.TwipsPerPixelY * WindowRatio) / Printer.TwipsPerPixelX
   End If

   ' Print the client area.
   On Error Resume Next
   ' The following must be entered on a single line:
   Call StretchFormToDC(CInt(Printer.hDC), 0, 0, PrinterWindowWidth, PrinterWindowHeight, hDCWindow, 0, 0, WindowWidth, WindowHeight)
   If Err Then
      MsgBox Err & ": Error Printing Client Area"
      ' Predefined error codes are commented in the general declarations.
   End If
   On Error GoTo 0

   ' End the print job.
   Printer.EndDoc
   frmSrc.ScaleMode = OldScaleMode ' Resotre OldScaleMode
   Screen.MousePointer = 0 ' Default

End Sub

Sub PrintForm256(frmSrc As Form)
'--------------------------------------------------------------------------
' PrintClient256: ' - Prints the client area of a form passed to it.
' - Renders 256-color bitmaps as they appear on the form.
' - Adjusts output to the size and orientation of the printer's page.
' - Calls StretchFormToDC to copy the contents of the form to the printer.
' - Starts and ends a print job. ' ' frmSrc:
' - The form object to print '
'Errors ' - Displays a message box for StrechFormToDC errors.
' - Otherwise, there is no error trapping.
' '--------------------------------------------------------------------------

   Dim RectWindow As RECT  'ZZ from Rect1
   Dim hDCWindow As Integer
   Dim WindowWidth As Integer
   Dim WindowHeight As Integer
   Dim WindowRatio As Double
   Dim PrinterWindowWidth As Integer
   Dim PrinterWindowHeight As Integer
   Dim PrinterRatio As Double
   Dim R
   
   Screen.MousePointer = 11 ' Hourglass

   ' Setup form.
   hDCWindow = GetWindowDC(frmSrc.hwnd) ' hDC of form, including borders
   R = GetWindowRect(frmSrc.hwnd, RectWindow) 'ZZ changed from Rect1 to RECT
 
   WindowWidth = Abs(RectWindow.rRight - RectWindow.rLeft)
   WindowHeight = Abs(RectWindow.rBottom - RectWindow.rTop)
   WindowRatio = (WindowWidth * Screen.TwipsPerPixelX) / (WindowHeight * Screen.TwipsPerPixelY)

   ' Setup printer.
   Printer.ScaleMode = PIXELS
   Printer.Print ""; ' Start print job; initialize printer object.
   ' The following must be entered on a single line:
   PrinterRatio = (Printer.ScaleWidth * Printer.TwipsPerPixelX) / (Printer.ScaleHeight * Printer.TwipsPerPixelY)
   ' Scale the output to the page size.
   If WindowRatio >= PrinterRatio Then
   PrinterWindowWidth = Printer.ScaleWidth
   PrinterWindowHeight = (PrinterWindowWidth * Printer.TwipsPerPixelX) / (WindowRatio * Printer.TwipsPerPixelY)
   Else
      PrinterWindowHeight = Printer.ScaleHeight
      ' The following must be entered on a single line:
      PrinterWindowWidth = (PrinterWindowHeight * Printer.TwipsPerPixelY * WindowRatio) / Printer.TwipsPerPixelX
   End If

   ' Print the form.
   On Error Resume Next
   ' The following must be entered on a single line:
   Call StretchFormToDC(CInt(Printer.hDC), 0, 0, PrinterWindowWidth, PrinterWindowHeight, hDCWindow, 0, 0, WindowWidth, WindowHeight)
   If Err Then
      MsgBox Err & ": Error Printing Form"
      ' Predefined error codes are commented in the general declarations.
   End If
   On Error GoTo 0

   ' Clean up.
   R = ReleaseDC(frmSrc.hwnd, hDCWindow) ' Free DC.

   ' End print job.
   Printer.EndDoc

   Screen.MousePointer = 0 ' Default

 End Sub

Sub StretchFormToDC(hDCDest As Integer, LeftDest, TopDest, WidthDest, HeightDest, hDCSrc As Integer, LeftSrc, TopSrc, WidthSrc, HeightSrc)

'--------------------------------------------------------------------------
' StretchFormToDC ' - Stretches a specified portion of a form to a device context.
' - Works with 256 colors.
' - Works on PostScript and PCL printers (driver must support 'StretchDIBits).
' - Allows you to output to other device contexts ' ' hDCDest: ' - Destination device context.
' - ScaleMode of device context must be pixels.
' - If using Printer object, the printer should be initialized. This can
'    be accomplished with Printer.Print ""; or any other printing.
' ' LeftDest, TopDest, WidthDest, HeightDest:
' - Describe the location and size of the image on the printer in pixels.
' ' hDCSrc: ' - The source device context; should be from a form.
' ' LeftSrc, TopSrc, WidthSrc, HeightSrc: ' - Describe the location and size of the source image in pixels.
' ' Errors: ' - Errors with a predefined code if necessary.
' '--------------------------------------------------------------------------
' The following must be entered on a single line:
   Dim BMI As BITMAPINFO
   Dim hMem As Integer
   Dim lpBits As Long
   Dim R As Integer
   Dim hDCMemory As Integer
   Dim hBmp As Integer
   Dim hBmpPrev As Integer
   Dim hPal As Integer
   Dim hPalPrev As Integer
   Dim RasterCapsDest As Integer
   Dim RasterCapsSrc As Integer
   Dim HasPaletteSrc As Integer
   Dim BitsPixelSrc As Integer
   Dim PlanesSrc As Integer
   Dim PaletteSizeSrc As Integer
   Dim LogPal As LOGPALETTE

   ' Set error trap.
   On Error GoTo SFTDC_ERRORS:

   ' Check that destination supports StretchDIBits.
   RasterCapsDest = GetDeviceCaps(hDCDest, RASTERCAPS)
   If (RasterCapsDest And RC_STRETCHDIB) <> RC_STRETCHDIB Then
      Error ERR_DEVSTRETCHDIB
   End If

   ' Get properties of source device context.
   RasterCapsSrc = GetDeviceCaps(hDCSrc, RASTERCAPS)
   HasPaletteSrc = RasterCapsSrc And RC_PALETTE
   BitsPixelSrc = GetDeviceCaps(hDCSrc, BITSPIXEL)
   PlanesSrc = GetDeviceCaps(hDCSrc, PLANES)
   PaletteSizeSrc = GetDeviceCaps(hDCSrc, SIZEPALETTE)

   ' Limit function use to 256-color palettes.
   If HasPaletteSrc And (PaletteSizeSrc <> 256) Then Error ERR_PALSIZE

   ' Copy source to a bitmap in memory.
   hDCMemory = CreateCompatibleDC(hDCSrc)
   If hDCMemory = 0 Then Error ERR_CREATEMEMDC
   hBmp = CreateCompatibleBitmap(hDCSrc, WidthSrc, HeightSrc)
   If hBmp = 0 Then Error ERR_CREATEBMP
   hBmpPrev = SelectObject(hDCMemory, hBmp)
   ' Create a copy of the system palette and realize it if necessary.
   If HasPaletteSrc Then
      LogPal.palVersion = &H300
      LogPal.palNumEntries = 256
      R = GetSystemPaletteEntries(hDCSrc, 0, 256, LogPal.palPalEntry(0))
      If R = 0 Then Error ERR_GETPALETTE
      hPal = CreatePalette(LogPal)
      If hPal = 0 Then Error ERR_CREATEPAL
      ' Select the palette into the destination and realize it.
      hPalPrev = SelectPalette(hDCMemory, hPal, 0)
      R = RealizePalette(hDCMemory)
   End If
   ' Copy the bitmap to the memory-device context.
   ' The following must be entered on a single line:
   R = BitBlt(hDCMemory, 0, 0, WidthSrc, HeightSrc, hDCSrc, LeftSrc, TopSrc, SRCCOPY)
   If R = 0 Then Error ERR_BITBLT
   hBmp = SelectObject(hDCMemory, hBmpPrev)

   ' Fill in necessary parts of bitmap info.
   BMI.bmiHeader.biSize = 40
   BMI.bmiHeader.biWidth = WidthSrc
   BMI.bmiHeader.biHeight = HeightSrc
   BMI.bmiHeader.biPlanes = 1
   If BitsPixelSrc * PlanesSrc >= 16 Then
      ' >= 16-bit True color may require too much memory so
      ' limit to 256-color DIB.
      ' You can get rid of this exception and the routine
      ' should copy 16-bit color bitmaps.
      BMI.bmiHeader.biBitCount = 8 ' 8 bits = 256 colors
   Else
      BMI.bmiHeader.biBitCount = BitsPixelSrc * PlanesSrc
   End If
   BMI.bmiHeader.biCompression = BI_RGB

   ' Allocate memory for bitmap bits.
   ' The following must be entered on a single line:
   hMem = GlobalAlloc(GMEM_MOVEABLE, (CLng(WidthSrc * BMI.bmiHeader.biBitCount + 31) \ 32) * 4 * HeightSrc)

   If hMem = 0 Then Error ERR_BITMEM
   lpBits = GlobalLock(hMem)

   ' Get the bits and color information from the bitmap.
   ' The following must be entered on a single line:
   R = GetDIBits(hDCMemory, hBmp, 0, HeightSrc, lpBits, BMI, DIB_RGB_COLORS)
   If R = 0 Then Error ERR_GETDIB

   ' Stretch the device-independent bitmap to the printer.
   ' The following must be entered on a single line:
   R = StretchDIBits(hDCDest, LeftDest, TopDest, WidthDest, HeightDest, 0, 0, WidthSrc, HeightSrc, lpBits, BMI, DIB_RGB_COLORS, SRCCOPY)
   If R = 0 Then Error ERR_STRETCHDIB

   ' Free up memory used for bitmap bits.
   R = GlobalUnlock(hMem)
   If R <> 0 Then Error ERR_UNLOCKMEM
   R = GlobalFree(hMem)
   If R <> 0 Then Error ERR_FREEMEM

   ' Select the default palette back if necessary.
   If HasPaletteSrc Then
      R = SelectPalette(hDCMemory, hPalPrev, 0)
      If R = 0 Then Error ERR_SELPAL
      R = DeleteObject(hPal)
      If R = 0 Then Error ERR_DELPAL
   End If

   ' Delete created objects.
   R = DeleteObject(hBmp)
   If R = 0 Then Error ERR_DELBMP
   R = DeleteDC(hDCMemory)
   If R = 0 Then Error ERR_DELMEMDC

   On Error GoTo 0
Exit Sub

' Clean up predefined errors if necessary.
SFTDC_ERRORS:

   Select Case Err
      Case ERR_CREATEBMP
         R = DeleteDC(hDCMemory)
         Error Err
      Case ERR_GETPALETTE, ERR_CREATEPAL
         hBmp = SelectObject(hDCMemory, hBmpPrev)
         R = DeleteObject(hBmp)
         R = DeleteDC(hDCMemory)
         Error Err
      Case ERR_BITBLT
         If HasPaletteSrc Then
            R = SelectPalette(hDCMemory, hPalPrev, 0)
            R = DeleteObject(hPal)
         End If
         hBmp = SelectObject(hDCMemory, hBmpPrev)
         R = DeleteObject(hBmp)
         R = DeleteDC(hDCMemory)
         Error Err
      Case ERR_BITMEM
         If HasPaletteSrc Then
            R = SelectPalette(hDCMemory, hPalPrev, 0)
            R = DeleteObject(hPal)
         End If
         R = DeleteObject(hBmp)
         R = DeleteDC(hDCMemory)
         Error Err
      Case ERR_GETDIB, ERR_STRETCHDIB
         R = GlobalUnlock(hMem)
         R = GlobalFree(hMem)
         If HasPaletteSrc Then
            R = SelectPalette(hDCMemory, hPalPrev, 0)
            R = DeleteObject(hPal)
         End If
         R = DeleteObject(hBmp)
         R = DeleteDC(hDCMemory)
         Error Err
      Case ERR_UNLOCKMEM, ERR_FREEMEM
         If HasPaletteSrc Then
            R = SelectPalette(hDCMemory, hPalPrev, 0)
            R = DeleteObject(hPal)
         End If
         R = DeleteObject(hBmp)
         R = DeleteDC(hDCMemory)
         Error Err
      Case ERR_SELPAL, ERR_DELPAL
         R = DeleteObject(hBmp)
         R = DeleteDC(hDCMemory)
         Error Err
      Case ERR_DELBMP
         R = DeleteDC(hDCMemory)
         Error Err
      Case Else
         Error Err
   End Select
   Error Err

End Sub

