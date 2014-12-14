Attribute VB_Name = "SC0"
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
Global arrayR(1010)         As Single  ' array to store sweep data
Global arrayQ(1010)         As Single
Global qstep                As Single  ' step size for Q Circles
Global vswrstep             As Single  ' step size for vswr circle
Global FreqStop!                    ' start freq for sweep -open data,last freq
Global FreqStart!                       ' stop freq for sweep
Global FreqPoints%                      ' number of frequency points
Global WindowVertical As Integer        ' window vertical status flag
Global WindowHorizontal As Integer      ' window horizontal status flag
Global CRLF               As String * 2
Global ArrayNum As Integer  ' used by editor to conunt the number of opened files
Global TOOL6WASCHECKED As Integer ' to remember the toolbars( form6) history
Global TOOL4WASCHECKED As Integer ' to remember the toolbars (form4) history
Global response             As Integer
Global EdopenFlag As Integer
Global PltopenFlag As Integer
Global HelpFilePath As String
Global CurrentDir As String
Global scrollvalue1 As Integer
Global scrollvalue2 As Integer
Global SweepopenFlag As Integer

Global Const WM_MDITILE = &H226
Global Const WM_MDICASCADE = &H227
Global Const WM_MDIICONARRANGE = &H228
' wParam values for WM_MDITILE and WM_MDICASCADE messages.
Global Const MDITILE_VERTICAL = &H0
Global Const MDITILE_HORIZONTAL = &H1
Global Const MDITILE_SKIPDISABLED = &H2 'Requires Win 3.1
Global Const OFN_READONLY = &H1
Global Const CDERR_CANCEL = &H7FF3
Global Const HELP_CONTEXT = &H1
Global Const HELP_INDEX = &H3
Global Const HELP_QUIT = &H2
Global Const HELP_HELPONHELP = &H4
Type RECT
        rLeft As Integer
        rTop As Integer
        rRight As Integer
        rBottom As Integer
End Type
Type MEMORYSTATUS
       dwLength As Long
       dwMemoryLoad As Long
       dwTotalPhys As Long
       dwAvailPhys As Long
       dwTotalPageFile As Long
       dwAvailPageFile As Long
       dwTotalVirtual As Long
       dwAvailVirtual As Long
End Type
Global Mem As MEMORYSTATUS

Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function WinHelp Lib "user32" Alias "WinHelpA" (ByVal hwnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Long) As Long
Declare Function GetSystemDirectory Lib "Kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetFreeSpace Lib "Kernel32" (ByVal wFlags) As Long
Declare Function GetWinFlags Lib "Kernel32" () As Long
Declare Function ScrollWindow Lib "user32" (ByVal hwnd As Long, ByVal XAmount As Long, ByVal YAmount As Long, lpRect As RECT, lpClipRect As RECT) As Long
Declare Sub UpdateWindow Lib "User" (ByVal hwnd As Integer)
Declare Sub GlobalMemoryStatus Lib "Kernel32" (lpBuffer As MEMORYSTATUS)


Sub checkdate()
Dim Filename, Setupdate
Dim DateDifference As Integer
Const Usageperiod = 30 ' set here the usage days from the setup
Dim sysDir_file As String
Dim recorddate As String
On Error GoTo endMe

sysDir_file = GetWindowsSysDir() & "syms.386"
If (FileExists(sysDir_file)) Then   ' if file exists
  Open sysDir_file For Input As #1
  Input #1, recorddate
  Close #1
  'MsgBox recorddate
  DateDifference = DateDiff("d", Now, recorddate)
  If Abs(DateDifference) >= Usageperiod Then GoTo endMe

Else ' file doesnot exist comes here during first execution
  Open sysDir_file For Output As #1
  Print #1, Date$
  Close #1
  SetAttr sysDir_file, 2 ' create hidden file
  Beep
End If

Exit Sub
endMe:
MsgBox "Invalid Executable File", , "QuickSmith Error"
End
End Sub

Function FileExists(Path$) As Integer
    Dim X As Integer
    X = FreeFile

    On Error Resume Next
    Open Path$ For Input As X
    If Err = 0 Then
        FileExists = True
    Else
        FileExists = False
    End If
    Close X

End Function

Function GetWindowsSysDir() As String
    Dim temp$
    Dim X
    temp$ = String$(145, 0)                 ' Size Buffer
    X = GetSystemDirectory(temp$, 145)      ' Make API Call
    temp$ = Left$(temp$, X)                 ' Trim Buffer

    If Right$(temp$, 1) <> "\" Then         ' Add \ if necessary
        GetWindowsSysDir$ = temp$ + "\"
    Else
        GetWindowsSysDir$ = temp$
    End If
End Function

Function GToZI!(R As Single, Q As Single)
On Error Resume Next
Dim Z1i!
Z1i! = (2 * R * Sin(pi * Q / 180)) / (1 + (R * R) - 2 * R * Cos(pi * Q / 180))
GToZI! = Z1i! * Z0


End Function

Function GToZR!(R As Single, Q As Single)
Dim Z1r!
On Error Resume Next
Z1r! = (1 - (R * R)) / (1 + (R * R) - (2 * R * Cos(Q * pi / 180)))
GToZR! = Z1r! * Z0
End Function

Function Help_File_In_Path()
Dim Path As String, CurrentDir As String
Dim Semicolon, found
    On Error Resume Next
    CurrentDir = App.Path
    If Right$(CurrentDir, 1) <> "\" Then CurrentDir = CurrentDir + "\"
    found = Dir$(CurrentDir + "QS.HLP") <> ""
    If Not found Then
        Path = Environ$("PATH")
        If Path <> "" Then
            If Right$(Path, 1) <> ";" Then Path = Path + ";"
            Semicolon = InStr(Path, ";")
            Do
                CurrentDir = Left$(Path, Semicolon - 1)
                If Right$(CurrentDir, 1) <> "\" Then CurrentDir = CurrentDir + "\"
                found = Dir$(CurrentDir + "QS.HLP") <> ""
                Path = Right$(Path, Len(Path) - Semicolon)
                Semicolon = InStr(Path, ";")
            Loop While ((Semicolon <> 0) And Not found)
        End If
    End If
    If found Then
        HelpFilePath = CurrentDir + "QS.HLP"
        App.HelpFile = CurrentDir + "QS.HLP"
    End If
    Help_File_In_Path = found
    
    On Error GoTo 0

End Function

 Sub mdiArrange(method%)

'Use the MDI Arrange methods rather than VB's because it can
  'ignore disabled (hidden) children.
    Dim Ret%
    Select Case method
      Case WM_MDICASCADE
    Ret = SendMessage(GetWindow(MDIForm1.hwnd, 5), WM_MDICASCADE, MDITILE_SKIPDISABLED, 0&)
      Case MDITILE_HORIZONTAL
    Ret = SendMessage(GetWindow(MDIForm1.hwnd, 5), WM_MDITILE, MDITILE_HORIZONTAL Or MDITILE_SKIPDISABLED, 0&)
      Case MDITILE_VERTICAL
    Ret = SendMessage(GetWindow(MDIForm1.hwnd, 5), WM_MDITILE, MDITILE_VERTICAL Or MDITILE_SKIPDISABLED, 0&)
      Case WM_MDIICONARRANGE
    Ret = SendMessage(GetWindow(MDIForm1.hwnd, 5), WM_MDIICONARRANGE, MDITILE_SKIPDISABLED, 0&)
    End Select
    
End Sub

Sub ScrollMe(CscrollMe As Form, NumHorzPixels As Integer, NumVertPixels As Integer)

Dim lprcScroll As RECT
Dim OldScaleMode, Result

' NumVertPixels' +/- the number of pixels to scroll Windows up/down.
' NumHorzPixels' +/- the number of pixels to scroll Windows left/right


OldScaleMode = CscrollMe.ScaleMode     ' Save it so we can set it back
CscrollMe.ScaleMode = 3 'Pixels

' Scroll area is the entire visible Window
lprcScroll.rTop = 0
lprcScroll.rLeft = 0
lprcScroll.rBottom = CscrollMe.ScaleHeight - CscrollMe.HScroll1.Height
lprcScroll.rRight = CscrollMe.ScaleWidth

' Do Scroll
Result = ScrollWindow(CscrollMe.hwnd, NumHorzPixels, NumVertPixels, lprcScroll, lprcScroll)
'Result = ScrollDC(CscrollMe.hDC, NumHorzPixels, NumVertPixels, lprcScroll, lprcScroll, 0, lprcScroll)

CscrollMe.ScaleMode = OldScaleMode ' restore the setting.

End Sub

Sub Yin()
Dim DRP#
DRP# = ZR(13) ^ 2 + ZI(13) ^ 2
On Error GoTo YinError
form1.Label1.Caption = Format(ZR(13) * 1000 / DRP#, "###0.0000") ' reflection coeff.
form1.Label2.Caption = Format(-ZI(13) * 1000 / DRP#, "###0.0000")

form1.Label5.Caption = "Y-In"
form1.Label6.Caption = "mil.S."
Exit Sub
YinError:
If (Err = 11) Then
response% = MsgBox("'A Division By Zero' Error Was Trapped", 49, "Error")
If response = 1 Then Resume Else Exit Sub
End If
End Sub

Function ZtoGammaA(ZXR As Double, ZXI As Double) As Double
  Dim ZRN, ZIN, a1, a2, D, SR, SI, SM, SQ As Double
   
   ZRN = (ZXR / Z0)
   ZIN = (ZXI / Z0)
   a1 = ZRN - 1
   a2 = ZRN + 1
   D = a2 ^ 2 + ZIN ^ 2
   SR = (ZRN ^ 2 - 1 + ZIN ^ 2) / D
   SI = (2 * ZIN) / D
   SM = Sqr(SR ^ 2 + SI ^ 2)
   If (SI = 0) Then SI = 0.00000000000001
   If (SR = 0) Then SR = 0.00000000000001
   SQ = (Atn(SI / SR) * 180 / pi)
   If (SR < 0 And SI > 0) Then SQ = 180 + SQ
   If (SR < 0 And SI < 0) Then SQ = -180 + SQ
   If (SM = 0) Then SQ = 0
   ZtoGammaA = SQ

 End Function

Function ZtoGammaM(ZXR As Double, ZXI As Double) As Double
  Dim ZRN, ZIN, a1, a2, D, SR, SI, SM, SQ As Double
   ZRN = (ZXR / Z0)
   ZIN = (ZXI / Z0)
   a1 = ZRN - 1
   a2 = ZRN + 1
   D = a2 ^ 2 + ZIN ^ 2
   SR = (ZRN ^ 2 - 1 + ZIN ^ 2) / D
   SI = (2 * ZIN) / D
   ZtoGammaM = Sqr(SR ^ 2 + SI ^ 2)

End Function

