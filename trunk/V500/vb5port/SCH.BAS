Attribute VB_Name = "SCH"
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
Global Const pi = 3.141592654
Global Const C = 299800000#

Global YR(14), YI(14), ZR(14), ZI(14) As Double 'ladder admittance and impedance
Global stepS(14)        As Single   ' step size for series part
Global stepP(14)        As Single   ' step size for parallel part
Global StepLR           As Single   ' step size for Load
Global StepLI           As Single
Global StepFF           As Single   ' step size for Frequency
'Global StepFT           As Single
Global s(14)            As Single   ' Series element values
Global p(14)            As Single   ' parallel element vlaues
Global SwpStp(14)       As Single   ' ??
Global Z0               As Single   ' Char. impedance
Global VF               As Single   ' velocity factor
Global freq             As Double   ' in Mhz
Global start1           As Single   ' Start value for sweep
Global stop1            As Single   ' stop value for sweep
Global step1            As Single   ' step value for sweep
Global Element          As Integer  ' value indicates current sweep element
Global Q(14)            As Single   ' Q values for each element
Global S21mag(1010)     As Single   ' S21Data Array
Global S21Angle(1010)   As Single
Global Swpfreq(1010)    As Single   ' Swept freq to write onto a file or graph
Global S21IM#, S21IA#, S21IR#, S21II# ' stores spot data of S21 Calculations
Global GammaMag(1010)   As Single   ' swept gamma to write onto a file
Global GammaAngle(1010) As Single
Global z(14)            As Single   'characterstic impedance of the transmissin line
Global TL               As Integer  'trans. line length units 0=inch,1=deg,2=mm
Global TLfreq           As Single  ' Transmissin line freq in Mhz
Global SCHl             As Integer ' to show default extn in Cmdialog
Global MULTITER         As Integer ' flag to determine termination mode
' Interpolate
Global xatemp(1010), y1atemp(1010), y2atemp(1010) As Double 'used for inputing data from gamma file for interpolation
Global xa(5), y1a(5), y2a(5) As Double  'adjacent values for interpolation
Global npoints          As Integer  'number of points interpolated
Global gammaM           As Single   'termination in terms of gamma
Global GammaA           As Single

Sub assign()
Dim w As Single
Dim I%, X!, beta#, rad!, LR!, LI!, a!, B!, Cd!, Dd!, xc!, xl!, R!
Dim n, j, k, m As Long  'fir min max interpolate
Dim ju, jm, jl As Long   ' for locate
Dim Ri, Qi As Double
On Error GoTo AssignError
'Assign Frequency
form1.Smith1.Z0 = Z0
If freq <= 0 Then freq = 0.000000000001
w = 2 * pi * freq * 1000000#

'Assign termination

If (MULTITER = 1) Then ' for mutilple termination interpolate Load

    n = npoints ' tolal number of data available
    X = freq ' locate index j closest to x use the xa Array for this
    jl = 0      ' initialize lower and upper limits
    ju = n + 1
    While ((ju - jl) > 1) 'if we are not done yet
        jm = (ju + jl) / 2    'compute mid point first
        If (X > xatemp(jm)) Then jl = jm Else ju = jm
    Wend
    j = jl  ' returns index j such as j < x < J+1

    ' get the leftmost index of a point m such that j is at the center
    m = 4 ' number of sample points
    k = min(max(j - (m - 1) / 2, 1), n + 1 - m) ' find midpoint offset
    If m >= n Then
        k = 1
        m = n
    End If
    ' Assign Values for interpolating
    For I = 1 To m Step 1
        xa(I) = xatemp(k - 1 + I)
        y1a(I) = y1atemp(k - 1 + I)
        y2a(I) = y2atemp(k - 1 + I)
    Next I
    Ri = interpolate(X, m, 1)
    Qi = interpolate(X, m, 2)

    ZR(1) = GToZR(CSng(Ri), CSng(Qi))
    ZI(1) = GToZI(CSng(Ri), CSng(Qi))
    gammaM = Ri
    GammaA = Qi

    If (form2.ImpeItem.Checked) Then
       If (Abs(ZR(1) > 99999.999)) Then form2.Panel3D4.Caption = Format$(ZR(1), "##.#E+##") Else form2.Panel3D4.Caption = Format(ZR(1), "0.000")
       If (Abs(ZI(1) > 99999.999)) Then form2.Panel3D5.Caption = Format$(ZI(1), "##.#E+##") Else form2.Panel3D5.Caption = Format(ZI(1), "0.000")
    Else
      If (Abs(gammaM) > 99999.999) Then form2.Panel3D4.Caption = Format$(gammaM, "##.#E+##") Else form2.Panel3D4.Caption = Format(gammaM, "0.000")
      If (Abs(GammaA) > 99999.999) Then form2.Panel3D5.Caption = Format$(GammaA, "##.#E+##") Else form2.Panel3D5.Caption = Format(GammaA, "0.000")
    End If

Else
   
   If (form2.ImpeItem.Checked) Then ' if termination is in impedance
        gammaM = ZtoGammaM(CDbl(ZR(1)), CDbl(ZI(1))) 'update gamma
        GammaA = ZtoGammaA(CDbl(ZR(1)), CDbl(ZI(1)))
       If (Abs(ZR(1) > 99999.999)) Then form2.Panel3D4.Caption = Format$(ZR(1), "##.#E+##") Else form2.Panel3D4.Caption = Format(ZR(1), "0.000")
       If (Abs(ZI(1) > 99999.999)) Then form2.Panel3D5.Caption = Format$(ZI(1), "##.#E+##") Else form2.Panel3D5.Caption = Format(ZI(1), "0.000")
   Else                    ' if termination is in gamma
        ZR(1) = GToZR(CSng(gammaM), CSng(GammaA)) 'update impedance
        ZI(1) = GToZI(CSng(gammaM), CSng(GammaA))
        If (Abs(gammaM) > 99999.999) Then form2.Panel3D4.Caption = Format$(gammaM, "##.#E+##") Else form2.Panel3D4.Caption = Format(gammaM, "0.000")
        If (Abs(GammaA) > 99999.999) Then form2.Panel3D5.Caption = Format$(GammaA, "##.#E+##") Else form2.Panel3D5.Caption = Format(GammaA, "0.000")
   End If

End If      ' end of interpolation

If ((ZR(1) = 0) And (ZI(1) = 0)) Then ZR(1) = 1E-45


  For I% = 3 To 9 Step 2    ' parallel
    Select Case form2.Image12(I%).Tag
        
        Case "C"
          If p(I%) = 0 Then p(I%) = 1E-16
          X! = (w * p(I%) * 0.000000000001)
         YR(I%) = (Q(I%) * X!) / (Q(I%) ^ 2 + 1)
         YI(I%) = (Q(I%) ^ 2 * X!) / (Q(I%) ^ 2 + 1)
        

        Case "L"
          If p(I%) = 0 Then p(I%) = 1E-16
          X! = 1 / (w * p(I%) * 0.000000001)
         YR(I%) = (Q(I%) * X!) / (Q(I%) ^ 2 + 1)
         YI(I%) = -(Q(I%) ^ 2 * X!) / (Q(I%) ^ 2 + 1)
         
        Case "S"
            beta# = w / (VF * C)
            Select Case TL
            Case 0
            rad! = beta# * p(I%) * 0.0254    ' P(I%) in inches
            Case 1
            rad! = beta# * p(I%) / 1000     ' P(I%) in mm
            Case 2                          ' P(I%) in degrees
            rad! = beta# * p(I%) * VF * C / (TLfreq * 1000000 * 360)
            Case 3                          ' P(I%) in lamda
            rad! = beta# * p(I%) * VF * C / (TLfreq * 1000000)
            Case 4                          ' P(I%) in meters
            rad! = beta# * p(I%)
            
            Case Else
            rad! = beta# * p(I%) * 0.0254     ' P(I%) in inches
            End Select
            
            If (rad! * 180 / pi) = 90 Then rad! = rad! + 0.000000000001
            If z(I%) = 0 Then z(I%) = 0.000000000001
            YI(I%) = -1 / (Abs(z(I%)) * Tan(rad!))
            YR(I%) = 0
        Case "O"
            beta# = w / (VF * C)
            Select Case TL
            Case 0
            rad! = beta# * p(I%) * 0.0254    ' P(I%) in inches
            Case 1
            rad! = beta# * p(I%) / 1000     ' P(I%) in mm
            Case 2                           ' P(I%) in degrees
            rad! = beta# * p(I%) * VF * C / (TLfreq * 1000000 * 360)
            Case 3                          ' P(I%) in lamda
            rad! = beta# * p(I%) * VF * C / (TLfreq * 1000000)
            Case 4                          ' P(I%) in meters
            rad! = beta# * p(I%)
            
            Case Else
            rad! = beta# * p(I%) * 0.0254     ' P(I%) in inches
            End Select
            
            'rad! = beta# * P(I%) * .0254
            If (rad! * 180 / pi) = 90 Then rad! = rad! + 0.000000000001
            If z(I%) = 0 Then z(I%) = 0.000000000001
            YI(I%) = Tan(rad!) / Abs(z(I%))
            YR(I%) = 0
        Case "R"
            YI(I%) = 0
            If p(I%) = 0 Then p(I%) = 1E-16
            YR(I%) = 1 / p(I%)
        Case "X"
            If p(I%) = 0 Then p(I%) = 1E-16
            YI(I%) = -1 / p(I%)
            YR(I%) = 0
        Case "W"
            YI(I%) = 0
            YR(I%) = 0
        Case "SLC"
           If p(I%) = 0 Then p(I%) = 0.0000000001
           X! = -1 / (w * p(I%) * 0.000000000001) + (w * z(I%) * 0.000000001)
           R! = (w * z(I%) * 0.000000001) / Q(I%)
           YI(I%) = -X! / (X! ^ 2 + R! ^ 2)
           YR(I%) = R! / (X! ^ 2 + R! ^ 2)
        Case "PLC"
          If p(I%) = 0 Then p(I%) = 1E-16
          If z(I%) = 0 Then z(I%) = 1E-16
          X! = ((z(I%) * 0.000000001) / (p(I%) * 0.000000000001)) / ((w * z(I%) * 0.000000001) - 1 / (w * p(I%) * 0.000000000001))
          YI(I%) = 1 / X!
          YR(I%) = 1 / (w * z(I%) * 0.000000001 * Q(I%))
        Case "SRC"
          If p(I%) = 0 Then p(I%) = 1E-16
          If z(I%) = 0 Then z(I%) = 1E-16
           X! = 1 / (w * p(I%) * 0.000000000001)  'cap
           R! = z(I%)
           YI(I%) = X! / (X! ^ 2 + R! ^ 2)
           YR(I%) = R! / (X! ^ 2 + R! ^ 2)
        Case "PRC"
          If p(I%) = 0 Then p(I%) = 1E-16
          If z(I%) = 0 Then z(I%) = 1E-16
          YI(I%) = (w * p(I%) * 0.000000000001)   ' capacitor
          YR(I%) = 1 / z(I%)


          
   
   End Select
 Next I%
 
 For I% = 2 To 10 Step 2            ' series
     Select Case form2.Image11(I%).Tag
      
        Case "C"
           If s(I%) = 0 Then s(I%) = 0.0000000001
           ZI(I%) = -1 / (w * s(I%) * 0.000000000001)
           ZR(I%) = -ZI(I%) / Q(I%)
       Case "L"
          ZI(I%) = w * s(I%) * 0.000000001
          ZR(I%) = ZI(I%) / Q(I%)
        Case "T"
          beta# = w / (VF * C)
          Select Case TL
            Case 0
            rad! = beta# * s(I%) * 0.0254    ' S(I%) in inches
            Case 1
            rad! = beta# * s(I%) / 1000     ' S(I%) in mm
            Case 2 ' S(I%) in degrees
            rad! = beta# * s(I%) * VF * C / (TLfreq * 1000000 * 360)
            Case 3                          ' S(I%) in lamda
            rad! = beta# * s(I%) * VF * C / (TLfreq * 1000000)
            Case 4                          ' S(I%) in meters
            rad! = beta# * s(I%)
            
            Case Else
            rad! = beta# * s(I%) * 0.0254     ' S(I%) in inches
            End Select
          
          'rad! = beta# * S(I%) * .0254
          LR! = ZR(1)
          LI! = ZI(1)
          If (I% >= 4) Then
          ZR(0) = ZR(2) + ZR(1)
          ZI(0) = ZI(2) + ZI(1)
          YR(2) = YR(3) + ZR(0) / (ZR(0) ^ 2 + ZI(0) ^ 2)
          YI(2) = YI(3) - ZI(0) / (ZR(0) ^ 2 + ZI(0) ^ 2)
          LR! = YR(2) / (YR(2) ^ 2 + YI(2) ^ 2)
          LI! = -YI(2) / (YR(2) ^ 2 + YI(2) ^ 2)
          End If
          
          If (I% >= 6) Then
          ZR(3) = ZR(4) + YR(2) / (YR(2) ^ 2 + YI(2) ^ 2)
          ZI(3) = ZI(4) - YI(2) / (YR(2) ^ 2 + YI(2) ^ 2)
          YR(4) = YR(5) + ZR(3) / (ZR(3) ^ 2 + ZI(3) ^ 2)
          YI(4) = YI(5) - ZI(3) / (ZR(3) ^ 2 + ZI(3) ^ 2)
          LR! = YR(4) / (YR(4) ^ 2 + YI(4) ^ 2)
          LI! = -YI(4) / (YR(4) ^ 2 + YI(4) ^ 2)
          End If
          
          If (I% >= 8) Then
          ZR(5) = ZR(6) + YR(4) / (YR(4) ^ 2 + YI(4) ^ 2)
          ZI(5) = ZI(6) - YI(4) / (YR(4) ^ 2 + YI(4) ^ 2)
          YR(6) = YR(7) + ZR(5) / (ZR(5) ^ 2 + ZI(5) ^ 2)
          YI(6) = YI(7) - ZI(5) / (ZR(5) ^ 2 + ZI(5) ^ 2)
          LR! = YR(6) / (YR(6) ^ 2 + YI(6) ^ 2)
          LI! = -YI(6) / (YR(6) ^ 2 + YI(6) ^ 2)
          End If
          
          If (I% = 10) Then
          ZR(7) = ZR(8) + YR(6) / (YR(6) ^ 2 + YI(6) ^ 2)
          ZI(7) = ZI(8) - YI(6) / (YR(6) ^ 2 + YI(6) ^ 2)
          YR(8) = YR(9) + ZR(7) / (ZR(7) ^ 2 + ZI(7) ^ 2)
          YI(8) = YI(9) - ZI(7) / (ZR(7) ^ 2 + ZI(7) ^ 2)
          LR! = YR(8) / (YR(8) ^ 2 + YI(8) ^ 2)
          LI! = -YI(8) / (YR(8) ^ 2 + YI(8) ^ 2)
          End If
          
          If z(I%) = 0 Then z(I%) = 0.000000000001
          z(I%) = Abs(z(I%))
          a! = LR! * Cos(rad!) ' EA
          B! = z(I%) * Sin(rad!) + (LI! * Cos(rad!)) 'B+ FA
          Cd! = z(I%) * Cos(rad!) - (LI! * Sin(rad!)) ' C-FD
          Dd! = LR! * Sin(rad!) 'ED
          ZR(I%) = z(I%) * DivRe(a!, B!, Cd!, Dd!) - LR!
          ZI(I%) = z(I%) * DivIm(a!, B!, Cd!, Dd!) - LI!
        Case "R"
            ZI(I%) = 0
            ZR(I%) = s(I%)
        Case "X"
            ZI(I%) = s(I%)
            ZR(I%) = 0
        Case "W"
            ZI(I%) = 0
            ZR(I%) = 0
        Case "SLC"
           If s(I%) = 0 Then s(I%) = 0.0000000001
           ZI(I%) = -1 / (w * s(I%) * 0.000000000001) + (w * z(I%) * 0.000000001)
           ZR(I%) = (w * z(I%) * 0.000000001) / Q(I%)

        Case "PLC"
          If s(I%) = 0 Then s(I%) = 1E-16 'Else s(i%) = s(i%) * .000000000001
          If z(I%) = 0 Then z(I%) = 1E-16 'Else z(i%) = z(i%) * .000000001
          X! = -((z(I%) * 0.000000001) / (s(I%) * 0.000000000001)) / ((w * z(I%) * 0.000000001) - 1 / (w * s(I%) * 0.000000000001))
          R! = (w * z(I%) * 0.000000001 * Q(I%))
          ZI(I%) = X! / (1 + X! ^ 2 / R! ^ 2)
          ZR(I%) = R! / (1 + R! ^ 2 / X! ^ 2)
        Case "SRC"
          If s(I%) = 0 Then s(I%) = 1E-16 'Else s(i%) = s(i%) * .000000000001
          If z(I%) = 0 Then z(I%) = 1E-16 'Else z(i%) = z(i%) * .000000001
          ZI(I%) = -1 / (w * s(I%) * 0.000000000001)  ' capacitor
          ZR(I%) = z(I%)  ' resistor
        Case "PRC"
          If s(I%) = 0 Then s(I%) = 1E-16 'Else s(i%) = s(i%) * .000000000001
          If z(I%) = 0 Then z(I%) = 1E-16 'Else z(i%) = z(i%) * .000000001
          X! = -(w * s(I%) * 0.000000000001) 'cap
          R! = 1 / z(I%)
          ZI(I%) = X! / (X! ^ 2 + R! ^ 2)
          ZR(I%) = R! / (X! ^ 2 + R! ^ 2)
    
    End Select
 Next I%

Exit Sub
AssignError:
Beep
If (Err = 11) Then
    response% = MsgBox("'A Division By Zero' Error Was Trapped", 49, "Error")
    If response = 1 Then Resume Else Exit Sub
Else: Exit Sub
End If

End Sub

Sub Ilcal(arg As Integer)
' 0 means spot, 1 means sweep , >1 means donot call assign
Dim f, fold As Single
ReDim CR(14), CI(14), VR(14), VI(14) As Double
Dim points%, k%, j%, I%, CRA#, CIA#, C1#, dB1!

  On Error GoTo ILCalError
  fold = freq
  points% = Abs((stop1 - start1) / step1) + 1 ' calculate number of points to sweep
  
  'form2.Print points%
  f = start1 - step1
  
  For k% = 1 To points% Step 1
  freq = (f + step1) ' frequency sweep
  If (arg <> 1) Then freq = fold
  If (arg < 2) Then Call assign
  If (ZR(1) = 0) Then ZR(1) = 1E-40
  YR(1) = 1 / ZR(1)
  YI(1) = 0

  YR(11) = 0    ' not used
  YI(11) = 0
  ZR(12) = 0
  ZI(12) = 0


  YR(13) = YR(1)
  YI(13) = YI(1)
  CR(2) = YR(1)
  CI(2) = 0
  VI(2) = YR(1) * ZI(2)
  VR(2) = YR(1) * ZR(2)
  VR(3) = 1 + VR(2)  ' VR(1) = 1 Volt
  VI(3) = VI(2)
  CR(3) = VR(3) * YR(3) - VI(3) * YI(3)
  CI(3) = VR(3) * YI(3) + VI(3) * YR(3)
  For j% = 4 To 12 Step 2
  CR(j%) = CR(j% - 2) + CR(j% - 1)
  CI(j%) = CI(j% - 2) + CI(j% - 1)
  VR(j%) = CR(j%) * ZR(j%) - CI(j%) * ZI(j%)
  VI(j%) = CR(j%) * ZI(j%) + CI(j%) * ZR(j%)
  VR(j% + 1) = VR(j%) + VR(j% - 1)
  VI(j% + 1) = VI(j%) + VI(j% - 1)
  CR(j% + 1) = VR(j% + 1) * YR(j% + 1) - VI(j% + 1) * YI(j% + 1)
  CI(j% + 1) = VR(j% + 1) * YI(j% + 1) + VI(j% + 1) * YR(j% + 1)
  Next j%
  CRA# = CR(13) + CR(12)
  CIA# = CI(13) + CI(12)
  If (CRA# = 0) Then CRA# = 1E-45
  C1# = Sqr(CRA# ^ 2 + CIA# ^ 2)
  'S21IM# = C1#    ' store global data
  'S21IR# = CRA#
  'S21II# = CIA#
  dB1! = -(20 / 2.30259) * Log(C1# / (2 * YR(1))) ' il loss = 20 log (v1/v2)
  f = freq
  Swpfreq(k%) = f
  S21mag(k%) = dB1!
  S21IM# = dB1!
  S21Angle(k%) = (Atn(CIA# / CRA#) * 180 / pi)
  If (CRA# < 0 And CIA# > 0) Then S21Angle(k%) = 180 + S21Angle(k%)
  If (CRA# < 0 And CIA# < 0) Then S21Angle(k%) = -180 + S21Angle(k%)
  S21Angle(k%) = -S21Angle(k%)
  S21IA# = S21Angle(k%) ' store global data
  If (f > 99999.999) Then form2.Panel3D10.Caption = Format$(f, "##.#E+##") Else form2.Panel3D10.Caption = Format$(f, "###0.000")
  form2.Panel3D12.Caption = Format(-dB1!, "###0.000")
  If (arg <> 1) Then k% = points%
 
  Next k%
  freq = fold ' get the old freq value

Exit Sub
ILCalError:
If (Err = 11) Then
    response% = MsgBox("'A Division By Zero' Error Was Trapped", 49, "Error")
    If response = 1 Then Resume Else Exit Sub
Else: Exit Sub

End If
  
End Sub

'Neville-Aitken interpolation Algorithm
'xa and ya are global arrays, needs number of points not greater than 200 ad x Value to interpolate
'column number decides the coulmns 1 for first, 2 for second
'n is usually 4
'for proper interpolation phase values needs to be unwrapped, phase is unwrapped in Link_item_click()
'
Function interpolate(ByVal X As Double, ByVal n As Integer, ByVal column As Integer) As Double

Static C(5), D(5)
Dim I, m, ns As Integer
Dim den, dif, dift, ho, hp, w, Y, dy As Double
If (n > 10) Then GoTo endinter
ns = 1
xa(0) = 0
'x = 23
dif = Abs(X - xa(1))    ' find the index ns closest to the entry
For I = 1 To n Step 1
 dift = Abs(X - xa(I))
    If (dift < dif) Then
        ns = I
        dif = dift
    End If
 If (column = 1) Then
 C(I) = y1a(I)
 D(I) = y1a(I)
 Else
 C(I) = y2a(I)
 D(I) = y2a(I)
 End If
Next I
   If (column = 1) Then Y = y1a(ns) Else Y = y2a(ns)
   
   ns = ns - 1
For m = 1 To (n - 1) Step 1
     For I = 1 To (n - m) Step 1
       ho = xa(I) - X
       hp = xa(I + m) - X
       w = C(I + 1) - D(I)
       den = ho - hp
       If (den = 0) Then GoTo endinter
       den = w / den
       D(I) = hp * den
       C(I) = ho * den
       Next I
   If (2 * ns < (n - m)) Then
   dy = C(ns + 1)
   Else
   
   dy = D(ns)
   ns = ns - 1
   End If
   Y = Y + dy
Next m
interpolate = Y
Exit Function

endinter:
       Beep
       MsgBox "Interpolation Error", 0, "QuickSmith"

End Function

Sub loadValue(Tag As Variant, dat1, dat2)
' this sub is called by schematic file_open routine
If (CStr(Tag) = "I") Then
  If (IsNumeric(dat1)) Then
    ZR(1) = CDbl(dat1)
    If (Abs(CDbl(dat1)) > 99999.99) Then form2.Panel3D4.Caption = Format(CDbl(dat1), "##.#E+##") Else form2.Panel3D4.Caption = Format(CDbl(dat2), "0.000")
    form2.Label1(1) = "R"
  End If
  If (IsNumeric(dat2)) Then
    ZI(1) = CDbl(dat2)
    If (Abs(CDbl(dat2)) > 99999.99) Then form2.Panel3D5.Caption = Format(CDbl(dat2), "##.#E+##") Else form2.Panel3D5.Caption = Format(CDbl(dat2), "0.000")
    form2.Label1(2) = "L"
  End If
form2.ImpeItem.Checked = True
form2.GammaItem.Checked = False

End If

If (CStr(Tag) = "G") Then
  If (IsNumeric(dat1)) Then
    gammaM = CDbl(dat1)
    If (Abs(CDbl(dat1)) > 99999.99) Then form2.Panel3D4.Caption = Format(CDbl(dat1), "##.#E+##") Else form2.Panel3D4.Caption = Format(CDbl(dat2), "0.000")
    form2.Label1(1) = "M"
  End If
  If (IsNumeric(dat2)) Then
    GammaA = CDbl(dat2)
    If (Abs(CDbl(dat2)) > 99999.99) Then form2.Panel3D5.Caption = Format(CDbl(dat2), "##.#E+##") Else form2.Panel3D5.Caption = Format(CDbl(dat2), "0.000")
    form2.Label1(2) = "A"
  End If
form2.ImpeItem.Checked = False
form2.GammaItem.Checked = True

End If





End Sub

Function max(ByVal X As Double, ByVal Y As Integer) As Integer
If X > Y Then max = X Else max = Y
End Function

Function min(ByVal X As Double, ByVal Y As Integer) As Integer
If X < Y Then min = X Else min = Y
End Function

Sub ParallelIO(Tag As Variant, X As Variant, index As Integer)
        Dim inp$, msg
        On Error Resume Next
        Select Case Tag
        

        Case "R"
            inp$ = X
            If (IsNumeric(inp$)) Then
            p(index) = CDbl(inp$)
            If (Abs(CDbl(inp$)) > 99999.99) Then form2.Panel3D6(index).Caption = Format$(CDbl(inp$), "##.#E+##") Else form2.Panel3D6(index).Caption = Format$(CDbl(inp$), "0.000")
            form2.Image12(index).Picture = LoadPicture(CurrentDir & "\rV.bmp")
            form2.Image12(index).Tag = "R"
            form2.Label2(index).Caption = "R" & index
            End If
        
        Case "C"
           inp$ = X
           If (IsNumeric(inp$)) Then
           p(index) = CDbl(inp$)  ' will be trapped in zcal
           If (Abs(CDbl(inp$)) > 99999.99) Then form2.Panel3D6(index).Caption = Format$(CDbl(inp$), "##.#E+##") Else form2.Panel3D6(index).Caption = Format$(CDbl(inp$), "0.000")
           form2.Image12(index).Picture = LoadPicture(CurrentDir & "\cV.bmp")
           form2.Image12(index).Tag = "C"
           form2.Label2(index).Caption = "C" & index
           End If
        
        Case "L"
           inp$ = X
           If (IsNumeric(inp$)) Then
           p(index) = CDbl(inp$)    ' will be trapped in zcal
           If (Abs(CDbl(inp$)) > 99999.99) Then form2.Panel3D6(index).Caption = Format$(CDbl(inp$), "##.#E+##") Else form2.Panel3D6(index).Caption = Format$(CDbl(inp$), "0.000")
           form2.Image12(index).Picture = LoadPicture(CurrentDir & "\IV.bmp")
           form2.Image12(index).Tag = "L"
           form2.Label2(index).Caption = "L" & index
           End If
        
        Case "W"
           form2.Image12(index).Picture = LoadPicture(CurrentDir & "\wV.bmp")
           form2.Panel3D6(index).Caption = ""
           s(index) = 0
           s(index) = 0
           form2.Image12(index).Tag = "W"
           form2.Label2(index).Caption = "W" & index
        Case "X"
            inp$ = X
            If (IsNumeric(inp$)) Then
            p(index) = CDbl(inp$)
            If (Abs(CDbl(inp$)) > 99999.99) Then form2.Panel3D6(index).Caption = Format$(CDbl(inp$), "##.#E+##") Else form2.Panel3D6(index).Caption = Format$(CDbl(inp$), "0.000")
            form2.Image12(index).Picture = LoadPicture(CurrentDir & "\XV.bmp")
            form2.Image12(index).Tag = "X"
            form2.Label2(index).Caption = "X" & index
            End If
        Case "O"
           inp$ = X
           If (IsNumeric(inp$)) Then
           p(index) = CDbl(inp$)  ' will be trapped in zcal
           If (Abs(CDbl(inp$)) > 99999.99) Then form2.Panel3D6(index).Caption = Format$(CDbl(inp$), "##.#E+##") Else form2.Panel3D6(index).Caption = Format$(CDbl(inp$), "0.000")
            form2.Option3D1.Value = True
           form2.Image12(index).Picture = LoadPicture(CurrentDir & "\stubo.bmp")
           form2.Image12(index).Tag = "O"
           form2.Label2(index).Caption = "O" & index
           End If
        
        Case "S"
           inp$ = X
           If (IsNumeric(inp$)) Then
           p(index) = CDbl(inp$)  ' will be trapped in zcal
           If (Abs(CDbl(inp$)) > 99999.99) Then form2.Panel3D6(index).Caption = Format$(CDbl(inp$), "##.#E+##") Else form2.Panel3D6(index).Caption = Format$(CDbl(inp$), "0.000")
           form2.Option3D1.Value = True
           form2.Image12(index).Picture = LoadPicture(CurrentDir & "\stubs.bmp")
           form2.Image12(index).Tag = "S"
           form2.Label2(index).Caption = "S" & index
           End If
        Case "SLC"
           inp$ = X
           If (IsNumeric(inp$)) Then
           p(index) = CDbl(inp$)  ' will be trapped in zcal
           If (Abs(CDbl(inp$)) > 99999.99) Then form2.Panel3D6(index).Caption = Format$(CDbl(inp$), "##.#E+##") Else form2.Panel3D6(index).Caption = Format$(CDbl(inp$), "0.000")
           form2.Option3D1.Value = True
           form2.Image12(index).Picture = LoadPicture(CurrentDir & "\CISV.bmp")
           form2.Image12(index).Tag = "SLC"
           form2.Label2(index).Caption = "SLC" & index
           End If
        Case "PLC"
           inp$ = X
           If (IsNumeric(inp$)) Then
           p(index) = CDbl(inp$)  ' will be trapped in zcal
           If (Abs(CDbl(inp$)) > 99999.99) Then form2.Panel3D6(index).Caption = Format$(CDbl(inp$), "##.#E+##") Else form2.Panel3D6(index).Caption = Format$(CDbl(inp$), "0.000")
           form2.Option3D1.Value = True
           form2.Image12(index).Picture = LoadPicture(CurrentDir & "\CIPV.bmp")
           form2.Image12(index).Tag = "PLC"
           form2.Label2(index).Caption = "PLC" & index
           End If
        Case "SRC"
           inp$ = X
           If (IsNumeric(inp$)) Then
           p(index) = CDbl(inp$)  ' will be trapped in zcal
           If (Abs(CDbl(inp$)) > 99999.99) Then form2.Panel3D6(index).Caption = Format$(CDbl(inp$), "##.#E+##") Else form2.Panel3D6(index).Caption = Format$(CDbl(inp$), "0.000")
           form2.Option3D1.Value = True
           form2.Image12(index).Picture = LoadPicture(CurrentDir & "\CRSV.bmp")
           form2.Image12(index).Tag = "SRC"
           form2.Label2(index).Caption = "SRC" & index
           End If
        Case "PRC"
           inp$ = X
           If (IsNumeric(inp$)) Then
           p(index) = CDbl(inp$)  ' will be trapped in zcal
           If (Abs(CDbl(inp$)) > 99999.99) Then form2.Panel3D6(index).Caption = Format$(CDbl(inp$), "##.#E+##") Else form2.Panel3D6(index).Caption = Format$(CDbl(inp$), "0.000")
           form2.Option3D1.Value = True
           form2.Image12(index).Picture = LoadPicture(CurrentDir & "\CRPV.bmp")
           form2.Image12(index).Tag = "PRC"
           form2.Label2(index).Caption = "PRC" & index
           End If


       Case "T"
            Beep
        End Select
       If Err Then
         msg = "QuickSmith could not find  picture file."
         MsgBox msg & CurrentDir, 0, "QuickSmith Error"
         Exit Sub
       End If
' stepS(index) = 1
 'Call Zcal

End Sub

Sub SeriesIO(Tag As Variant, X As Variant, index As Integer)
         Dim inp$, msg
         On Error Resume Next
        Select Case Tag
        
        Case "R"
            inp$ = X
            If (IsNumeric(inp$)) Then
            s(index) = CDbl(inp$)
            If (Abs(CDbl(inp$)) > 99999.99) Then form2.Panel3D3(index).Caption = Format$(CDbl(inp$), "##.#E+##") Else form2.Panel3D3(index).Caption = Format$(CDbl(inp$), "0.000")
            form2.Image11(index).Picture = LoadPicture(CurrentDir & "\r.bmp")
            form2.Image11(index).Tag = "R"
            form2.Label2(index).Caption = "R" & index
            End If
        
        Case "C"
           inp$ = X
           If (IsNumeric(inp$)) Then
           s(index) = CDbl(inp$)  ' will be trapped in zcal
           If (Abs(CDbl(inp$)) > 99999.99) Then form2.Panel3D3(index).Caption = Format$(CDbl(inp$), "##.#E+##") Else form2.Panel3D3(index).Caption = Format$(CDbl(inp$), "0.000")
           form2.Image11(index).Picture = LoadPicture(CurrentDir & "\c.bmp")
           form2.Image11(index).Tag = "C"
           form2.Label2(index).Caption = "C" & index
           End If
        
        Case "L"
           inp$ = X
           If (IsNumeric(inp$)) Then
           s(index) = CDbl(inp$)    ' will be trapped in zcal
           If (Abs(CDbl(inp$)) > 99999.99) Then form2.Panel3D3(index).Caption = Format$(CDbl(inp$), "##.#E+##") Else form2.Panel3D3(index).Caption = Format$(CDbl(inp$), "0.000")
           form2.Image11(index).Picture = LoadPicture(CurrentDir & "\I.bmp")
           form2.Image11(index).Tag = "L"
           form2.Label2(index).Caption = "L" & index
           End If
        
        Case "W"
           form2.Image11(index).Picture = LoadPicture(CurrentDir & "\w.bmp")
           form2.Panel3D3(index).Caption = ""
           s(index) = 0
           s(index) = 0
           form2.Image11(index).Tag = "W"
           form2.Label2(index).Caption = "W" & index
        Case "X"
            inp$ = X
            If (IsNumeric(inp$)) Then
            s(index) = CDbl(inp$)
            If (Abs(CDbl(inp$)) > 99999.99) Then form2.Panel3D3(index).Caption = Format$(CDbl(inp$), "##.#E+##") Else form2.Panel3D3(index).Caption = Format$(CDbl(inp$), "0.000")
            form2.Image11(index).Picture = LoadPicture(CurrentDir & "\X.bmp")
            form2.Image11(index).Tag = "X"
            form2.Label2(index).Caption = "X" & index
            End If
        
        Case "T"
           inp$ = X
           If (IsNumeric(inp$)) Then
           s(index) = CDbl(inp$)   ' will be trapped in zcal
           If (Abs(CDbl(inp$)) > 99999.99) Then form2.Panel3D3(index).Caption = Format$(CDbl(inp$), "##.#E+##") Else form2.Panel3D3(index).Caption = Format$(CDbl(inp$), "0.000")
           form2.Option3D1.Value = True
           form2.Image11(index).Picture = LoadPicture(CurrentDir & "\slug.bmp")
           form2.Image11(index).Tag = "T"
           form2.Label2(index).Caption = "T" & index
           End If
        Case "SLC"
           inp$ = X
           If (IsNumeric(inp$)) Then
           s(index) = CDbl(inp$)   ' will be trapped in zcal
           If (Abs(CDbl(inp$)) > 99999.99) Then form2.Panel3D3(index).Caption = Format$(CDbl(inp$), "##.#E+##") Else form2.Panel3D3(index).Caption = Format$(CDbl(inp$), "0.000")
           form2.Option3D1.Value = True
           form2.Image11(index).Picture = LoadPicture(CurrentDir & "\CIS.bmp")
           form2.Image11(index).Tag = "SLC"
           form2.Label2(index).Caption = "SLC" & index
           End If
        Case "PLC"
           inp$ = X
           If (IsNumeric(inp$)) Then
           s(index) = CDbl(inp$)   ' will be trapped in zcal
           If (Abs(CDbl(inp$)) > 99999.99) Then form2.Panel3D3(index).Caption = Format$(CDbl(inp$), "##.#E+##") Else form2.Panel3D3(index).Caption = Format$(CDbl(inp$), "0.000")
           form2.Option3D1.Value = True
           form2.Image11(index).Picture = LoadPicture(CurrentDir & "\CIP.bmp")
           form2.Image11(index).Tag = "PLC"
           form2.Label2(index).Caption = "PLC" & index
           End If
        Case "SRC"
           inp$ = X
           If (IsNumeric(inp$)) Then
           s(index) = CDbl(inp$)   ' will be trapped in zcal
           If (Abs(CDbl(inp$)) > 99999.99) Then form2.Panel3D3(index).Caption = Format$(CDbl(inp$), "##.#E+##") Else form2.Panel3D3(index).Caption = Format$(CDbl(inp$), "0.000")
           form2.Option3D1.Value = True
           form2.Image11(index).Picture = LoadPicture(CurrentDir & "\CRS.bmp")
           form2.Image11(index).Tag = "SRC"
           form2.Label2(index).Caption = "SRC" & index
           End If
        Case "PRC"
           inp$ = X
           If (IsNumeric(inp$)) Then
           s(index) = CDbl(inp$)   ' will be trapped in zcal
           If (Abs(CDbl(inp$)) > 99999.99) Then form2.Panel3D3(index).Caption = Format$(CDbl(inp$), "##.#E+##") Else form2.Panel3D3(index).Caption = Format$(CDbl(inp$), "0.000")
           form2.Option3D1.Value = True
           form2.Image11(index).Picture = LoadPicture(CurrentDir & "\CRP.bmp")
           form2.Image11(index).Tag = "PRC"
           form2.Label2(index).Caption = "PRC" & index
           End If
        
        
        Case "S", "O"
            Beep
        End Select
       
       If Err Then
         msg = "QuickSmith could not find picture file."
         MsgBox msg, 0, "QuickSmith Error"
         Exit Sub
       End If
     
 'stepS(index) = 1
 'Call Zcal

End Sub

Sub Zcalsweep(index As Integer)
' 0-100 for the first element eg length and 101-200 for sub element Z0 etc
Dim w, f, oldf, oldS, oldP, oldZ As Single
Dim points%, k%, j%, I%, ZRN#, ZIN#, a1#, a2#, D#, SR#, SI#, SM!, SQ!
            'save freq. value
Dim temp% ' save old sweep value for spot only

On Error GoTo ZCalSweepError

points% = Abs((stop1 - start1) / step1)  ' calculate number of points to sweep
   
   Select Case index
    Case 0     ' spot
    Case 1     'freq
       oldf = freq
       If (start1 < 0) Then
        points% = Abs((stop1 - 0) / step1) ' if startvalue for freq sweep is -ve then start sweep from zero
        freq = 0.0000000000001  'start sweep from 0
       Else
       freq = start1
       End If
   Case 2, 4, 6, 8, 10
    oldS = s(index)
    s(index) = start1   ' save old values of the present sweep element
   Case 3, 5, 7, 9
    oldP = p(index)
    p(index) = start1    ' save old values of the present sweep element
   Case Is > 100
    oldZ = z(index - 100)
    z(index - 100) = start1
   End Select
  
   
   
   
   If (index > 0) Then
    form1.Smith1.NumPoints = points% + 2
    form1.Smith1.ThisPoint = 0
    form1.Smith1.sweep = True
   End If


For k% = 0 To points% Step 1
   Call assign ' equate impedance/admittance values to network variables

   ZR(0) = ZR(2) + ZR(1)
   ZI(0) = ZI(2) + ZI(1)
   YR(2) = YR(3) + ZR(0) / (ZR(0) ^ 2 + ZI(0) ^ 2)
   YI(2) = YI(3) - ZI(0) / (ZR(0) ^ 2 + ZI(0) ^ 2)
   For j% = 3 To 7 Step 2
    ZR(j%) = ZR(j% + 1) + YR(j% - 1) / (YR(j% - 1) ^ 2 + YI(j% - 1) ^ 2)
    ZI(j%) = ZI(j% + 1) - YI(j% - 1) / (YR(j% - 1) ^ 2 + YI(j% - 1) ^ 2)
    YR(j% + 1) = YR(j% + 2) + ZR(j%) / (ZR(j%) ^ 2 + ZI(j%) ^ 2)
    YI(j% + 1) = YI(j% + 2) - ZI(j%) / (ZR(j%) ^ 2 + ZI(j%) ^ 2)
   Next j%
   ZR(13) = ZR(10) + YR(8) / (YR(8) ^ 2 + YI(8) ^ 2)
   ZI(13) = ZI(10) - YI(8) / (YR(8) ^ 2 + YI(8) ^ 2)

   If (Abs(ZR(13)) > 99999.9999) Then form1.Label1.Caption = Format$(ZR(13), "##.###E+##") Else form1.Label1.Caption = Format(ZR(13), "####0.0000")
    form2.Panel3D8.Caption = form1.Label1.Caption
   If (Abs(ZI(13)) > 99999.9999) Then form1.Label2.Caption = Format$(ZI(13), "##.###E+##") Else form1.Label2.Caption = Format(ZI(13), "####0.0000")
    form2.Panel3D9.Caption = form1.Label2.Caption
   If (ZI(13) = 0) Then ZI(13) = 1E-45
   If (form1.Option3D1.Value = True) Then Call Yin
    
    SM! = ZtoGammaM(CDbl(ZR(13)), CDbl(ZI(13))) 'convert to gamma
    SQ! = ZtoGammaA(CDbl(ZR(13)), CDbl(ZI(13))) ' for smith chart
   
   Select Case index
    Case 0     ' spot
     k% = points%
     temp% = form1.Smith1.sweep
     form1.Smith1.sweep = False
    Case 1
     freq = freq + step1
     form2.Panel3D10.Caption = Format$(freq, "###0.000")
    Case 2, 4, 6, 8, 10
     s(index) = s(index) + step1
    Case 3, 5, 7, 9
     p(index) = p(index) + step1
    Case Is > 100
    z(index - 100) = z(index - 100) + step1
   End Select
    form1.Smith1.DataM = Format(SM!, ".000")
    form1.Smith1.DataQ = Format(SQ!, ".000")
  If (index > 0) Then
    GammaMag(k%) = SM!   ' Save data for exporting to a file
    GammaAngle(k%) = SQ!
  End If

Next k%

   Select Case index
    Case 0     ' spot
      form1.Smith1.sweep = temp%
    Case 1
     freq = oldf
    Case 2, 4, 6, 8, 10
     s(index) = oldS
    Case 3, 5, 7, 9
     p(index) = oldP
    Case Is > 100
    z(index - 100) = oldZ
   End Select

Call Ilcal(2)
Exit Sub
Beep
ZCalSweepError:
If (Err = 11) Then
    response% = MsgBox("'A Division By Zero' Error Was Trapped", 49, "Error")
    If response = 1 Then Resume Else Exit Sub
Else: Exit Sub
End If

End Sub

