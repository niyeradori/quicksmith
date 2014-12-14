Attribute VB_Name = "AMP"


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
'Global Const Pi = 3.1415926
Global RS, CSA, CSM, RL, CLA, CLM, AK As Single
Global S11R, S11I, S21R, S21I, S12R, S12I, S22R, S22I, DeltaM As Single
Global C1R, C1I, C2R, C2I, CenterR, CenterI, radius, GPMAX As Single
Global LoadAngle, AVLoadAngle As Single
Global STFLAG As Integer
Global Const OPERATING = 1
Global Const AVAILABLE = 0
Global fi, Fmin  As Single
Global AVCenterR, AVCenterI, AVradius As Single ' availabel Data to be ale to access from spincontrol
Global NOISEOPENFLAG, SPAROPENFLAG, AMPOPENFLAG As Integer ' flag to indicate if forms are loaded or not
Global LoadReal, LoadImag, SourceReal, SourceImag
Global vswrstep1  As Single

Function AddIm(a, B, C, D) As Single
AddIm = B + D
End Function

Function AddRe(a, B, C, D) As Single
AddRe = a + C
End Function

Function AmpGTOZI(R, Q, Z00) As Single
Dim tmp As Double
tmp = (2 * R * Sin(pi * Q / 180)) / (1 + (R * R) - 2 * R * Cos(pi * Q / 180))
AmpGTOZI = tmp * Z00
End Function

Function AmpGTOZR(R, Q, Z00) As Single
Dim tmp As Double
tmp = (1 - (R * R)) / (1 + (R * R) - (2 * R * Cos(Q * pi / 180)))
AmpGTOZR = tmp * Z00
End Function

 Sub CircleGamma(CX As Single, CY As Single, R As Single, a As Single, opera As Single)
        'Find a GammaS(Operating) or GammaL(available) on the circle for a given angle
        Dim DataM, DataQ, datax, dataY As Single
        Dim DR, DI, NR, NI As Single
        On Error GoTo CircleGammaError
        'find gammaL
        datax = R * Sin(a * pi / 180)
        dataY = R * Cos(a * pi / 180)
        datax = datax + CX
        dataY = dataY + CY
        DataM = Abs(Sqr(datax ^ 2 + dataY ^ 2))
        DataQ = Atn(dataY / datax) * 180 / pi
        If (datax < 0 And dataY > 0) Then DataQ = 180 + DataQ
        If (datax < 0 And dataY < 0) Then DataQ = -180 + DataQ
        'plot current first and then calculate and plot the corresponding source/load
        If (opera = 1) Then ' write Load
            Form6.Smith2.DataM = DataM
            Form6.Smith2.DataQ = DataQ
            Form6.Panel3D10.Caption = Format(DataM, "0.000") & "  < " & Format(DataQ, "0.000")
            LoadReal = AmpGTOZR(DataM, DataQ, CSng(Form6.Smith2.Z0))
            LoadImag = AmpGTOZI(DataM, DataQ, CSng(Form6.Smith2.Z0))
            Form6.Panel3D9.Caption = Format(LoadReal, "0.000") & " + j " & Format(LoadImag, "0.000")
          Else              ' write Source
            Form6.Smith1.DataM = DataM
            Form6.Smith1.DataQ = DataQ
            SourceReal = AmpGTOZR(DataM, DataQ, CSng(Form6.Smith1.Z0))
            SourceImag = AmpGTOZI(DataM, DataQ, CSng(Form6.Smith1.Z0))
            Form6.Panel3D8.Caption = Format(DataM, "0.000") & "  < " & Format(DataQ, "0.000")
            Form6.Panel3D11.Caption = Format(SourceReal, "0.000") & " + j " & Format(SourceImag, "0.000")
        End If
        'find the corresponding GammaS Gammal and wrie /plot
   If (opera = 1) Then    ' operating   pg122 GG
        NR = MulRe(MulRe(S12R, S12I, S21R, S21I), MulIm(S12R, S12I, S21R, S21I), datax, dataY)
        NI = MulIm(MulRe(S12R, S12I, S21R, S21I), MulIm(S12R, S12I, S21R, S21I), datax, dataY)
        DR = MulRe(S22R, S22I, datax, dataY)
        DI = MulIm(S22R, S22I, datax, dataY)
        DR = AddRe(1, 0, -DR, -DI)
        DI = AddIm(1, 0, -DR, -DI)
        datax = DivRe(NR, NI, DR, DI)
        dataY = DivIm(NR, NI, DR, DI)
        datax = AddRe(S11R, S11I, datax, dataY)
        dataY = AddIm(S11R, S11I, datax, dataY)
        dataY = -dataY ' The required GammaS for maximum power is conjugate
        DataM = Abs(Sqr(datax ^ 2 + dataY ^ 2))
        DataQ = Atn(dataY / datax) * 180 / pi
        If (datax < 0 And dataY > 0) Then DataQ = 180 + DataQ
        If (datax < 0 And dataY < 0) Then DataQ = -180 + DataQ
            
        Form6.Smith1.DataM = DataM        ' write source
        Form6.Smith1.DataQ = DataQ
        Form6.Panel3D8.Caption = Format(DataM, "0.000") & "  < " & Format(DataQ, "0.000")
        SourceReal = AmpGTOZR(DataM, DataQ, CSng(Form6.Smith2.Z0))
        SourceImag = AmpGTOZI(DataM, DataQ, CSng(Form6.Smith2.Z0))
        
        Form6.Panel3D11.Caption = Format(SourceReal, "0.000") & " + j " & Format(SourceImag, "0.000")
  
  Else            ' available pg 147 GG
        NR = MulRe(MulRe(S12R, S12I, S21R, S21I), MulIm(S12R, S12I, S21R, S21I), datax, dataY)
        NI = MulIm(MulRe(S12R, S12I, S21R, S21I), MulIm(S12R, S12I, S21R, S21I), datax, dataY)
        DR = MulRe(S11R, S11I, datax, dataY)
        DI = MulIm(S11R, S11I, datax, dataY)
        DR = AddRe(1, 0, -DR, -DI)
        DI = AddIm(1, 0, -DR, -DI)
        datax = DivRe(NR, NI, DR, DI)
        dataY = DivIm(NR, NI, DR, DI)
        datax = AddRe(S22R, S22I, datax, dataY)
        dataY = AddIm(S22R, S22I, datax, dataY)
        dataY = -dataY ' The required GammaS for maximum power is conjugate
        DataM = Abs(Sqr(datax ^ 2 + dataY ^ 2))
        DataQ = Atn(dataY / datax) * 180 / pi
        If (datax < 0 And dataY > 0) Then DataQ = 180 + DataQ
        If (datax < 0 And dataY < 0) Then DataQ = -180 + DataQ
            
            
            Form6.Smith2.DataM = DataM      ' write load
            Form6.Smith2.DataQ = DataQ
            Form6.Panel3D10.Caption = Format(DataM, "0.000") & "  < " & Format(DataQ, "0.000")
            LoadReal = AmpGTOZR(DataM, DataQ, CSng(Form6.Smith2.Z0))
            LoadImag = AmpGTOZI(DataM, DataQ, CSng(Form6.Smith2.Z0))
            Form6.Panel3D9.Caption = Format(LoadReal, "0.000") & " + j " & Format(LoadImag, "0.000")
       End If
Exit Sub
CircleGammaError:
If (Err = 11) Then
    response% = MsgBox("'A Division By Zero' Error Was Trapped", 49, "Error")
    If response = 1 Then Resume Else Exit Sub
Else: Exit Sub
End If

End Sub

Function DivIm(a, B, C, D) As Single
On Error GoTo DivImError
'(A+ Bj)/(C+ dj)
DivIm = ((B * C) - (a * D)) / (C ^ 2 + D ^ 2)

DivImError:
If (Err = 11) Then
    response% = MsgBox("'A Division By Zero' Error Was Trapped", 49, "Error")
    If response = 1 Then Resume Else Exit Function
Else: Exit Function
End If

End Function

Function DivRe(a, B, C, D) As Single
On Error GoTo DivReError
'(A+ Bj)/(C+ dj)
DivRe = ((a * C) + (B * D)) / (C ^ 2 + D ^ 2)
DivReError:
If (Err = 11) Then
    response% = MsgBox("'A Division By Zero' Error Was Trapped", 49, "Error")
    If response = 1 Then Resume Else Exit Function
Else: Exit Function
End If

End Function

Sub GainCirc(G As Single, Gmax As Single, Oper As Integer)  'opreating or availabel circles
 Dim SR, SI, CR, CI, H1, GI, H2, H2A, H3, Distan, angle, RN, tmp As Single
 'Dim AV, CenterR, AVCenterI, AVRadius As Single  ' Local variables so than spin control does not access it
   'we control the load plane data points only
On Error GoTo GainCircError
         
         If (Oper) Then ' Load plane
         SR = S22R
         SI = S22I
         CR = C2R
         CI = C2I

         H1 = Abs(Sqr(MulRe(S12R, S12I, S21R, S21I) ^ 2 + MulIm(S12R, S12I, S21R, S21I) ^ 2))
         H2 = Abs(Sqr(MulRe(SR, SI, SR, SI) ^ 2 + MulIm(SR, SI, SR, SI) ^ 2)) ' |S11*S11|
         H2A = Abs(Sqr(MulRe(S21R, S21I, S21R, S21I) ^ 2 + MulIm(S21R, S21I, S21R, S21I) ^ 2)) ' |S21*S21|
         H3 = H2 - DeltaM ^ 2
         GI = 10 ^ (G / 10) / H2A
         tmp = GI / (1 + GI * H3)
         CenterR = tmp * CR
         CenterI = tmp * -CI
         Distan = Abs(Sqr(CenterR ^ 2 + CenterI ^ 2))
         angle = Atn(CenterI / CenterR) * 180 / pi
         If (CenterR < 0 And CenterI > 0) Then angle = 180 + angle
         If (CenterR < 0 And CenterI < 0) Then angle = -180 + angle
         If (G < Gmax) Then RN = Sqr(1 - 2 * AK * H1 * GI + (H1 * GI) ^ 2) Else RN = 0
         radius = RN / (1 + GI * H3)
         
         'If (oper) Then     'for load plane
         Form6.Smith2.CircleIndex = 2
         Form6.Smith2.CircleM = Distan
         Form6.Smith2.CircleQ = angle
         Form6.Smith2.CircleR = radius     ' find a point on the circle
         
         Call CircleGamma(CSng(CenterR), CSng(CenterI), CSng(radius), CSng(LoadAngle), OPERATING)
 Else     ' AVAILABLE OR SOURCE PLANE CIRCLE
         SR = S11R
         SI = S11I
         CR = C1R
         CI = C1I
         H1 = Abs(Sqr(MulRe(S12R, S12I, S21R, S21I) ^ 2 + MulIm(S12R, S12I, S21R, S21I) ^ 2))
         H2 = Abs(Sqr(MulRe(SR, SI, SR, SI) ^ 2 + MulIm(SR, SI, SR, SI) ^ 2)) ' |S11*S11|
         H2A = Abs(Sqr(MulRe(S21R, S21I, S21R, S21I) ^ 2 + MulIm(S21R, S21I, S21R, S21I) ^ 2)) ' |S21*S21|
         H3 = H2 - DeltaM ^ 2
         GI = 10 ^ (G / 10) / H2A
         tmp = GI / (1 + GI * H3)
         AVCenterR = tmp * CR
         AVCenterI = tmp * -CI
         Distan = Abs(Sqr(AVCenterR ^ 2 + AVCenterI ^ 2))
         angle = Atn(AVCenterI / AVCenterR) * 180 / pi
         If (AVCenterR < 0 And AVCenterI > 0) Then angle = 180 + angle
         If (AVCenterR < 0 And AVCenterI < 0) Then angle = -180 + angle
         If (G < Gmax) Then RN = Sqr(1 - 2 * AK * H1 * GI + (H1 * GI) ^ 2) Else RN = 0
         AVradius = RN / (1 + GI * H3)
         
         Form6.Smith1.CircleIndex = 2
         Form6.Smith1.CircleM = Distan
         Form6.Smith1.CircleQ = angle
         Form6.Smith1.CircleR = AVradius
         Call CircleGamma(CSng(AVCenterR), CSng(AVCenterI), CSng(AVradius), AVLoadAngle, AVAILABLE)
         radius = AVradius ' for msg box
         End If
        If Form6.ShowMsgItem.Checked Then
        MsgBox "Center = " & Format(Distan, ".000") & " < " & Format(angle, ".000") & Chr$(13) & "Radius = " & Format(radius, ".000"), 0, "QuickSmith"
        End If
Exit Sub
GainCircError:
If (Err = 11) Then
    response% = MsgBox("'A Division By Zero' Error Was Trapped", 49, "Error")
    If response = 1 Then Resume Else Exit Sub
Else: Exit Sub
End If
        

End Sub

Function MulIm(a, B, C, D) As Single
' (A+Bj)* (C+Dj)
MulIm = (a * D) + (B * C)


End Function

Function MulRe(a, B, C, D) As Single
' (A+Bj)* (C+Dj)
 MulRe = (a * C) - (B * D)
End Function

Sub Noise()
Dim GOMAG, GOANG, RN, Fil, FminL As Single
Dim GoR, GoI, tmpR, tmpI, tmp, ani, cfimag, cfiang, RFi As Single
   On Error GoTo NoiseError
   
  ' Fmin = form10.Text1.Text
   GOMAG = Form10.Text2.Text
   GOANG = Form10.Text3.Text
   RN = Form10.text4.Text
   'fi = form10.Text5.Text
   
   If (fi < Fmin) Then
    fi = Fmin
    'form10.Text5.Text = fi
   End If
   
   Form6.Panel3D13.Caption = Format$(fi, "###0.000")
   
   FminL = 10 ^ (Fmin / 10)  ' dB to linear
   Fil = 10 ^ (fi / 10)
   If (Form6.Smith1.Z0 > 0) Then RN = RN / Form6.Smith1.Z0 Else RN = RN / 50
   GoR = GOMAG * Cos(GOANG * pi / 180)
   GoI = GOMAG * Sin(GOANG * pi / 180)
   tmpR = MulRe(AddRe(1, 0, GoR, GoI), AddIm(1, 0, GoR, GoI), AddRe(1, 0, GoR, GoI), AddIm(1, 0, GoR, GoI))
   tmpI = MulIm(AddRe(1, 0, GoR, GoI), AddIm(1, 0, GoR, GoI), AddRe(1, 0, GoR, GoI), AddIm(1, 0, GoR, GoI))
   tmp = Abs(Sqr(tmpR ^ 2 + tmpI ^ 2))
   If (RN <> 0) Then ani = (Fil - FminL) * tmp / (4 * RN)
   
   cfimag = GOMAG / (1 + ani)
   cfiang = GOANG
   If (Fil > FminL) Then RFi = Sqr((ani) ^ 2 + ani * (1 - GOMAG ^ 2)) / (1 + ani) Else RFi = 0
   'MsgBox cfimag & " " & cfiang & "  " & rfi
   If Form6.ShowMsgItem.Checked Then
   MsgBox "Center = " & Format(cfimag, ".000") & " < " & Format(GOANG, ".000") & Chr$(13) & "Radius = " & Format(RFi, ".000"), 0, "QuickSmith"
   End If
   Form6.Smith1.CircleIndex = 3
   Form6.Smith1.CircleM = cfimag
   Form6.Smith1.CircleQ = GOANG
   Form6.Smith1.CircleR = RFi
Exit Sub
NoiseError:
If (Err = 11) Then
    response% = MsgBox("'A Division By Zero' Error Was Trapped", 49, "Error")
    If response = 1 Then Resume Else Exit Sub
Else: Exit Sub
End If
        

End Sub

