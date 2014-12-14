VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form Form9 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Enter S-Parameters"
   ClientHeight    =   3315
   ClientLeft      =   1665
   ClientTop       =   2115
   ClientWidth     =   4680
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "SPAR.frx":0000
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3315
   ScaleWidth      =   4680
   Begin Threed.SSPanel Panel3D1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   18
      Top             =   2940
      Width           =   4680
      _Version        =   65536
      _ExtentX        =   8255
      _ExtentY        =   661
      _StockProps     =   15
      ForeColor       =   255
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   2
      BevelOuter      =   1
      BevelInner      =   1
      Autosize        =   1
   End
   Begin VB.OptionButton Option2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Rect."
      ForeColor       =   &H80000008&
      Height          =   288
      Left            =   3600
      TabIndex        =   17
      Top             =   240
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Polar"
      ForeColor       =   &H80000008&
      Height          =   288
      Left            =   2640
      TabIndex        =   16
      Top             =   240
      Value           =   -1  'True
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      Caption         =   "Stability"
      Height          =   396
      Left            =   3000
      TabIndex        =   15
      Top             =   960
      Width           =   972
   End
   Begin VB.TextBox Text8 
      Appearance      =   0  'Flat
      Height          =   288
      Left            =   1440
      TabIndex        =   13
      Text            =   "-95.7"
      Top             =   1560
      Width           =   732
   End
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      Height          =   288
      Left            =   720
      TabIndex        =   12
      Text            =   ".572"
      Top             =   1560
      Width           =   732
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      Height          =   288
      Left            =   1440
      TabIndex        =   10
      Text            =   "16.3"
      Top             =   1200
      Width           =   732
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      Height          =   288
      Left            =   720
      TabIndex        =   9
      Text            =   ".057"
      Top             =   1200
      Width           =   732
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      Height          =   288
      Left            =   1440
      TabIndex        =   7
      Text            =   "28.5"
      Top             =   840
      Width           =   732
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   288
      Left            =   720
      TabIndex        =   6
      Text            =   "2.058"
      Top             =   840
      Width           =   732
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   288
      Left            =   1440
      TabIndex        =   2
      Text            =   "-171.3"
      Top             =   480
      Width           =   732
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   288
      Left            =   720
      TabIndex        =   1
      Text            =   ".641"
      Top             =   480
      Width           =   732
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "O K"
      Height          =   396
      Left            =   3000
      TabIndex        =   0
      Top             =   1560
      Width           =   972
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Delta = "
      ForeColor       =   &H00FF0000&
      Height          =   252
      Left            =   672
      TabIndex        =   20
      Top             =   2244
      Width           =   1512
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "K ="
      ForeColor       =   &H00FF0000&
      Height          =   264
      Left            =   672
      TabIndex        =   19
      Top             =   1980
      Width           =   1512
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "S21"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   120
      TabIndex        =   14
      Top             =   840
      Width           =   396
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "S22"
      ForeColor       =   &H80000008&
      Height          =   276
      Left            =   120
      TabIndex        =   11
      Top             =   1560
      Width           =   396
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "S12"
      ForeColor       =   &H80000008&
      Height          =   276
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   396
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "S11"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   396
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Ang."
      ForeColor       =   &H80000008&
      Height          =   276
      Left            =   1524
      TabIndex        =   4
      Top             =   144
      Width           =   396
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Mag."
      ForeColor       =   &H80000008&
      Height          =   276
      Left            =   720
      TabIndex        =   3
      Top             =   132
      Width           =   396
   End
   Begin VB.Menu FileItem 
      Caption         =   ""
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim DeltaR, DeltaI, KR, KI, CL, CS, DM, DA  As Single
Option Explicit

Private Sub Command1_Click()
Form9.Hide
'MdiForm1.Show
End Sub

Private Sub Command2_Click()
Dim S11M, S11A, S21M, S21A, S12M, S12A, S22M, S22A As Single
Dim H1, H2, H3, RLD, CLR, CLI, RSD, CSR, CSI As Single
Dim GMSG As Single
On Error GoTo Command2Error
If (IsNumeric(Text1.Text) And IsNumeric(Text2.Text) And IsNumeric(Text3.Text) And IsNumeric(Text4.Text) And IsNumeric(Text5.Text) And IsNumeric(Text6.Text) And IsNumeric(Text7.Text) And IsNumeric(Text8.Text)) Then
   S11M = Text1.Text
   S11A = Text2.Text
   S21M = Text3.Text
   S21A = Text4.Text
   S12M = Text5.Text
   S12A = Text6.Text
   S22M = Text7.Text
   S22A = Text8.Text
   If (Option1.Value) Then   ' for Polar input
   S11R = S11M * Cos(S11A * pi / 180)
   S11I = S11M * Sin(S11A * pi / 180)
   S21R = S21M * Cos(S21A * pi / 180)
   S21I = S21M * Sin(S21A * pi / 180)
   S12R = S12M * Cos(S12A * pi / 180)
   S12I = S12M * Sin(S12A * pi / 180)
   S22R = S22M * Cos(S22A * pi / 180)
   S22I = S22M * Sin(S22A * pi / 180)
   End If
  If (Option2.Value) Then   ' for Rect input
   S11R = S11M
   S11I = S11A
   S21R = S21M
   S21I = S21A
   S12R = S12M
   S12I = S12A
   S22R = S22M
   S22I = S22A
  End If
   DeltaR = AddRe(MulRe(S11R, S11I, S22R, S22I), MulIm(S11R, S11I, S22R, S22I), -MulRe(S12R, S12I, S21R, S21I), -MulIm(S12R, S12I, S21R, S21I))
   DeltaI = AddIm(MulRe(S11R, S11I, S22R, S22I), MulIm(S11R, S11I, S22R, S22I), -MulRe(S12R, S12I, S21R, S21I), -MulIm(S12R, S12I, S21R, S21I))
   DeltaM = Abs(Sqr(DeltaR ^ 2 + DeltaI ^ 2))
   Label7.Caption = "Delta : " & Format(DeltaM, "0.000")
' Calculate K
   H1 = Abs(Sqr(MulRe(S21R, S21I, S12R, S12I) ^ 2 + MulIm(S21R, S21I, S12R, S12I) ^ 2))
      
      If (H1 <> 0) Then
        H2 = Abs(Sqr(MulRe(S11R, S11I, S11R, S11I) ^ 2 + MulIm(S11R, S11I, S11R, S11I) ^ 2)) ' |S11*S11|
        H3 = Abs(Sqr(MulRe(S22R, S22I, S22R, S22I) ^ 2 + MulIm(S22R, S22I, S22R, S22I) ^ 2)) ' |S12*S12|
        AK = (1 - H2 - H3 + DeltaM ^ 2) / (2 * H1)
      Else: AK = 1E+20
      End If
   Label8.Caption = "K : " & Format(AK, "0.000")
If (DeltaM < 1 And AK > 1) Then
 Panel3D1.Caption = "UNCONDITIONALLY STABLE TRANSISTOR"
   STFLAG = True
Else
   Panel3D1.Caption = "POTENTIALLY UNSTABLE TRANSISTOR"
   STFLAG = False
End If
'Calculate Global Variable C1 & C2
C1R = AddRe(S11R, S11I, -MulRe(DeltaR, DeltaI, S22R, -S22I), -MulIm(DeltaR, DeltaI, S22R, -S22I))
C1I = AddIm(S11R, S11I, -MulRe(DeltaR, DeltaI, S22R, -S22I), -MulIm(DeltaR, DeltaI, S22R, -S22I))
C2R = AddRe(S22R, S22I, -MulRe(DeltaR, DeltaI, S11R, -S11I), -MulIm(DeltaR, DeltaI, S11R, -S11I))
C2I = AddIm(S22R, S22I, -MulRe(DeltaR, DeltaI, S11R, -S11I), -MulIm(DeltaR, DeltaI, S11R, -S11I))
End If
RLD = S22R ^ 2 + S22I ^ 2 - DeltaM ^ 2
If (RLD = 0) Then RLD = 0.000000000001
RL = H1 / RLD
CLR = C2R / RLD
CLI = -C2I / RLD
CLM = Abs(Sqr(CLR ^ 2 + CLI ^ 2))
CLA = Atn(CLI / CLR) * 180 / pi
  If (CLR < 0 And CLI > 0) Then CLA = 180 + CLA
  If (CLR < 0 And CLI < 0) Then CLA = -180 + CLA
   
RSD = S11R ^ 2 + S11I ^ 2 - DeltaM ^ 2
If (RSD = 0) Then RSD = 0.000000000001
RS = H1 / RSD
CSR = C1R / RSD
CSI = -C1I / RSD
CSM = Abs(Sqr(CSR ^ 2 + CSI ^ 2))
CSA = Atn(CSI / CSR) * 180 / pi
  If (CSR < 0 And CSI > 0) Then CSA = 180 + CSA
  If (CSR < 0 And CSI < 0) Then CSA = -180 + CSA

' Calculate Maximum Gain  GPMAX
    GMSG = Abs(Sqr(DivRe(S21R, S21I, S12R, S12I) ^ 2 + DivIm(S21R, S21I, S12R, S12I) ^ 2))
    If (STFLAG = False) Then  ' Unstable case
        GPMAX = (Log(GMSG) / Log(10)) * 10
    Else
        GPMAX = (AK - Sqr(AK ^ 2 - 1)) * GMSG
        GPMAX = (Log(GPMAX) / Log(10)) * 10
    End If

Form6.OperItem.Enabled = True
Form6.AvItem.Enabled = True
Form6.StabilityShowitem.Enabled = True
'form6.CirclesMenu.Enabled = True
Form6.StabilityShowitem.Checked = False
Form6.NoiseShowItem.Checked = False
Form6.OperItem.Checked = False
Form6.AvItem.Checked = False

Exit Sub
Command2Error:
If (Err = 11) Then
    response% = MsgBox("'A Division By Zero' Error Was Trapped", 49, "Error")
    If response = 1 Then Resume Else Exit Sub
Else: Exit Sub
End If

End Sub

Private Sub Form_Load()
 Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
 Form9.Show
 SPAROPENFLAG = True
 Text1.SetFocus
'

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then ' Enter
   SendKeys " {Tab}"
   KeyAscii = 0
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then ' Enter
   SendKeys " {Tab}"
   KeyAscii = 0
End If


End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then ' Enter
   SendKeys " {Tab}"
   KeyAscii = 0
End If

End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then ' Enter
   SendKeys " {Tab}"
   KeyAscii = 0
End If

End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then ' Enter
   SendKeys " {Tab}"
   KeyAscii = 0
End If

End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then ' Enter
   SendKeys " {Tab}"
   KeyAscii = 0
End If


End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then ' Enter
   SendKeys " {Tab}"
   KeyAscii = 0
End If

End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then ' Enter
   SendKeys " {Tab}"
   KeyAscii = 0
End If

End Sub

