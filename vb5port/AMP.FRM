VERSION 5.00
Object = "{B16553C3-06DB-101B-85B2-0000C009BE81}#1.0#0"; "SPIN32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5AE352A3-E9A4-11D0-936D-444553540000}#1.0#0"; "Smith.ocx"
Begin VB.Form Form6 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Amplifier Design/Analysis"
   ClientHeight    =   5715
   ClientLeft      =   1425
   ClientTop       =   4350
   ClientWidth     =   12165
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   Icon            =   "AMP.frx":0000
   LinkTopic       =   "Form6"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5715
   ScaleWidth      =   12165
   Begin Threed.SSPanel Panel3D1 
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   2985
      _Version        =   65536
      _ExtentX        =   5265
      _ExtentY        =   873
      _StockProps     =   15
      Caption         =   "SOURCE PLANE"
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
      Alignment       =   0
      Autosize        =   1
   End
   Begin Threed.SSPanel Panel3D2 
      Height          =   495
      Left            =   6360
      TabIndex        =   0
      Top             =   0
      Width           =   2385
      _Version        =   65536
      _ExtentX        =   4207
      _ExtentY        =   873
      _StockProps     =   15
      Caption         =   "LOAD PLANE"
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
      Alignment       =   0
      Autosize        =   1
   End
   Begin SmithLibCtl.Smith Smith2 
      Height          =   855
      Left            =   7560
      TabIndex        =   22
      Top             =   3120
      Width           =   975
      _Version        =   65536
      _ExtentX        =   1720
      _ExtentY        =   1508
      _StockProps     =   13
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin SmithLibCtl.Smith Smith1 
      Height          =   735
      Left            =   1080
      TabIndex        =   21
      Top             =   3360
      Width           =   735
      _Version        =   65536
      _ExtentX        =   1296
      _ExtentY        =   1296
      _StockProps     =   13
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSPanel Panel3D3 
      Align           =   2  'Align Bottom
      Height          =   1035
      Left            =   0
      TabIndex        =   2
      Top             =   4680
      Width           =   12165
      _Version        =   65536
      _ExtentX        =   21458
      _ExtentY        =   1826
      _StockProps     =   15
      ForeColor       =   -2147483640
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
      BevelOuter      =   0
      BevelInner      =   1
      Alignment       =   1
      Begin Spin.SpinButton Spin5 
         Height          =   372
         Left            =   11220
         TabIndex        =   16
         Top             =   300
         Width           =   372
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   73
         BorderThickness =   2
         Delay           =   200
         ShadowThickness =   1
         TdThickness     =   2
      End
      Begin Threed.SSPanel Panel3D14 
         Height          =   372
         Left            =   9960
         TabIndex        =   15
         Top             =   300
         Width           =   1272
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   15
         Caption         =   "VSWR: 1"
         ForeColor       =   4210752
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   16.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
         BevelOuter      =   1
         BevelInner      =   1
      End
      Begin Spin.SpinButton Spin4 
         Height          =   516
         Left            =   2628
         TabIndex        =   17
         Top             =   228
         Width           =   372
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   73
         BorderThickness =   2
         Delay           =   200
         ShadowThickness =   1
         TdThickness     =   2
      End
      Begin Threed.SSPanel Panel3D5 
         Height          =   450
         Left            =   120
         TabIndex        =   14
         Top             =   480
         Width           =   990
         _Version        =   65536
         _ExtentX        =   1746
         _ExtentY        =   794
         _StockProps     =   15
         Caption         =   "ZS:"
         ForeColor       =   -2147483640
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "System"
            Size            =   29.26
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         Autosize        =   1
      End
      Begin Threed.SSPanel Panel3D7 
         Height          =   450
         Left            =   6120
         TabIndex        =   13
         Top             =   510
         Width           =   945
         _Version        =   65536
         _ExtentX        =   1667
         _ExtentY        =   794
         _StockProps     =   15
         Caption         =   "ZL:"
         ForeColor       =   -2147483640
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "System"
            Size            =   29.26
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         Autosize        =   1
      End
      Begin Threed.SSPanel Panel3D6 
         Height          =   204
         Left            =   6120
         TabIndex        =   12
         Top             =   270
         Width           =   792
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   15
         Caption         =   "GammaL:"
         ForeColor       =   -2147483640
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   29.47
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         Autosize        =   1
      End
      Begin Spin.SpinButton Spin3 
         Height          =   528
         Left            =   8640
         TabIndex        =   18
         Top             =   225
         Width           =   372
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   73
         BorderThickness =   2
         Delay           =   200
         ShadowThickness =   1
         TdThickness     =   2
      End
      Begin Spin.SpinButton Spin2 
         Height          =   252
         Left            =   5280
         TabIndex        =   19
         Top             =   480
         Width           =   252
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   73
         Delay           =   200
         ShadowThickness =   1
         TdThickness     =   1
      End
      Begin Spin.SpinButton Spin1 
         Height          =   252
         Left            =   5280
         TabIndex        =   20
         Top             =   240
         Width           =   252
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   73
         Delay           =   200
         ShadowThickness =   1
         TdThickness     =   1
      End
      Begin Threed.SSPanel Panel3D13 
         Height          =   252
         Left            =   4200
         TabIndex        =   11
         Top             =   480
         Width           =   1092
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   15
         ForeColor       =   -2147483640
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   24.27
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
      End
      Begin Threed.SSPanel Panel3D12 
         Height          =   252
         Left            =   4200
         TabIndex        =   9
         Top             =   240
         Width           =   1092
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   15
         ForeColor       =   -2147483640
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   24.27
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
      End
      Begin Threed.SSPanel Panel3D11 
         Height          =   252
         Left            =   960
         TabIndex        =   7
         Top             =   480
         Width           =   1680
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   15
         ForeColor       =   -2147483640
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   24.27
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
      End
      Begin Threed.SSPanel Panel3D8 
         Height          =   252
         Left            =   960
         TabIndex        =   6
         Top             =   240
         Width           =   1668
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   15
         ForeColor       =   -2147483640
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   24.27
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
      End
      Begin Threed.SSPanel Panel3D9 
         Height          =   252
         Left            =   6960
         TabIndex        =   5
         Top             =   480
         Width           =   1680
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   15
         ForeColor       =   -2147483640
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   24.27
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
      End
      Begin Threed.SSPanel Panel3D10 
         Height          =   264
         Left            =   6960
         TabIndex        =   4
         Top             =   225
         Width           =   1680
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   15
         ForeColor       =   -2147483640
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   22.92
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
      End
      Begin Threed.SSPanel Panel3D4 
         Height          =   204
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   816
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   15
         Caption         =   "GammaS:"
         ForeColor       =   -2147483640
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   29.47
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         Autosize        =   1
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ga[dB]: "
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3480
         TabIndex        =   10
         Top             =   240
         Width           =   720
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fi [dB]: "
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3480
         TabIndex        =   8
         Top             =   495
         Width           =   735
      End
   End
   Begin MSComDlg.CommonDialog CMDialog1 
      Left            =   0
      Top             =   5760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FontSize        =   0
      MaxFileSize     =   256
   End
   Begin VB.Menu EnterMenu 
      Caption         =   "&Enter Parameters"
      Begin VB.Menu SparMenu 
         Caption         =   "&S-Parameters"
      End
      Begin VB.Menu NoiseMenu 
         Caption         =   "&Noise-Parameters"
      End
      Begin VB.Menu CharMenu 
         Caption         =   "&Char. Impedance"
      End
   End
   Begin VB.Menu CirclesMenu 
      Caption         =   "&Circles"
      Begin VB.Menu StabilityShowitem 
         Caption         =   "&Stability Circles"
         Enabled         =   0   'False
      End
      Begin VB.Menu OperItem 
         Caption         =   "&Operating Gain Circles"
         Enabled         =   0   'False
      End
      Begin VB.Menu AvItem 
         Caption         =   "&Available Gain Circles"
         Enabled         =   0   'False
      End
      Begin VB.Menu NoiseShowItem 
         Caption         =   "&Noise Circles"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu TransferMenu 
      Caption         =   "&Transfer"
      Begin VB.Menu SourceItem 
         Caption         =   "&Source Impedance"
      End
      Begin VB.Menu LoadItem 
         Caption         =   "&Load Impedance"
      End
   End
   Begin VB.Menu RedrawMenu 
      Caption         =   "&Redraw"
   End
   Begin VB.Menu OptionsMenu 
      Caption         =   "&Options"
      Begin VB.Menu ShowMsgItem 
         Caption         =   "Show &Message Box"
      End
      Begin VB.Menu SetColorItem 
         Caption         =   "Set &Color"
         Begin VB.Menu StabilitySetColorItem 
            Caption         =   "&Stability Circles"
         End
         Begin VB.Menu GainSetColorItem 
            Caption         =   "&Gain Circles"
         End
         Begin VB.Menu NoiseSetColorItem 
            Caption         =   "&Noise Circles"
         End
      End
   End
   Begin VB.Menu ExitMenu 
      Caption         =   "&Back"
   End
   Begin VB.Menu PrintMenu 
      Caption         =   "&Print"
   End
   Begin VB.Menu HelpMenu 
      Caption         =   "&Help"
      Begin VB.Menu ContentsItem 
         Caption         =   "&Contents"
      End
      Begin VB.Menu ExamplesItem 
         Caption         =   "&Examples"
      End
      Begin VB.Menu AboutHelpItem 
         Caption         =   "&About Help"
      End
      Begin VB.Menu AboutBar 
         Caption         =   "-"
      End
      Begin VB.Menu AboutItem 
         Caption         =   "A&bout..."
      End
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StepG, ThisG, StepAngle, StepN, AVStepangle As Single
Option Explicit

Private Sub AboutHelpItem_Click()
Dim dummy As String
Dim R
 R = WinHelp(hwnd, dummy, HELP_HELPONHELP, CLng(0))

End Sub

Private Sub AboutItem_Click()
AboutBox.Show 1  'modal
End Sub

Private Sub AvItem_Click()
Dim Gastr As String
If AvItem.Checked = False Then
    Gastr = InputBox$("Maximum Gain = " & Format(GPMAX, "0.000") & " dB " & Chr$(13) & Chr$(13) & "Enter Gain In dB : ", "QuickSmith", Format(CStr(ThisG), "0.000"))
           If (Gastr = "") Then GoTo EndGainA
           Gastr = Trim(Gastr)
           If (IsNumeric(Gastr)) Then ThisG = CDbl(Gastr)
           If (ThisG > GPMAX) Then ThisG = GPMAX
           Panel3D12.Caption = Format(ThisG, "0.000")
           Call GainCirc(CSng(ThisG), CSng(GPMAX), AVAILABLE)
           AvItem.Checked = True
Else
    AvItem.Checked = False
   Smith1.CircleIndex = 2
   Smith1.CircleQ = 0
   Smith1.CircleM = 0
   Smith1.CircleR = 0
   Smith1.DataM = 0
   Smith1.DataQ = 0

    Panel3D11.Caption = ""
    Panel3D8.Caption = ""
 If (OperItem.Checked = False) Then Panel3D12.Caption = ""

End If
EndGainA:


End Sub

Private Sub CharMenu_Click()
Dim ZZ0, R
R = InputBox$("Enter Characteristic Impedance of the Charts", "QuickSmith", CStr(Smith1.Z0))
If (IsNumeric(R) = False) Then GoTo EndZZ0
If (CDbl(R) <= 0) Then GoTo EndZZ0
ZZ0 = CDbl(R)
Smith1.Z0 = ZZ0
Smith2.Z0 = ZZ0
EndZZ0:

End Sub

Private Sub ContentsItem_Click()
Dim R
R = WinHelp(form1.hwnd, HelpFilePath, HELP_CONTEXT, CLng(1))
End Sub

Private Sub ExamplesItem_Click()
Dim R
R = WinHelp(form1.hwnd, HelpFilePath, HELP_CONTEXT, CLng(2))
End Sub

Private Sub ExitMenu_Click()
MDIForm1.Show
Form6.Hide

End Sub

Private Sub Form_Load()
Move 0, 0, Screen.Width, Screen.Height
Smith1.Move 0, 0, Width / 2, 0.8 * (Height - Panel3D3.Height)
Smith2.Move Width / 2, 0, Width / 2, 0.8 * (Height - Panel3D3.Height)
Form6.DrawWidth = 1
Panel3D1.Left = 0
Panel3D2.Left = Smith1.Width
Smith1.CircleIndex = 1
Smith1.CircleColor = RGB(255, 0, 0)
Smith2.CircleIndex = 1
Smith2.CircleColor = RGB(255, 0, 0)

Smith1.CircleIndex = 2
Smith1.CircleColor = RGB(0, 255, 0)
Smith2.CircleIndex = 2
Smith2.CircleColor = RGB(0, 255, 0)

Smith1.CircleIndex = 3
'Smith1.CircleColor = RGB(0, 0, 255)
Smith1.CircleColor = RGB(128, 0, 0)
Smith2.CircleIndex = 3
'smith2.CircleColor = RGB(0, 0, 255)
Smith2.CircleColor = RGB(128, 0, 0)
Smith2.CircleIndex = 0
Smith1.CircleIndex = 0

Panel3D3.Align = 2
STFLAG = 0
StepG = 0.1
ThisG = 10
StepN = 0.1
StepAngle = 1
LoadAngle = 45
AVLoadAngle = 45
AVStepangle = 1
SPAROPENFLAG = False
NOISEOPENFLAG = False
AMPOPENFLAG = True
vswrstep1 = 1
End Sub

Private Sub GainOnItem_Click()
 
End Sub

Private Sub GainSetColorItem_Click()
On Error Resume Next
CMDialog1.CancelError = -1
Smith1.CircleIndex = 2
Smith2.CircleIndex = 2
CMDialog1.Color = Smith1.CircleColor
CMDialog1.DialogTitle = "Set Color"
CMDialog1.Flags = &H1 ' show initial value
CMDialog1.Action = 3
If Err = 0 Then
    Smith1.CircleColor = CMDialog1.Color
    Smith2.CircleColor = CMDialog1.Color
End If
End Sub

Private Sub LoadItem_Click()
 Dim prompt
If Panel3D9.Caption = "" Then
Beep
Exit Sub
Else
prompt = MsgBox("Load Impedance in Schematic will be Set to :" & Chr(13) & " R = " & Format(LoadReal, "0.000") & " Ohms " & Chr(13) & " X = " & Format(-LoadImag, "0.000") & " Ohms ", 1, "QuickSmith")
If (prompt = 2) Then Exit Sub
ZR(1) = LoadReal
ZI(1) = -LoadImag
If (Abs(ZR(1)) > 99999.99) Then form2.Panel3D4.Caption = Format$(ZR(1), "##.#E+##") Else form2.Panel3D4.Caption = Format$(ZR(1), "0.000")
If (Abs(ZR(1)) > 99999.99) Then form2.Panel3D5.Caption = Format$(ZI(1), "##.#E+##") Else form2.Panel3D5.Caption = Format$(ZI(1), "0.000")
If (form2.ImpeItem.Checked = False) Then
    form2.Label1(1) = "R"
    form2.Label1(2) = "X"
    form2.ImpeItem.Checked = True
End If
Call Zcalsweep(0)
End If
End Sub

Private Sub NoiseMenu_Click()
Form10.Show
End Sub

Private Sub NoiseSetColorItem_Click()
On Error Resume Next
CMDialog1.CancelError = -1
Smith1.CircleIndex = 3
Smith2.CircleIndex = 3
CMDialog1.Color = Smith1.CircleColor
CMDialog1.DialogTitle = "Set Color"
CMDialog1.Flags = &H1 ' show initial value
CMDialog1.Action = 3
If Err = 0 Then
    Smith1.CircleColor = CMDialog1.Color
    Smith2.CircleColor = CMDialog1.Color
End If
End Sub

Private Sub NoiseShowItem_Click()
Dim Gastr
If NoiseShowItem.Checked = False Then
   
Gastr = InputBox$("Minimum Noise Figure = " & Format(Fmin, "0.000") & " dB " & Chr$(13) & Chr$(13) & "Enter Noise Figure in dB : ", "QuickSmith", Format(CStr(fi), "0.000"))
If (Gastr = "") Then GoTo EndNoise
Gastr = Trim(Gastr)
If (IsNumeric(Gastr)) Then fi = CDbl(Gastr)
If (Fmin > fi) Then fi = Fmin
Panel3D13.Caption = Format(fi, "0.000")

Call Noise
NoiseShowItem.Checked = True
'form10.Show
Else
NoiseShowItem.Checked = False
   Smith1.CircleIndex = 3
   Smith1.CircleQ = 0
   Smith1.CircleM = 0
   Smith1.CircleR = 0
   Panel3D13.Caption = ""
End If
EndNoise:
End Sub

Private Sub OperItem_Click()

Dim Gastr As String
If OperItem.Checked = False Then
   
           Gastr = InputBox$("Maximum Gain = " & Format(GPMAX, "0.000") & " dB " & Chr$(13) & Chr$(13) & "Enter Gain In dB : ", "QuickSmith", Format(CStr(ThisG), "0.000"))
           If (Gastr = "") Then GoTo EndGain
           Gastr = Trim(Gastr)
           If (IsNumeric(Gastr)) Then ThisG = CDbl(Gastr)
           If (ThisG > GPMAX) Then ThisG = GPMAX
           Panel3D12.Caption = Format(ThisG, "0.000")
           Call GainCirc(CSng(ThisG), GPMAX, OPERATING)
            OperItem.Checked = True
Else
    OperItem.Checked = False
   Smith2.CircleIndex = 2
   Smith2.CircleQ = 0
   Smith2.CircleM = 0
   Smith2.CircleR = 0
   Smith2.DataM = 0
   Smith2.DataQ = 0
    Panel3D10.Caption = ""
    Panel3D9.Caption = ""
 If (AvItem.Checked = False) Then Panel3D12.Caption = ""
End If
EndGain:

End Sub

Private Sub Panel3D10_DblClick()
Dim stepsize
stepsize = InputBox$("Enter Step Size", "QuickSmith", CStr(StepAngle))
If (IsNumeric(stepsize)) Then StepAngle = Abs(CDbl(stepsize))

End Sub

Private Sub Panel3D11_DblClick()
Dim stepsize
stepsize = InputBox$("Enter Step Size", "QuickSmith", CStr(AVStepangle))
If (IsNumeric(stepsize)) Then AVStepangle = Abs(CDbl(stepsize))

End Sub

Private Sub Panel3D12_DblClick()
Dim stepsize
stepsize = InputBox$("Enter Step Size", "QuickSmith", CStr(StepG))
If (IsNumeric(stepsize)) Then StepG = Abs(CDbl(stepsize))

End Sub

Private Sub Panel3D13_DblClick()
Dim stepsize
stepsize = InputBox$("Enter Step Size", "QuickSmith", CStr(StepN))
If (IsNumeric(stepsize)) Then StepN = CDbl(stepsize)

End Sub

Private Sub Panel3D14_DblClick()
Dim Title As String
Dim prompt As String
Dim Default As String
Dim stepsize
Title = "Amplifier Design"
prompt = "Enter Step Size for VSWR Circles"
Default = CStr(vswrstep1)
stepsize = InputBox$(prompt$, Title$, Default$)
If (IsNumeric(stepsize)) Then vswrstep1 = CDbl(stepsize)



End Sub

Private Sub Panel3D8_DblClick()
Dim stepsize
stepsize = InputBox$("Enter Step Size", "QuickSmith", CStr(AVStepangle))
If (IsNumeric(stepsize)) Then AVStepangle = Abs(CDbl(stepsize))

End Sub

Private Sub Panel3D9_DblClick()
Dim stepsize
stepsize = InputBox$("Enter Step Size", "QuickSmith", CStr(StepAngle))
If (IsNumeric(stepsize)) Then StepAngle = Abs(CDbl(stepsize))

End Sub

Private Sub PrintMenu_Click()
    On Error GoTo PrintSetupError
    CMDialog1.CancelError = -1
    CMDialog1.Flags = &H40& 'PD_PRINTSETUP     ' Just enable print setup
    CMDialog1.Action = 5 ' DLG_PRINT        ' Show printer setup dialog box
    
    If Err = 0 Then
   ' print form
    Call PrintClient256(Form6)
    End If

PrintSetupExit:
Exit Sub

PrintSetupError:
    If Err <> CDERR_CANCEL Then         ' If user didn't select CANCEL
    '    Msg$ = "Error" & Str$(Err) & ": " & Error$
     '   MsgBox Msg$, 48, "Error"
    End If
    Resume PrintSetupExit

  

End Sub

Private Sub RedrawMenu_click()
  Smith1.Redraw = True
  Smith2.Redraw = True
End Sub

Private Sub ShowMsgItem_Click()
If ShowMsgItem.Checked = False Then ShowMsgItem.Checked = True Else ShowMsgItem.Checked = False
End Sub

Private Sub Smith1_ClickIn(m As Single, Q As Single)
Dim str1 As String
Dim Swr As Double
Dim S11 As Double
On Error GoTo ClickinError
If (m = 0) Then m = 1E-36
If (m = 1) Then m = 0.9999999999
Swr = (1 + m) / (1 - m)
S11 = -20 * (Log(m) / Log(10))
If (S11 >= 720) Then str1 = " > " Else str1 = ""
MsgBox ("Gamma [Mag]: " & Format(m, "#0.000") & Chr$(13) & "Gamma [Ang]: " & Format(Q, "###.00") & Chr$(176) & Chr$(13) & "Z [Real]: " & Format(AmpGTOZR(m, Q, CSng(Form6.Smith1.Z0)), "0.000") & Chr$(13) & "Z [Imag]: " & Format(AmpGTOZI(m, Q, CSng(Form6.Smith1.Z0)), "0.000")) & Chr(10) & "VSWR: " & Format(Swr, "###.00") & Chr(10) & "Return Loss: " & str1 & Format(S11, "##0.00") & " dB", , "QuickSmith"
Exit Sub
ClickinError:
If (Err = 11) Then
response% = MsgBox("'A Division By Zero' Error Was Trapped", 49, "Error")
If response = 1 Then Resume Else Exit Sub
End If



End Sub

Private Sub Smith1_MouseMove(m As Single, Q As Single)
'ZZ If (m > 1.5) Then Smith1.MousePointer = 0 Else Smith1.MousePointer = 10
End Sub

Private Sub Smith2_ClickIn(m As Single, Q As Single)
Dim str1 As String
Dim Swr As Double
Dim S11 As Double
On Error GoTo ClickinError1
If (m = 0) Then m = 1E-36
If (m = 1) Then m = 0.9999999999
Swr = (1 + m) / (1 - m)
S11 = -20 * (Log(m) / Log(10))
If (S11 >= 720) Then str1 = " > " Else str1 = ""

MsgBox ("Gamma [Mag]: " & Format(m, "#0.000") & Chr$(13) & "Gamma [Ang]: " & Format(Q, "###.00") & Chr$(176) & Chr$(13) & "Z [Real]: " & Format(AmpGTOZR(m, Q, CSng(Smith2.Z0)), "0.000") & Chr$(13) & "Z [Imag]: " & Format(AmpGTOZI(m, Q, CSng(Smith2.Z0)), "0.000")) & Chr(10) & "VSWR: " & Format(Swr, "###.00") & Chr(10) & "Return Loss: " & str1 & Format(S11, "##0.00") & " dB", , "QuickSmith"
Exit Sub
ClickinError1:
If (Err = 11) Then
response% = MsgBox("'A Division By Zero' Error Was Trapped", 49, "Error")
If response = 1 Then Resume Else Exit Sub
End If

End Sub

Private Sub Smith2_MouseMove(m As Single, Q As Single)
'ZZ If (m > 1.5) Then Smith2.MousePointer = 0 Else Smith2.MousePointer = 10

End Sub

Private Sub SourceItem_Click()
Dim prompt
If Panel3D11.Caption = "" Then
Beep
Exit Sub
Else
prompt = MsgBox("Load Impedance in Schematic will be Set to :" & Chr(13) & " R = " & Format(SourceReal, "0.000") & " Ohms " & Chr(13) & " X = " & Format(-SourceImag, "0.000") & " Ohms ", 1, "QuickSmith")
If (prompt = 2) Then Exit Sub
ZR(1) = SourceReal
ZI(1) = -SourceImag
If (Abs(ZR(1)) > 99999.99) Then form2.Panel3D4.Caption = Format$(ZR(1), "##.#E+##") Else form2.Panel3D4.Caption = Format$(ZR(1), "0.000")
If (Abs(ZR(1)) > 99999.99) Then form2.Panel3D5.Caption = Format$(ZI(1), "##.#E+##") Else form2.Panel3D5.Caption = Format$(ZI(1), "0.000")
If (form2.ImpeItem.Checked = False) Then
    form2.Label1(1) = "R"
    form2.Label1(2) = "X"
    form2.ImpeItem.Checked = True
End If
Call Zcalsweep(0)
End If
End Sub

Private Sub SparMenu_Click()
Form9.Show
End Sub

Private Sub Spin1_SpinDown()
If (Panel3D12.Caption <> "") Then
  If ((ThisG - StepG) > 0) Then ThisG = ThisG - StepG
  If (ThisG > GPMAX) Then ThisG = GPMAX
  Panel3D12.Caption = Format$(ThisG, "###0.000")
  
  If OperItem.Checked Then Call GainCirc(CSng(ThisG), GPMAX, OPERATING)
  If AvItem.Checked Then Call GainCirc(CSng(ThisG), GPMAX, AVAILABLE)
End If

End Sub

Private Sub Spin1_SpinUp()
If (Panel3D12.Caption <> "") Then
  If ((ThisG + StepG) > 0) Then ThisG = ThisG + StepG Else ThisG = 0.000000000001
  If (ThisG > GPMAX) Then ThisG = GPMAX
  Panel3D12.Caption = Format$(ThisG, "###0.000")
  'If (freq > 9999.99) Then Panel3D10.Caption = Format$(freq, "##.#E+##")
  If OperItem.Checked Then Call GainCirc(CSng(ThisG), GPMAX, OPERATING)
  If AvItem.Checked Then Call GainCirc(CSng(ThisG), GPMAX, AVAILABLE)
End If

End Sub

Private Sub Spin2_SpinDown()
If (Panel3D13.Caption <> "") Then
  If ((fi - StepN) > 0) Then fi = fi - StepN
  'If (Fi < Fmin) Then Fi = Fmin
  'panel3d13.Caption = Format$(Fi, "###0.000")
  'form10.Text5.Text = Format$(Fi, "###0.000")
  Call Noise
  End If

End Sub

Private Sub Spin2_SpinUp()
If (Panel3D13.Caption <> "") Then
  If ((fi + StepN) > 0) Then fi = fi + StepN
  'If (Fi < Fmin) Then Fi = Fmin
  'panel3d13.Caption = Format$(Fi, "###0.000")
  'form10.Text5.Text = Format$(Fi, "###0.000")
  Call Noise
  End If


End Sub

Private Sub Spin3_SpinDown()
If (OperItem.Checked) Then
  LoadAngle = (LoadAngle - StepAngle)
  Call CircleGamma(CSng(CenterR), CSng(CenterI), CSng(radius), CSng(LoadAngle), OPERATING)

 
End If

End Sub

Private Sub Spin3_SpinUp()
If (OperItem.Checked) Then
LoadAngle = (LoadAngle + StepAngle)
  Call CircleGamma(CSng(CenterR), CSng(CenterI), CSng(radius), CSng(LoadAngle), OPERATING)
  
 
End If

End Sub

Private Sub Spin4_SpinDown()
If (AvItem.Checked) Then
  AVLoadAngle = (AVLoadAngle - AVStepangle)
  Call CircleGamma(CSng(AVCenterR), CSng(AVCenterI), CSng(AVradius), CSng(AVLoadAngle), AVAILABLE)
  'MsgBox CStr(AVLoadangle)
  End If


End Sub

Private Sub Spin4_SpinUp()
If (AvItem.Checked) Then
  AVLoadAngle = (AVLoadAngle + AVStepangle)
  Call CircleGamma(CSng(AVCenterR), CSng(AVCenterI), CSng(AVradius), CSng(AVLoadAngle), AVAILABLE)
   'MsgBox AVStepangle
  End If


End Sub

Private Sub Spin5_SpinDown()
Dim vswrtext, start
Dim vswr As Single
vswrtext = Panel3D14.Caption
start = InStr(vswrtext, "VSWR: ")
Mid(vswrtext, start, 6) = "      "
vswr = vswrtext
vswr = vswr - vswrstep1
If (vswr < 1) Then vswr = 1
'Smith1.Redraw = True
Smith1.VSWRCircle = vswr
'smith2.Redraw = True
Smith2.VSWRCircle = vswr
Panel3D14.Caption = "VSWR: " & Format(vswr, ".0#")
 
End Sub

Private Sub Spin5_SpinUp()
Dim vswrtext, start
Dim vswr As Single
vswrtext = Panel3D14.Caption
start = InStr(vswrtext, "VSWR: ")
Mid(vswrtext, start, 6) = "      "
vswr = vswrtext
vswr = vswr + vswrstep1
If (vswr < 1) Then vswr = 1
'Smith1.Redraw = True
Smith1.VSWRCircle = vswr
'smith2.Redraw = True
Smith2.VSWRCircle = vswr
Panel3D14.Caption = "VSWR: " & Format(vswr, ".0#")


End Sub

Private Sub StabilitySetColorItem_Click()
On Error Resume Next
CMDialog1.CancelError = -1
Smith1.CircleIndex = 1
Smith2.CircleIndex = 1
CMDialog1.Color = Smith1.CircleColor
CMDialog1.DialogTitle = "Set Color"
CMDialog1.Flags = &H1 ' show initial value
CMDialog1.Action = 3
If Err = 0 Then
    Smith1.CircleColor = CMDialog1.Color
    Smith2.CircleColor = CMDialog1.Color
End If
End Sub

Private Sub StabilityShowitem_Click()
Dim str1, Str2 As String
If StabilityShowitem.Checked = False Then
   StabilityShowitem.Checked = True
   
  'Input Stability Circle
  
   Smith1.CircleIndex = 1
   Smith1.CircleQ = CSA
   Smith1.CircleM = CSM
   Smith1.CircleR = RS
   If Form6.ShowMsgItem.Checked Then
   MsgBox "Center = " & Format(CSM, ".000") & " < " & Format(CSA, ".000") & Chr$(13) & "Radius = " & Format(RS, ".000"), 0, "QuickSmith"
   End If
 
   'Output Stability Circle
   Smith2.CircleIndex = 1
   Smith2.CircleQ = CLA
   Smith2.CircleM = CLM
   Smith2.CircleR = RL
   If Form6.ShowMsgItem.Checked Then
   MsgBox "Center = " & Format(CLM, ".000") & " < " & Format(CLA, ".000") & Chr$(13) & "Radius = " & Format(RL, ".000"), 0, "QuickSmith"
   End If
   
   If (Sqr(S11R ^ 2 + S11I ^ 2) < 1) Then
    If (CSM > RS) Then str1 = "The Stability Region is Outside the Circle" Else str1 = "The Stability Region is Inside the Circle"
    If (CLM > RL) Then Str2 = "The Stability Region is Outside the Circle" Else Str2 = "The Stability Region is Inside the Circle"
   Else
    If (CSM > RS) Then str1 = "The Stability Region is Inside the Circle" Else str1 = "The Stability Region is Outside the Circle"
    If (CLM > RL) Then Str2 = "The Stability Region is Inside the Circle" Else Str2 = "The Stability Region is Outside the Circle"
   End If
   If (STFLAG = False) Then MsgBox "SOURCE PLANE : " & str1 & Chr$(13) & Chr$(13) & "LOAD PLANE : " & Str2, 0, "QuickSmith"

Else
   StabilityShowitem.Checked = False
   Smith1.CircleIndex = 1
   Smith1.CircleQ = 0
   Smith1.CircleM = 0
   Smith1.CircleR = 0

   Smith2.CircleIndex = 1
   Smith2.CircleQ = 0
   Smith2.CircleM = 0
   Smith2.CircleR = 0
End If

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

