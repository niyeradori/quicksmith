VERSION 5.00
Object = "{B16553C3-06DB-101B-85B2-0000C009BE81}#1.0#0"; "SPIN32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5AE352A3-E9A4-11D0-936D-444553540000}#1.0#0"; "Smith.ocx"
Begin VB.Form form1 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   Caption         =   "SmithChart"
   ClientHeight    =   6105
   ClientLeft      =   1125
   ClientTop       =   2160
   ClientWidth     =   9600
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
   Icon            =   "SC0.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "SC0.frx":030A
   ScaleHeight     =   6105
   ScaleWidth      =   9600
   Begin SmithLibCtl.Smith Smith1 
      Height          =   735
      Left            =   2040
      TabIndex        =   18
      Top             =   1560
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
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
   Begin Threed.SSFrame Frame3D1 
      Height          =   600
      Left            =   -36
      TabIndex        =   1
      Top             =   -108
      Width           =   9900
      _Version        =   65536
      _ExtentX        =   17462
      _ExtentY        =   1058
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin Threed.SSCommand Command3D2 
         Height          =   405
         Left            =   6600
         TabIndex        =   14
         ToolTipText     =   "Amplifier Design"
         Top             =   108
         Width           =   465
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   78
         BevelWidth      =   1
         Outline         =   0   'False
         MouseIcon       =   "SC0.frx":074C
         Picture         =   "SC0.frx":0B9E
      End
      Begin Threed.SSCommand Command3D1 
         Height          =   405
         Left            =   7092
         TabIndex        =   13
         ToolTipText     =   "Component Values "
         Top             =   108
         Width           =   465
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   78
         BevelWidth      =   1
         Outline         =   0   'False
         Picture         =   "SC0.frx":0FF0
      End
      Begin Spin.SpinButton Spin1 
         Height          =   372
         Left            =   4380
         TabIndex        =   16
         Top             =   120
         Width           =   372
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   73
         ForeColor       =   14737632
         BackColor       =   12632256
         ShadowBackColor =   12632256
         TdThickness     =   3
      End
      Begin Spin.SpinButton Spin2 
         Height          =   372
         Left            =   2880
         TabIndex        =   17
         Top             =   120
         Width           =   384
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   73
         ForeColor       =   4210752
         BackColor       =   12632256
         ShadowBackColor =   12632256
         TdThickness     =   3
      End
      Begin Threed.SSPanel Panel3D2 
         Height          =   372
         Left            =   1680
         TabIndex        =   0
         Top             =   120
         Width           =   1212
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
      Begin Threed.SSPanel Panel3D3 
         Height          =   372
         Left            =   4770
         TabIndex        =   5
         Top             =   120
         Width           =   1824
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   15
         Caption         =   "Z:"
         ForeColor       =   16711680
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   16.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
         BevelOuter      =   1
         BevelInner      =   1
      End
      Begin Threed.SSPanel Panel3D1 
         Height          =   372
         Left            =   3300
         TabIndex        =   4
         Top             =   120
         Width           =   1068
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   15
         Caption         =   "Q: 0"
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
      Begin Threed.SSOption Option3D1 
         Height          =   195
         Left            =   960
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   240
         Width           =   735
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   78
         Caption         =   "Y-In"
         ForeColor       =   16448
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   29.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSOption Option3D2 
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   765
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   78
         Caption         =   "Z-In"
         ForeColor       =   16448
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   29.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   -1  'True
      End
   End
   Begin Threed.SSPanel Panel3D4 
      Height          =   732
      Left            =   0
      TabIndex        =   15
      Top             =   432
      Width           =   1776
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   15
      ForeColor       =   4210752
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      BevelOuter      =   0
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "test"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   244
         Left            =   324
         TabIndex        =   6
         Top             =   478
         Width           =   974
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "test"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   324
         TabIndex        =   7
         Top             =   226
         Width           =   972
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Re"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   192
         Left            =   72
         TabIndex        =   8
         Top             =   252
         Width           =   216
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Im"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   192
         Left            =   72
         TabIndex        =   9
         Top             =   478
         Width           =   168
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Ohms"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   192
         Left            =   1332
         TabIndex        =   10
         Top             =   360
         Width           =   420
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Z-In"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   570
         TabIndex        =   11
         Top             =   30
         Width           =   360
      End
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   260
      LargeChange     =   500
      Left            =   360
      Max             =   5000
      SmallChange     =   50
      TabIndex        =   12
      Top             =   5430
      Visible         =   0   'False
      Width           =   2208
   End
   Begin MSComDlg.CommonDialog CMDialog2 
      Left            =   7800
      Top             =   5640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FontSize        =   0
      MaxFileSize     =   256
   End
   Begin MSComDlg.CommonDialog CMDialog1 
      Left            =   0
      Top             =   7920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FontSize        =   0
      MaxFileSize     =   256
   End
   Begin VB.Menu FileMenu 
      Caption         =   "&File"
      Begin VB.Menu ImportItem 
         Caption         =   "&Import Data"
      End
      Begin VB.Menu ExportItem 
         Caption         =   "Exp&ort Data"
      End
      Begin VB.Menu EditFileItem 
         Caption         =   "&Edit Data"
      End
      Begin VB.Menu PrintItem 
         Caption         =   "&Print Chart"
      End
      Begin VB.Menu ExitMenu 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu DisplayMenu 
      Caption         =   "&Display"
      Begin VB.Menu InitializeItem 
         Caption         =   "&New Chart"
         Shortcut        =   ^N
      End
      Begin VB.Menu AdmitItem 
         Caption         =   "Show &Admittance Circles"
         Checked         =   -1  'True
         Shortcut        =   ^A
      End
      Begin VB.Menu MarkerItem 
         Caption         =   "Show &Marker"
         Shortcut        =   ^M
      End
      Begin VB.Menu RedrawItem 
         Caption         =   "&Redraw Chart"
         Shortcut        =   ^R
      End
      Begin VB.Menu SeriesItem 
         Caption         =   "&Series Equivalent."
         Shortcut        =   ^S
      End
      Begin VB.Menu ParallelItem 
         Caption         =   "&Parallel Equivalent"
         Shortcut        =   ^P
      End
      Begin VB.Menu ConversItem 
         Caption         =   "Tr. &Line Equivalents"
         Shortcut        =   ^L
      End
      Begin VB.Menu LastItem 
         Caption         =   "Last Import &Freq."
         Shortcut        =   ^F
      End
      Begin VB.Menu togglehist 
         Caption         =   "Show &History"
         Checked         =   -1  'True
         Shortcut        =   ^H
      End
   End
   Begin VB.Menu SweepMenu 
      Caption         =   "&Sweep"
      Begin VB.Menu SweepItem 
         Caption         =   "&Generate Sweep"
         Shortcut        =   ^G
      End
      Begin VB.Menu ClearTrcItem 
         Caption         =   "&Clear Trace"
         Shortcut        =   ^C
      End
      Begin VB.Menu SweepTraceItem 
         Caption         =   "Sweep-&Trace Color"
         Shortcut        =   ^T
      End
   End
   Begin VB.Menu AmplifierMenu 
      Caption         =   "&Amplifier Design"
      Begin VB.Menu AmpItem 
         Caption         =   "&S-Parameter Design/Analysis "
      End
   End
   Begin VB.Menu WindowMenu 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu WinItem 
         Caption         =   "&Cascade"
         Index           =   0
      End
      Begin VB.Menu WinItem 
         Caption         =   "Tile &Horizontal"
         Index           =   1
      End
      Begin VB.Menu WinItem 
         Caption         =   "Tile &Vertical"
         Index           =   2
      End
      Begin VB.Menu WinItem 
         Caption         =   "&Arrange Icons"
         Index           =   3
      End
      Begin VB.Menu WinItem 
         Caption         =   "&Normal"
         Index           =   4
      End
      Begin VB.Menu WinItem 
         Caption         =   "&Minimize"
         Index           =   5
      End
      Begin VB.Menu WinItem 
         Caption         =   "Ma&ximize"
         Index           =   6
      End
      Begin VB.Menu toolbar1item 
         Caption         =   "-"
      End
      Begin VB.Menu ShowToolItem 
         Caption         =   "Show Componen&ts Bar"
         Checked         =   -1  'True
      End
      Begin VB.Menu EditorItem 
         Caption         =   "Show &Editor"
      End
   End
   Begin VB.Menu Help1Item 
      Caption         =   "&Help"
      Begin VB.Menu ContentsItem 
         Caption         =   "&Contents"
      End
      Begin VB.Menu ExamplesItem 
         Caption         =   "&Examples"
      End
      Begin VB.Menu aboutbar 
         Caption         =   "-"
      End
      Begin VB.Menu AboutMenu 
         Caption         =   "&About..."
      End
   End
End
Attribute VB_Name = "form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub AboutMenu_Click()
AboutBox.Show 1  'modal
End Sub

Private Sub AdmitItem_Click()
If (AdmitItem.Checked) Then
AdmitItem.Checked = False
Smith1.ShowAdmittance = False
Else
AdmitItem.Checked = True
Smith1.ShowAdmittance = True
End If
End Sub

Private Sub AmpItem_Click()
   Form6.Show
   MDIForm1.Hide

End Sub

Private Sub AmplifierMenu_Click()
  ' form6.Show
  ' mdiform1.Hide
End Sub

Private Sub ArrangeItem_Click()
MDIForm1.Arrange ARRANGE_ICONS
If (ShowToolItem.Checked = True) Then form4.Show
End Sub

Private Sub CascadeItem_Click()
'mdiform1.Arrange CASCADE
'If (showtoolitem.Checked = True) Then form4.Show
End Sub

Private Sub ChartItem_Click()
form1.Show
End Sub

Private Sub ClearTrcItem_Click()
Smith1.DataReset = 1
Zcalsweep (0)

End Sub

Private Sub CloseItem_Click()
If (ShowToolItem.Checked = True) Then
form4.Hide
ShowToolItem.Checked = False
End If

Unload Me
End Sub

Private Sub Command1_Click()
Frame3D1.Visible = False
End Sub

Private Sub Command2_Click()
form2.Show
End Sub

Private Sub Command3D1_Click()
form2.Show
If (form1.ShowToolItem.Checked = True) Then
form1.ShowToolItem.Checked = False
form4.Hide
Else: form1.ShowToolItem.Checked = True
'ret = SetParent(form4.hWnd, form1.hWnd) 'this makes the toolbox float
'form4.Left = -200 'sets position to top left corner of parent
'form4.Top = 1320
form4.Show 'makes toolbox visible
End If
End Sub

Private Sub Command3D2_Click()
   Form6.Show
   MDIForm1.Hide

End Sub

Private Sub Command3D3_Click()
On Error GoTo SweepSmithErr
 Call Ilcal(1)
Dim final$
Dim ThisFreq As Single
Form3.List1.Clear
points% = Abs((stop1 - start1) / step1) + 1
If points% = 1 Then GoTo Ax1
Form3.Graph1.NumPoints = points%
Form3.Graph2.NumPoints = points%
ThisFreq = start1
For I% = 1 To points%
Form3.Graph1.ThisPoint = I%
Form3.Graph1.GraphData = S21mag(I%) ' insertion loss
Form3.Graph2.GraphData = S21Angle(I%)
final$ = Format(ThisFreq, "###0.000") & "           " & Format(S21mag(I%), "0.000") & "        " & Format(S21Angle(I%), "0.000")
Form3.List1.AddItem final$ ' Add items in the list
ThisFreq = ThisFreq + step1
Next I%

Form3.Graph1.BottomTitle = start1 & "                              To                         " & stop1 & "  [MHz]"
Form3.Graph2.BottomTitle = Form3.Graph1.BottomTitle
Form3.Graph2.Width = Form3.Graph1.Width
Form3.Graph2.Height = Form3.Graph1.Height
Form3.Graph2.Left = Form3.Graph1.Left
Form3.Graph2.Top = Form3.Graph1.Top

Form3.Graph1.DrawMode = 2
Form3.Graph2.DrawMode = 2
GoTo AX2:
Ax1:
Form3.Graph1.NumPoints = 2
Form3.Graph2.NumPoints = 2

Form3.Graph1.GraphData = S21mag(1)
Form3.Graph2.GraphData = S21Angle(1)
Form3.Graph1.GraphData = S21mag(1)
Form3.Graph2.GraphData = S21Angle(1)

Form3.Graph1.BottomTitle = start1 & "                              To                         " & stop1 & "  [MHz]"
Form3.Graph2.BottomTitle = Form3.Graph1.BottomTitle

Form3.Graph2.Width = Form3.Graph1.Width
Form3.Graph2.Height = Form3.Graph1.Height
Form3.Graph2.Left = Form3.Graph1.Left
Form3.Graph2.Top = Form3.Graph1.Top
Form3.Graph2.Top = Form3.Graph1.Top
Form3.Graph1.DrawMode = 2
Form3.Graph2.DrawMode = 2
AX2:
Form3.Graph1.Visible = False
Form3.Graph2.Visible = False
Form3.List1.Visible = False
If (Form3.Option3D1.Value = True) Then Form3.Graph1.Visible = True
If (Form3.Option3D2.Value = True) Then Form3.Graph2.Visible = True
If (Form3.Option3D3.Value = True) Then Form3.List1.Visible = True
If (Form3.Option3D3.Value = True) Then Form3.Panel3D2.Visible = True
Call Ilcal(0)

Form3.Show
MDIForm1.Hide
Exit Sub

SweepSmithErr:
If (Err = 11) Then
    response% = MsgBox("' A Division By Zero' Error Was Trapped", 49, "Error")
    If response = 1 Then Resume Else Exit Sub
End If



End Sub

Private Sub Command3D4_Click()
Smith1.ResetAll = True
Panel3D1.Caption = "Q: 0 "
Panel3D2.Caption = "VSWR: 1"
End Sub

Private Sub Command3D5_Click()
Smith1.Redraw = True
End Sub

Private Sub Command3D6_Click()
Dim NL, Filedata0, Filedata1, Filedata2
Dim n As Long
Dim Z0 As Single
NL = Chr(10)
On Error Resume Next
Smith1.Redraw = True
CMDialog1.CancelError = -1
CMDialog1.DialogTitle = "Import"
CMDialog1.Filter = "Gamma Data Files | *.Gam"
CMDialog1.Flags = OFN_READONLY
CMDialog1.Action = 1
If Err = 0 Then
    Open CMDialog1.Filename For Input As #1
    Input #1, Z0
    Smith1.Z0 = Z0
    Do While Not EOF(1)
         Input #1, Filedata0, Filedata1, Filedata2
         Smith1.DrawM = Filedata1
         Smith1.DrawQ = Filedata2
    Loop
    'Print n

    Close #1
End If
End Sub

Private Sub ContentsItem_Click()
R = WinHelp(form1.hwnd, HelpFilePath, HELP_CONTEXT, CLng(1))
End Sub

Private Sub ConversItem_Click()
On Error GoTo ConverError

MZ! = Sqr(ZR(13) ^ 2 + ZI(13) ^ 2)           ' Zin
AZ! = Atn(ZI(13) / ZR(13)) * 180 / pi
   'If (MZ! = 50) Then AZ! = 0

                                            ' Yin
dx# = ZR(13) ^ 2 + ZI(13) ^ 2
YRX! = ZR(13) * 1000 / dx#
YIX! = -ZI(13) * 1000 / dx#
MY! = Sqr(YRX! ^ 2 + YIX! ^ 2)
If (YRX! = 0) Then YRX! = 0.000000000001
AY! = Atn(YIX! / YRX!) * 180 / pi
   SM! = ZtoGammaM(CDbl(ZR(13)), CDbl(ZI(13)))                                     ' S11/Gamma
   SQ! = ZtoGammaA(CDbl(ZR(13)), CDbl(ZI(13)))
   If (SM! = 0) Then SQ! = 0
   If (SM! = 0) Then SM! = 1E-36
   SR# = SM! * Cos(SQ! * pi / 180)
   SI# = SM! * Sin(SQ! * pi / 180)

Swr! = (1 + SM!) / (1 - SM!)                    ' Swr
RL = Abs(20 * (Log(SM!) / Log(10)))                ' Rl
If (RL >= 720) Then st$ = ">" Else st$ = ""

Call Ilcal(0)                                  ' S21

msg$ = " ZIN : " & Format(ZR(13), "0.000") & " + " & Format(ZI(13), "0.000") & "j" & "   " & Format(MZ!, "0.000") & " < " & Format(AZ!, "0.000") & CRLF & "YIN(mS) : " & Format(YRX!, "0.000") & " + " & Format(YIX!, "0.000") & "j" & "   " & Format(MY!, "0.000") & " < " & Format(AY!, "0.000") & CRLF & "GAM : " & Format(SR#, "0.000") & " + " & Format(SI#, "0.000") & "j" & "   " & Format(SM!, "0.000") & " < " & Format(SQ!, "0.000") & CRLF & "S21 : " & Format(S21IM#, "0.000") & " dB" & " < " & Format(S21IA#, "0.000") & CRLF & "VSWR : " & Format(Swr!, "0.000") & CRLF & "Return Loss : " & st$ & Format(RL, "0.000") & " dB" & CRLF & "Insertion Loss : " & Format(-S21IM#, "0.000") & " dB"
msg1$ = "Reflection Loss : " & Format(-10 * Log(1 - SM! ^ 2) / Log(10), "0.000") & " dB" & CRLF & "Power Reflection Coefficient : " & Format((SM! ^ 2), "0.000") & CRLF & "Transmission Loss Coefficient : " & Format(((1 + (Swr!) ^ 2) / (2 * Swr!)), "0.000") & CRLF & "Maximum Standing Wave: " & Format(Sqr(Swr!), "0.000") & CRLF & "Minimum Standing Wave: " & Format(1 / Sqr(Swr!), "0.000")
'msg2$ = " Note: S21 vaild only if Z0 = Load "
MsgBox msg$ & CRLF & msg1$, 0, "Transmission Line Calculations"

Exit Sub
ConverError:
If (Err = 11) Then
response% = MsgBox("'A Division By Zero' Error Was Trapped", 49, "Error")
If response = 1 Then Resume Else Exit Sub
End If



End Sub

Private Sub EditFileItem_Click()

SCHl = 0
Form5.Show
End Sub

Private Sub EditorItem_Click()
Form5.Show
End Sub

Private Sub ExamplesItem_Click()
R = WinHelp(form1.hwnd, HelpFilePath, HELP_CONTEXT, CLng(2))
End Sub

Private Sub ExitMenu_Click()
If (EdopenFlag = True) Then Unload Form5
If (PltopenFlag = True) Then Unload Form3
If (SPAROPENFLAG = True) Then Unload Form9
If (NOISEOPENFLAG = True) Then Unload Form10
If (AMPOPENFLAG = True) Then Unload Form6
If (SweepopenFlag = True) Then Unload Form8
Unload MDIForm1


End Sub

Private Sub ExportItem_Click()
On Error Resume Next
CMDialog2.CancelError = -1
CMDialog2.DialogTitle = "Export Data"
CMDialog2.Filter = "Gamma Data Files| *.Gam"
CMDialog2.Flags = &H4    ' hide read only box
CMDialog2.Filename = "*.Gam"
CMDialog2.Action = 2

points% = Abs(CInt((stop1 - start1) / step1))
f! = start1

If Err = 0 Then
    If (Dir(CMDialog2.Filename) <> "") Then 'file already exist
        response = MsgBox("Overwrite Existing File ?", 4 + 32 + 256, "QuickSmith")
        If response = 7 Then Exit Sub
    End If
   
   Open CMDialog2.Filename For Output As #1
   Print #1, "QuickSmith GAM"
   Print #1, Z0
   
   For I% = 0 To points% Step 1
   Print #1, Format(f!, "######0.000") & "    " & Format(GammaMag(I%), "##0.000") & "      " & Format(GammaAngle(I%), "#00.000")
   f! = f! + step1
   Next I%
   
   Close #1
End If

End Sub

Private Sub Form_Activate()
If (ShowToolItem.Checked = True) Then form4.Show

For I% = 2 To 10 Step 2
form4.Panel3D3(I%) = form2.Panel3D3(I%)
form4.Label1(I%).Caption = form2.Image11(I%).Tag & I%
Next I%
For I% = 3 To 9 Step 2
form4.Panel3D6(I%) = form2.Panel3D6(I%)
form4.Label1(I%).Caption = form2.Image12(I%).Tag & I%
Next I%
form4.Panel3D10 = form2.Panel3D10

End Sub

Private Sub Form_Deactivate()
 form4.Hide
End Sub

Private Sub Form_Load()

SCHl = 0
qstep = 1
vswrstep = 1
Command3 = 0
Command2 = 0
Smith1.Z0 = 50
Smith1.Font.Size = 6.6
ShowToolItem.Checked = True
Ret = SetParent(form4.hwnd, form1.hwnd) 'this makes the toolbox float
form4.Show 'makes toolbox visible
form4.Move 0, 0.95 * (Frame3D1.Height + Panel3D4.Height)
WindowVertical = True ' changed to true ver4.2
WindowHorizontal = False
form4.Hide 'hides toolbox  added v4.2
ShowToolItem.Checked = False   ' added V4.2
 ' added v4.2 begin
    If (WindowVertical) Then
    form1.HScroll1.Visible = True
    form2.HScroll1.Visible = True
    form1.HScroll1.Move 0, form1.ScaleHeight - form1.HScroll1.Height, form1.ScaleWidth
    If (form2.VScroll1.Visible) Then
        form2.HScroll1.Move 0, form2.ScaleHeight - form2.HScroll1.Height, form2.ScaleWidth - form2.VScroll1.Width
    Else
        form2.HScroll1.Move 0, form2.ScaleHeight - form2.HScroll1.Height, form2.ScaleWidth
    End If
    form1.Frame3D1.Left = -HScroll1.Value
    form1.Frame3D1.Width = form1.Width + HScroll1.Value
    form2.Picture2.Left = -form2.HScroll1.Value
    Else
    form1.HScroll1.Visible = False
    form2.HScroll1.Visible = False
    form1.Frame3D1.Left = 0
    End If
 ' added v4.2 end
    

EdopenFlag = False ' Flag to notify exit event if editor is closed
AMPOPENFLAG = False
NOISEOPENFLAF = False
SPAROPENFLAG = False
PltopenFlag = False
Frame3D1.Left = 0
Frame3D1.Width = Width
form2.Show
Smith1.ShowAdmittance = True

 End Sub

Private Sub Form_Resize()
On Error Resume Next
Dim OldScaleMode As Integer
OldScaleMode = form1.ScaleMode
form1.ScaleMode = 1 ' Twips
If form1.WindowState = 2 Then WindowVertical = False ' maximized
If (WindowVertical = True) Then
    Smith1.Top = -35 + Frame3D1.Height + Panel3D4.Height * 0.8
    Smith1.Left = 0
Else
    Smith1.Top = 10 + Frame3D1.Height
    Smith1.Left = form1.Width / 3.5

End If

Frame3D1.Width = form1.Width
If form1.WindowState = 2 Or ShowToolItem.Checked Then
    ShowToolItem.Checked = True
    form4.Show 'makes toolbox visible
    form1.HScroll1.Visible = False
    WindowVertical = False
End If



   ' Enable or Disable HScroll1 depending on new size of Form
  HScroll1.Value = 0
  HDiff = ScaleWidth - (Command3D1.Left + Command3D1.Width)
  HScroll1.Visible = HDiff <= 0
  Smith1.Height = Abs(form1.ScaleHeight - Smith1.Top) ' negative value crashes
  If HScroll1.Visible Then
      HScroll1.Move 0, ScaleHeight - HScroll1.Height, ScaleWidth
      HScroll1.max = Abs(HDiff)
      HScroll1.LargeChange = HScroll1.max \ 10
      Smith1.Height = Abs(form1.ScaleHeight - Smith1.Top - HScroll1.Height) ' negative value crashes
  End If

Smith1.Width = Abs(form1.ScaleWidth - Smith1.Left)
If Smith1.Width > Smith1.Height Then Smith1.Width = Smith1.Height Else Smith1.Height = Smith1.Width

form1.ScaleMode = OldScaleMode
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload form4

End Sub

Private Sub HScroll1_Change()
Frame3D1.Left = -HScroll1.Value
 Frame3D1.Width = form1.Width + HScroll1.Value
'Static scrollvalue As Integer
'ScrollMe form1, CInt(HScroll1.Value) - scrollvalue, 0

'scrollvalue = HScroll1.Value


End Sub

Private Sub ImportItem_Click()
Dim NL, dat
Dim Filedata0 As Double
Dim Filedata1 As Double
Dim Filedata2 As Double
Dim Z0 As Single
Dim n As Long
NL = Chr(10)
Dim gamh As Integer ' used as glag to detect HZ file
gamh = 0
On Error Resume Next
Smith1.Redraw = True
CMDialog1.CancelError = -1
CMDialog1.DialogTitle = "Import"
CMDialog1.Filter = "Gamma Data Files | *.Gam"
CMDialog1.Flags = OFN_READONLY
CMDialog2.Filename = "*.Gam"
CMDialog1.Action = 1
If Err = 0 Then
    Open CMDialog1.Filename For Input As #1
    Input #1, dat
    If (dat <> "QuickSmith GAM") Then
        If (dat = "QuickSmith GAMH") Then
            gmah = 1
            GoTo label101
        End If
    Beep
    MsgBox "Bad Gamma File !" & CRLF & " Check QuickSmith Data File Format ", 48, "QuickSmith"
    Close #1
    Exit Sub
label101:
    End If
   Input #1, Z0
    If (Z0 = 0) Then Z0 = 0.0000000001 ' avoid divide by zero errors
    Smith1.Z0 = Z0
    Input #1, Filedata0, Filedata1, Filedata2
              Smith1.DrawM = Filedata1
              Smith1.DrawQ = Filedata2
              FreqStart! = Filedata0
              If gmah = 1 Then FreqStart! = FreqStart! / 1000000#
              n = 1
              'MsgBox Filedata0 & " " & Filedata1 & "  " & Filedata2
    Do While Not EOF(1)
         Input #1, Filedata0, Filedata1, Filedata2
         If Filedata0 = 0 Then Exit Do
         Smith1.DrawM = Filedata1
         Smith1.DrawQ = Filedata2
         n = n + 1
         FreqStop! = Filedata0
         
         If gmah = 1 Then FreqStop! = FreqStop! / 1000000#
        'MsgBox Format(FreqStop!, "0.000")
        'MsgBox Filedata0 & " " & Filedata1 & "  " & Filedata2
    Loop
     FreqPoints% = n
  '   MsgBox n & " " & FreqStart! & "  " & FreqStop!
    Close #1
'If (FreqPoints% > 0) Then FreqStep! = Abs(freqstop! - Freqstart! + 1) / FreqPoints% Else FreqStep! = 0
'msg$ = "Start Frequency : " & Format(Freqstart!, "0.000") & "  MHz " & CRLF & CRLF & " Stop Frequency :  " & Format(freqstop!, "0.000") & "  MHz " & CRLF & CRLF & " Frequency Step :  " & Format(FreqStep!, "0.000") & "  MHz " & CRLF & CRLF & " Frequency Points :  " & Format(FreqPoints%, "#####000")

'MsgBox msg$, 0, "Import-Data Frequency Range"


End If


End Sub

Private Sub InitializeItem_Click()
Smith1.ResetAll = 1
AdmitItem.Checked = False
Panel3D2.Caption = "VSWR: " & Format(1, ".0#")
MarkerItem.Checked = False
Panel3D1.Caption = "Q: " & Format(0, "0.0#")
Zcalsweep (0)
End Sub

Private Sub ItemCircles_Click()
If (ShowToolItem.Checked = True) Then
ShowToolItem.Checked = False
form4.Hide
Else: ShowToolItem.Checked = True
'ret = SetParent(form4.hWnd, form1.hWnd) 'this makes the toolbox float
'form4.Left = -200 'sets position to top left corner of parent
'form4.Top = 1320
form4.Show 'makes toolbox visible
End If

End Sub

Private Sub ItemEnter_Click()
MDIForm1.Hide
Form6.Show
End Sub

Private Sub LastItem_Click()
 
'If (FreqPoints% > 0) Then FreqStep! = (Abs(freqstop! - Freqstart!) / FreqPoints%) + 1 Else FreqStep! = 0
msg$ = "Start Frequency : " & Format(FreqStart!, "0.000") & "  MHz " & CRLF & CRLF & " Stop Frequency :  " & Format(FreqStop!, "0.000") & "  MHz " & CRLF & CRLF & " Frequency Points :  " & Format(FreqPoints%, "#####000")

MsgBox msg$, 0, "Last Import-Data Frequency Range"

End Sub

Private Sub MarkerItem_Click()
Dim R, I  As Double
If (MarkerItem.Checked) Then
MarkerItem.Checked = False
Smith1.ShowMarker = False
Else
Title$ = "Smith Chart"
prompt$ = "Enter Series Resistance Value"
Default$ = ""
RX$ = InputBox$(prompt$, Title$, Default$)
If (RX$ = "") Then GoTo Enda
If (IsNumeric(RX$)) Then R = CDbl(RX$)

prompt$ = "Enter Series Reactance Value"
Q1$ = InputBox$(prompt$, Title$, Default$)
If (Q1$ = "") Then GoTo Enda
Smith1.ShowMarker = True
MarkerItem.Checked = True
If (IsNumeric(Q1$)) Then I = CDbl(Q1$)
Smith1.MarkerM = ZtoGammaM(CDbl(R), CDbl(I))
Smith1.MarkerQ = ZtoGammaA(CDbl(R), CDbl(I))
Enda:
End If

End Sub

Private Sub Option3D1_Click(Value As Integer)
Call Yin
End Sub

Private Sub Option3D2_Click(Value As Integer)
form1.Label1.Caption = Format(ZR(13), "###0.0000") ' reflection coeff.
form1.Label2.Caption = Format(ZI(13), "###0.0000")
Label5.Caption = "Z-In"
Label3.Caption = "Re"
Label4.Caption = "Im"
Label6.Caption = "Ohms"
End Sub

Private Sub Panel3D1_DblClick()
Title$ = "Smith Chart"
prompt$ = "Enter Step Size for Q Circles"
Default$ = CStr(qstep)
stepsize$ = InputBox$(prompt$, Title$, Default$)
If (IsNumeric(stepsize$)) Then qstep = CDbl(stepsize$)

End Sub

Private Sub Panel3D2_DblClick()
Title$ = "Smith Chart"
prompt$ = "Enter Step Size for VSWR Circles"
Default$ = CStr(vswrstep)
stepsize$ = InputBox$(prompt$, Title$, Default$)
If (IsNumeric(stepsize$)) Then vswrstep = CDbl(stepsize$)
'stepsize$ = InputBox$("Enter Step Size ", "QuickSmith", CStr(stepS(Index)))
'If (IsNumeric(stepsize$)) Then stepS(Index) = Abs(CDbl(stepsize$))


End Sub

Private Sub Panel3D4_DblClick()
stepsize$ = InputBox$("Enter Step Size ", "Smith Chart", "")
If (IsNumeric(stepsize$)) Then stepS(index) = CDbl(stepsize$)

End Sub

Private Sub Panel3D5_Click()
Title$ = "Smith Chart"
prompt$ = "Enter Chart Z0"
Default$ = ""
stepsize$ = InputBox$(prompt$, Title$, Default$)
If (IsNumeric(stepsize$)) Then z0step = CDbl(stepsize$)

End Sub

Private Sub ParallelItem_Click()
ZinR! = Format(ZR(13), ".0000")    'Panel3d8.Caption
ZinI! = Format(ZI(13), ".0000")   'label2.caption

On Error GoTo ParallelItemError

If (ZinR! = 0) Then ZinR! = 0.00000000000001
Qu! = ZinI! / ZinR!
ZinR! = (Qu! ^ 2 + 1) * ZinR!
If (Qu! <> 0) Then ZinI! = ZinR! / Qu!

If (ZinI! < 0) Then                ' capacitor
LC! = 1 / (2 * pi * freq * ZinI! * 0.000001)
EL$ = "C"
PN$ = " pf"
ElseIf (ZinI! > 0) Then            ' inductor
LC! = ZinI! / (2 * pi * freq * 0.001)
EL$ = "L"
PN$ = " nh"
Else                             ' neither cap nor ind
LC! = 0
EL$ = " "
End If

If (LC! = 0) Then
msg$ = "R = " & Format(ZinR!, ".000") & " ohms"
Else: msg$ = "R = " & Format(ZinR!, ".000") & " ohms" & CRLF & CRLF & "in Parallel with " & CRLF & CRLF & EL$ & " = " & Format(Abs(LC!), "0.000") & PN$
End If
MsgBox msg$, 0, "Parallel Equivalent"

Exit Sub

ParallelItemError:
If (Err = 11) Then
    response% = MsgBox("' A Division By Zero' Error Was Trapped", 49, "Error")
    If response = 1 Then Resume Else Exit Sub
End If



End Sub

Private Sub PrintItem_Click()

Dim msg$, prompt$, cap$
Dim hscroll_flag

    
    On Error GoTo PrintSetupError
    CMDialog1.CancelError = -1
    CMDialog1.Flags = &H40& 'PD_PRINTSETUP     ' Just enable print setup
    CMDialog1.Action = 5 ' DLG_PRINT        ' Show printer setup dialog box
    
If Err = 0 Then
    If (ShowToolItem.Checked = True) Then form4.Hide
  Frame3D1.Visible = False
    Command3D1.Visible = False
    Command3D2.Visible = False
    If (form1.HScroll1.Visible) Then
    hscroll_flag = True
    form1.HScroll1.Visible = False
    End If
    'modified code here to call new module print1.bas
    'instead of print.bas, to facilate high resolution printing

    Call PrintClient256(form1)
    
     If (ShowToolItem.Checked = True) Then form4.Show
     Frame3D1.Visible = True
     Command3D1.Visible = True
     Command3D2.Visible = True
     If (hscroll_flag) Then form1.HScroll1.Visible = True
End If

PrintSetupExit:
    Exit Sub

PrintSetupError:
    If Err <> CDERR_CANCEL Then         ' If user didn't select CANCEL
    End If
    Resume PrintSetupExit

  
End Sub

Private Sub RedrawItem_Click()
Smith1.Redraw = True
End Sub

Private Sub ReseDtItem_Click()
Smith1.DataReset = True
Call Zcalsweep(0)
End Sub

Private Sub ResetItem_Click()
Smith1.ResetAll = True
Panel3D1.Caption = "Q: 0 "
Panel3D2.Caption = "VSWR: 1"
End Sub

Private Sub SchMenu_Click()
'If form2.WindowState = 3 Then form2.WindowState = 0
form2.Show
End Sub

Private Sub SeriesItem_Click()
ZinR! = Format(ZR(13), ".0000")    'label1.Caption
ZinI! = Format(ZI(13), ".0000")   'label2.caption

On Error GoTo SeriesItemError

If (ZinI! < 0) Then      ' capacitor
LC! = 1 / (2 * pi * freq * ZinI! * 0.000001)
EL$ = "C"
PN$ = " pf"
ElseIf (ZinI! > 0) Then      ' inductor
LC! = ZinI! / (2 * pi * freq * 0.001)
EL$ = "L"
PN$ = " nh"
Else                       ' neither cap nor ind
LC! = 0
EL$ = " "
End If

If (LC! = 0) Then
msg$ = "R = " & Format(ZinR!, ".000") & " ohms"
Else: msg$ = "R = " & Format(ZinR!, ".000") & " ohms" & CRLF & CRLF & "in Series with " & CRLF & CRLF & EL$ & " = " & Format(Abs(LC!), "0.000") & PN$
End If
MsgBox msg$, 0, "Series Equivalent"

Exit Sub
SeriesItemError:
If (Err = 11) Then
    response% = MsgBox("' A Division By Zero' Error Was Trapped", 49, "Error")
    If response = 1 Then Resume Else Exit Sub
End If

End Sub

Private Sub ShowToolItem_Click()
If (ShowToolItem.Checked = True) Then
ShowToolItem.Checked = False
form4.Hide
Else: ShowToolItem.Checked = True
'ret = SetParent(form4.hWnd, form1.hWnd) 'this makes the toolbox float
'form4.Left = -200 'sets position to top left corner of parent
'form4.Top = 1320
form4.Show 'makes toolbox visible
End If

End Sub

Private Sub Smith1_ClickIn(m As Single, Q As Single)
Dim str1 As String
On Error GoTo ClickinError
If (m = 0) Then m = 1E-36
If (m = 1) Then m = 0.9999999999
Swr! = (1 + m) / (1 - m)
S11 = -20 * (Log(m) / Log(10))
If (S11 >= 720) Then str1 = " > " Else str1 = ""

MsgBox ("Gamma [Mag]: " & Format(m, "#0.000") & Chr(10) & "Gamma [Ang]: " & Format(Q, "###.00") & Chr(176) & Chr(10) & "VSWR: " & Format(Swr!, "###.00") & Chr(10)) & "Return Loss: " & str1 & Format(S11, "##0.00") & " dB", 0, "QuickSmith"

Exit Sub
ClickinError:
If (Err = 11) Then
response% = MsgBox("'A Division By Zero' Error Was Trapped", 49, "Error")
If response = 1 Then Resume Else Exit Sub
End If

End Sub

Private Sub Smith1_ClickOut()
Beep
End Sub

Private Sub Smith1_MouseMove(m As Single, Q As Single)
If (Smith1.Z0 = 0) Then
Title$ = "Smith Chart"
prompt$ = "Enter Z0 Value"
Default$ = ""
ZX$ = InputBox$(prompt$, Title$, Default$)
If (ZX$ = "") Then GoTo EndE
If (IsNumeric(ZX$)) Then Smith1.Z0 = CDbl(ZX$)
EndE:
End If
Z1r! = GToZR(m, Q)
Z1i! = GToZI(m, Q)

If (m > 1.5) Then
Panel3D3.Caption = "Z: "
form1.MousePointer = 0
Else
Panel3D3.Caption = "Z:   " & Format(Z1r!, "######0.##") & Format(Z1i!, "+#####0.##;-###0.##") & "j"
 form1.MousePointer = 10
End If

End Sub

Private Sub Spin1_SpinDown()
Dim qtext, start
Dim Q As Single
qtext = Panel3D1.Caption
start = InStr(qtext, "Q: ")
Mid(qtext, start, 4) = "   "
Q = qtext
Q = Q - qstep
If (Q < 0) Then Q = 0
Smith1.QCircle = Q
Panel3D1.Caption = "Q: " & Format(Q, "0.0#")


End Sub

Private Sub Spin1_SpinUp()
Dim qtext, start
Dim Q As Single
qtext = Panel3D1.Caption
start = InStr(qtext, "Q: ")
Mid(qtext, start, 4) = "   "
Q = qtext
Q = Q + qstep
If (Q < 0) Then Q = 0
Smith1.QCircle = Q
Panel3D1.Caption = "Q: " & Format(Q, "0.0#")

End Sub

Private Sub Spin2_SpinDown()
Dim vswrtext, start
Dim vswr As Single
vswrtext = Panel3D2.Caption
start = InStr(vswrtext, "VSWR: ")
Mid(vswrtext, start, 6) = "      "
vswr = vswrtext
vswr = vswr - vswrstep
If (vswr < 1) Then vswr = 1
Smith1.VSWRCircle = vswr
Panel3D2.Caption = "VSWR: " & Format(vswr, ".0#")

End Sub

Private Sub Spin2_SpinUp()
Dim vswrtext, start
Dim vswr As Single
vswrtext = Panel3D2.Caption
start = InStr(vswrtext, "VSWR: ")
Mid(vswrtext, start, 6) = "      "
vswr = vswrtext
vswr = vswr + vswrstep
If (vswr < 1) Then vswr = 1
Smith1.VSWRCircle = vswr
Panel3D2.Caption = "VSWR: " & Format(vswr, ".0#")
End Sub

Private Sub Spin3_SpinDown(index As Integer)
 '       Select Case form2.Image11(Index).Tag
  '
    '    Case "R"
   '         Z1r(Index) = Z1r(Index) - stepS(Index)
     '       Z1i(Index) = 0
      '      form2.Panel3D3(Index).Caption = Format(Z1r(Index), "0.000")
       '     If (Z1r(Index) > 99999 Or Z1r(Index) < .99) Then form2.Panel3D3(Index).Caption = Format$(Z1r(Index), "##.#E+##")
     
       ' Case "C"
        '   Z1r(Index) = 0
           
         '  ZC(Index) = ZC(Index) - stepS(Index)
           'If (ZC(index) = 0) Then ZC(index) = 1E-45
          ' If (ZC(Index) > 99999.99 Or ZC(Index) < -99999.99) Then form2.Panel3D3(Index).Caption = Format$(C!, "##.#E+##")
          'form2.Panel3D3(Index).Caption = Format(ZC(Index), "0.000")
          'Case "I"
          ' Z1r(Index) = 0
          ' ZL(Index) = ZL(Index) - stepS(Index)
           'If (ZL(index) = 0) Then ZL(index) = 1E-45
          ' If (ZL(Index) > 99999.99 Or ZL(Index) < -99999.99) Then form2.Panel3D3(Index).Caption = Format$(C!, "##.#E+##")
          'form2.Panel3D3(Index).Caption = Format(ZL(Index), "0.000")
        'Case "W"
        
        'Case "X"
         '   Z1i(Index) = Z1i(Index) - stepS(Index)
          '  Z1r(Index) = 0
           'form2.Panel3D3(Index).Caption = Format(Z1i(Index), "0.000")
           ' If (Z1i(Index) > 99999.99 Or Z1i(Index) < -99999.99) Then form2.Panel3D3(Index).Caption = Format$(Z1i(Index), "##.#E+##")
        
       ' End Select
' Call zcal




End Sub

Private Sub sweepItem_click()

If (Element = 1) Then GoTo sweep
    If (form2.Option3D2.Value) Then
         If (Element Mod 2 = 0) Then
         Select Case form2.Image11(Element).Tag
         Case "S", "T", "O", "SLC", "PLC"
         Zcalsweep (Element + 100)
         End Select
         Else
           Select Case form2.Image12(Element).Tag
           Case "S", "T", "O", "SLC", "PLC"
           Zcalsweep (Element + 100)
           End Select
         End If
    Else: Zcalsweep (Element)
    End If
    Zcalsweep (0) ' spot.. bug fixed to redisplay
    'spot value after sweep
Exit Sub
sweep:
 Zcalsweep (1) ' sweep freq
 Zcalsweep (0) ' spot.. bug fixed to redisplay
    'spot value after sweep

End Sub

Private Sub SweepTraceItem_Click()
On Error Resume Next
CMDialog1.CancelError = -1
CMDialog1.Color = Smith1.ForeColor
CMDialog1.Flags = &H1 ' show initial value
CMDialog1.Action = 3
If Err = 0 Then Smith1.ForeColor = CMDialog1.Color
End Sub

Private Sub Text1_DblClick()
Title$ = "Smith Chart"
prompt$ = "Enter Step Size for VSWR Circles"
Default$ = ""
stepsize$ = InputBox$(prompt$, Title$, Default$)
If (IsNumeric(stepsize$)) Then qstep = CDbl(stepsize$)


End Sub

Private Sub TileHItem_Click()
'mdiform1.Arrange TILE_HORIZONTAL
'If (showtoolitem.Checked = True) Then
'form4.Hide
'showtoolitem.Checked = False
'End If
End Sub

Private Sub TileVItem_Click()
'mdiform1.Arrange TILE_VERTICAL
'If (showtoolitem.Checked = True) Then
'form4.Hide
'showtoolitem.Checked = False
'End If

End Sub

Private Sub UsingItem_Click()

End Sub

Private Sub togglehist_Click()
If (togglehist.Checked) Then
togglehist.Checked = False
form1.Smith1.ShowHistory = False
Form6.Smith2.ShowHistory = False
Form6.Smith1.ShowHistory = False
Else
togglehist.Checked = True
form1.Smith1.ShowHistory = True
Form6.Smith2.ShowHistory = True
Form6.Smith1.ShowHistory = True
End If
End Sub

Private Sub WinItem_Click(index As Integer)
  
  On Error GoTo form1WindowError
  Select Case index
      Case 0:
            ShowToolItem.Checked = True
            form4.Show 'makes toolbox visible
            WindowVertical = False
            WindowHorizontal = False
            mdiArrange WM_MDICASCADE

      Case 1:
            form4.Hide 'makes toolbox visible
            ShowToolItem.Checked = False
            WindowVertical = False
            WindowHorizontal = True
            mdiArrange MDITILE_HORIZONTAL
             'End If
      Case 2:
             WindowVertical = True
             WindowHorizontal = False
             form4.Hide 'makes toolbox visible
             ShowToolItem.Checked = False
             mdiArrange MDITILE_VERTICAL
      
      Case 3: mdiArrange WM_MDIICONARRANGE
      Case 4:
             WindowVertical = False
             WindowHorizontal = False
            If (form1.WindowState Or form2.WindowState) Then
            form1.WindowState = 0
            form2.WindowState = 0
            mdiArrange WM_MDICASCADE
            End If
            form1.Move 0, 0, 0.84 * MDIForm1.Width, 0.85 * MDIForm1.Height
            form2.Move 400, 0, 0.84 * MDIForm1.Width, 0.75 * MDIForm1.Height
            ShowToolItem.Checked = True
            form4.Show 'makes toolbox visible
           
                
      Case 5:
              WindowVertical = False
              WindowHorizontal = False
              form1.WindowState = 1
      
      Case 6:
              WindowVertical = False
              WindowHorizontal = False
              form1.WindowState = 2
    End Select
    
    form2.Picture1.Move 0, 0, form2.ScaleWidth, form2.ScaleHeight
    form2.Picture2.Move 0, 0
    
    If (WindowVertical) Then
    form1.HScroll1.Visible = True
    form2.HScroll1.Visible = True
    form1.HScroll1.Move 0, form1.ScaleHeight - form1.HScroll1.Height, form1.ScaleWidth
    If (form2.VScroll1.Visible) Then
        form2.HScroll1.Move 0, form2.ScaleHeight - form2.HScroll1.Height, form2.ScaleWidth - form2.VScroll1.Width
    Else
        form2.HScroll1.Move 0, form2.ScaleHeight - form2.HScroll1.Height, form2.ScaleWidth
    End If
    
    'form2.HScroll1.Move 0, form2.ScaleHeight - form2.HScroll1.Height, form2.ScaleWidth
    
    form1.Frame3D1.Left = -HScroll1.Value
    form1.Frame3D1.Width = form1.Width + HScroll1.Value
    form2.Picture2.Left = -form2.HScroll1.Value
    Else
    form1.HScroll1.Visible = False
    form2.HScroll1.Visible = False
    form1.Frame3D1.Left = 0
    End If

    If (WindowHorizontal) Then
    form2.VScroll1.Visible = True
    If (form2.HScroll1.Visible) Then
    form2.VScroll1.Move form2.ScaleWidth - form2.VScroll1.Width, 0, form2.VScroll1.Width, form2.ScaleHeight - form2.HScroll1.Height
    Else
    form2.VScroll1.Move form2.ScaleWidth - form2.VScroll1.Width, 0, form2.VScroll1.Width, form2.ScaleHeight
    End If
    form2.Picture2.Left = -form2.VScroll1.Value
    Else
    form2.VScroll1.Visible = False
    End If

Exit Sub
form1WindowError:
If (Err) Then
response% = MsgBox("QuickSmith Detected An Error", 49, "Error")
If response = 1 Then Resume Else Exit Sub
End If



End Sub

