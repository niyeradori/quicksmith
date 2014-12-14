VERSION 5.00
Object = "{0842D103-1E19-101B-9AAF-1A1626551E7C}#1.0#0"; "GRAPH32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form3 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   Caption         =   "QuickSmith-S21"
   ClientHeight    =   8400
   ClientLeft      =   1275
   ClientTop       =   2340
   ClientWidth     =   11625
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
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form3"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8400
   ScaleWidth      =   11625
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "Edit"
      Height          =   372
      Left            =   840
      TabIndex        =   13
      Top             =   120
      Width           =   732
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H00FF0000&
      Height          =   312
      Left            =   1680
      TabIndex        =   11
      Top             =   660
      Width           =   6012
   End
   Begin Threed.SSPanel Panel3D2 
      Height          =   375
      Left            =   2760
      TabIndex        =   10
      Top             =   1200
      Visible         =   0   'False
      Width           =   3495
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   15
      Caption         =   "Freq[MHz]       Mag.[dB]       Ang.[D]"
      BevelWidth      =   2
      BorderWidth     =   1
      BevelOuter      =   0
      Outline         =   -1  'True
      Alignment       =   1
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   4320
      Left            =   2760
      TabIndex        =   9
      Top             =   1572
      Visible         =   0   'False
      Width           =   3492
   End
   Begin Threed.SSOption Option3D3 
      Height          =   372
      Left            =   5520
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   120
      Width           =   852
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   78
      Caption         =   "LIST"
      ForeColor       =   4210752
   End
   Begin MSComDlg.CommonDialog CMDialog3 
      Left            =   0
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FontSize        =   0
      MaxFileSize     =   256
   End
   Begin GraphLib.Graph Graph2 
      Height          =   1692
      Left            =   0
      TabIndex        =   7
      Top             =   5520
      Visible         =   0   'False
      Width           =   1932
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   2646
      _StockProps     =   96
      DrawMode        =   3
      GraphTitle      =   "S21-  ANGLE"
      GraphType       =   6
      GridStyle       =   1
      LabelEvery      =   10
      LeftTitle       =   "Degrees"
      NumPoints       =   10
      PrintStyle      =   2
      RandomData      =   0
      ThickLines      =   0
      TickEvery       =   10
      ColorData       =   0
      ExtraData       =   0
      ExtraData[]     =   0
      FontFamily      =   1
      FontSize        =   0
      FontStyle       =   1
      GraphData       =   0
      GraphData[]     =   0
      LabelText       =   0
      LegendText      =   0
      PatternData     =   0
      SymbolData      =   0
      XPosData        =   0
      XPosData[]      =   0
   End
   Begin Threed.SSOption Option3D2 
      Height          =   375
      Left            =   4440
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   120
      Width           =   1095
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   78
      Caption         =   "ANGLE"
   End
   Begin Threed.SSOption Option3D1 
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   120
      Width           =   855
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   78
      Caption         =   "MAG."
      Value           =   -1  'True
   End
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      Caption         =   "Back"
      Height          =   372
      Left            =   2268
      TabIndex        =   2
      Top             =   120
      Width           =   732
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Caption         =   "Print"
      Height          =   372
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   732
   End
   Begin VB.CommandButton Command4 
      Appearance      =   0  'Flat
      Caption         =   "Save"
      Height          =   372
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   732
   End
   Begin Threed.SSPanel Panel3D1 
      Height          =   612
      Left            =   -360
      TabIndex        =   6
      Top             =   0
      Width           =   12552
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   15
      BackColor       =   -2147483633
      BevelWidth      =   2
      BorderWidth     =   1
   End
   Begin GraphLib.Graph Graph1 
      Height          =   5415
      Left            =   600
      TabIndex        =   0
      Top             =   1260
      Width           =   8895
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   2646
      _StockProps     =   96
      DrawMode        =   3
      GraphTitle      =   "S21 - MAGNITUDE"
      GraphType       =   6
      GridStyle       =   1
      LabelEvery      =   10
      LeftTitle       =   "dB"
      NumPoints       =   2
      PrintStyle      =   2
      RandomData      =   0
      ThickLines      =   0
      TickEvery       =   10
      ColorData       =   0
      ExtraData       =   0
      ExtraData[]     =   0
      FontFamily      =   1
      FontSize        =   0
      FontStyle       =   1
      GraphData       =   0
      GraphData[]     =   0
      LabelText       =   0
      LegendText      =   0
      PatternData     =   0
      SymbolData      =   0
      XPosData        =   0
      XPosData[]      =   0
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Caption :"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   720
      TabIndex        =   12
      Top             =   720
      Width           =   990
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
SCHl = 3
Form5.Show
End Sub

Private Sub Command2_Click()
   On Error Resume Next
  
   CMDialog3.CancelError = -1
    CMDialog3.Flags = &H40&  'PD_PRINTSETUP     ' Just enable print setup
    CMDialog3.DialogTitle = "Print File"
   'CmDialog3.FromPage = 1
   'CmDialog3.ToPage = 1
   CMDialog3.Action = 5
  If Err = 0 Then
Command1.Visible = False
Command2.Visible = False
Command3.Visible = False
Command4.Visible = False
Panel3D1.Visible = False
Option3D1.Visible = False
Option3D2.Visible = False
Option3D3.Visible = False
Label1.Visible = False
Form3.PrintForm
Command1.Visible = True
Command2.Visible = True
Command3.Visible = True
Command4.Visible = True
Panel3D1.Visible = True
Option3D1.Visible = True
Option3D2.Visible = True
Option3D3.Visible = True
Label1.Visible = True
End If
'If (option3d1.Value = True) Then
'graph1.DrawMode = 5
'ElseIf (option3d2.Value = True) Then graph2.DrawMode = 5
'Else
'End If
End Sub

Private Sub Command3_Click()
MDIForm1.Show
Form3.Hide

Zcalsweep (0)

End Sub

Private Sub Command4_Click()
On Error Resume Next
CMDialog3.CancelError = -1
CMDialog3.DialogTitle = "Save File"
CMDialog3.Filter = "S21 Data Files |*.S21"
CMDialog3.Flags = &H4
CMDialog3.Action = 2

points% = Abs(CInt((stop1 - start1) / step1)) + 1
f! = start1
If Err = 0 Then
   Open CMDialog3.Filename For Output As #1
   Print #1, "QuickSmith IL"
   Print #1, "Freq      S21 dB      Angle"
   For I% = 1 To points%
   Print #1, Format(f!, "####0.000") & "    " & Format(S21mag(I%), "###00.00") & "      " & Format(S21Angle(I%), "#00.00")
   f! = f! + step1
   Next I%



   Close #1
End If

End Sub

Private Sub EndMenu_Click()

End Sub

Private Sub Form_Load()
PltopenFlag = True
Height = Screen.Height
Width = Screen.Width
Left = 0
Top = 0
Graph1.Move (Screen.Width - Width) \ 2, (Screen.Height - Height + 2040) \ 2
'Graph2.Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
Form3.Graph2.Width = Form3.Graph1.Width
Form3.Graph2.Height = Form3.Graph1.Height
Form3.Graph2.Left = Form3.Graph1.Left
Form3.Graph2.Top = Form3.Graph1.Top
'list1.Move (Screen.Width - Width) \ 2, (Screen.Height - Height + 2040) \ 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
PltopenFlag = False
End Sub

Private Sub Option3D1_Click(Value As Integer)
If (Option3D1.Value = True) Then
Graph1.Visible = True
Graph2.Visible = False
List1.Visible = False
Panel3D2.Visible = False
Else

Graph1.Visible = False

End If
End Sub

Private Sub Option3D2_Click(Value As Integer)
If (Value = True) Then
Graph2.Visible = True
Graph1.Visible = False
List1.Visible = False
Panel3D2.Visible = False
Else
Graph2.Visible = False
End If
End Sub

Private Sub Option3D3_Click(Value As Integer)
If (Value = True) Then
Graph2.Visible = False
Graph1.Visible = False
List1.Visible = True
Panel3D2.Visible = True
Else
List1.Visible = False
Panel3D2.Visible = False
End If

End Sub

