VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form Form8 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sweep Setup"
   ClientHeight    =   3585
   ClientLeft      =   1905
   ClientTop       =   1725
   ClientWidth     =   4950
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
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3585
   ScaleWidth      =   4950
   Begin Threed.SSOption Option3D10 
      Height          =   252
      Left            =   3240
      TabIndex        =   8
      Top             =   420
      Width           =   1572
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   78
      Caption         =   "Element 9"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   24.27
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSOption Option3D9 
      Height          =   252
      Left            =   1680
      TabIndex        =   7
      Top             =   1500
      Width           =   1692
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   78
      Caption         =   "Element 8"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   24.27
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSOption Option3D8 
      Height          =   252
      Left            =   1680
      TabIndex        =   6
      Top             =   1140
      Width           =   1692
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   78
      Caption         =   "Element 7"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   24.27
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSOption Option3D7 
      Height          =   252
      Left            =   1680
      TabIndex        =   5
      Top             =   780
      Width           =   1572
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   78
      Caption         =   "Element 6"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   24.27
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSOption Option3D6 
      Height          =   252
      Left            =   1680
      TabIndex        =   4
      Top             =   420
      Width           =   1572
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   78
      Caption         =   "Element 5"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   24.27
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSOption Option3D5 
      Height          =   252
      Left            =   120
      TabIndex        =   3
      Top             =   1500
      Width           =   1452
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   78
      Caption         =   "Element 4"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   24.27
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSOption Option3D4 
      Height          =   252
      Left            =   120
      TabIndex        =   2
      Top             =   1140
      Width           =   1452
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   78
      Caption         =   "Element 3"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   24.27
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   288
      Left            =   1500
      TabIndex        =   18
      TabStop         =   0   'False
      Text            =   "Sweep Variable"
      Top             =   60
      Width           =   2052
   End
   Begin Threed.SSOption Option3D3 
      Height          =   252
      Left            =   120
      TabIndex        =   1
      Top             =   780
      Width           =   1452
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   78
      Caption         =   "Element 2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   24.27
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSOption Option3D2 
      Height          =   252
      Left            =   3240
      TabIndex        =   9
      Top             =   780
      Width           =   1632
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   78
      Caption         =   "Element 10"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   24.27
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSOption Option3D1 
      Height          =   252
      Left            =   120
      TabIndex        =   0
      Top             =   420
      Width           =   1452
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   78
      Caption         =   "Frequency"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   24.27
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Value           =   -1  'True
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   288
      Left            =   480
      TabIndex        =   17
      TabStop         =   0   'False
      Text            =   "Units: MHz,pf,nh,ohms.."
      Top             =   3300
      Width           =   3972
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "O K"
      Height          =   372
      Left            =   3060
      TabIndex        =   14
      Top             =   2340
      Width           =   1212
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   288
      Left            =   1560
      TabIndex        =   10
      Top             =   1980
      Width           =   1140
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   288
      Left            =   1560
      TabIndex        =   11
      Top             =   2400
      Width           =   1128
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   288
      Left            =   1560
      TabIndex        =   12
      Top             =   2820
      Width           =   1092
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Start"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   480
      TabIndex        =   16
      Top             =   1980
      Width           =   972
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Stop"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   480
      TabIndex        =   15
      Top             =   2400
      Width           =   948
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Step"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   480
      TabIndex        =   13
      Top             =   2820
      Width           =   948
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim points As Long
R$ = Text1.Text
If (IsNumeric(R$) = False) Then GoTo SweepEnd
If (CDbl(R$) = 0) Then start1 = 0.000000000001 Else start1 = CDbl(R$)

R$ = Text2.Text
If (IsNumeric(R$) = False) Then GoTo SweepEnd
If (CDbl(R$) = 0) Then stop1 = 0.000000000001 Else stop1 = CDbl(R$)


R$ = Text3.Text
If (IsNumeric(R$) = False) Then GoTo SweepEnd

If (CDbl(R$) <= 0) Then GoTo SweepEnd ' zero condition
step1 = CDbl(R$)
points = Abs((stop1 - start1) / step1) + 1
If (points > 1000) Then
step1 = Abs((stop1 - start1)) / 999
MsgBox ("Number of Points exceeds 1000. Step Size Set to " & step1)
End If



SweepEnd:
Form8.Hide
End Sub

Private Sub Form_Activate()
Text1.Text = start1
Text2.Text = stop1
Text3.Text = step1

End Sub

Private Sub Form_Load()
 SweepopenFlag = True
 Form8.Show
 Text1.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
SweepopenFlag = False
End Sub

Private Sub Option3D1_Click(Value As Integer)
Element = 1
End Sub

Private Sub Option3D10_Click(Value As Integer)
Element = 9
End Sub

Private Sub Option3D2_Click(Value As Integer)
Element = 10
End Sub

Private Sub Option3D3_Click(Value As Integer)
Element = 2
End Sub

Private Sub Option3D4_Click(Value As Integer)
Element = 3
End Sub

Private Sub Option3D5_Click(Value As Integer)
Element = 4
End Sub

Private Sub Option3D6_Click(Value As Integer)
Element = 5
End Sub

Private Sub Option3D7_Click(Value As Integer)
Element = 6
End Sub

Private Sub Option3D8_Click(Value As Integer)
Element = 7
End Sub

Private Sub Option3D9_Click(Value As Integer)
Element = 8
End Sub

