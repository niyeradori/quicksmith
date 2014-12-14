VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{7A080CC8-26E2-101B-AEBD-04021C009402}#1.0#0"; "GAUGE32.OCX"
Begin VB.Form Form7 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   4080
   ClientLeft      =   1770
   ClientTop       =   2070
   ClientWidth     =   7665
   FillColor       =   &H00808080&
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
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4080
   ScaleWidth      =   7665
   Begin GaugeLib.Gauge Gauge1 
      Height          =   288
      Left            =   4752
      TabIndex        =   6
      Top             =   3060
      Width           =   2640
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   73
      ForeColor       =   16711680
      Max             =   400
      NeedleWidth     =   1
      Picture         =   "STARTUP.frx":0000
   End
   Begin Threed.SSCommand Command3D1 
      Height          =   948
      Left            =   3252
      TabIndex        =   1
      Top             =   708
      Width           =   3600
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   78
      Caption         =   "QuickSmith"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   8
   End
   Begin Threed.SSPanel Panel3D2 
      Height          =   264
      Left            =   2496
      TabIndex        =   2
      Top             =   -420
      Width           =   1716
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   15
      Caption         =   "Version 1.00"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      BevelOuter      =   0
      Font3D          =   3
   End
   Begin Threed.SSPanel Panel3D1 
      Height          =   3912
      Left            =   0
      TabIndex        =   0
      Top             =   -12
      Width           =   7656
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   15
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   2
      BorderWidth     =   6
      BevelInner      =   1
      Outline         =   -1  'True
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         ForeColor       =   &H80000008&
         Height          =   3168
         Left            =   288
         Picture         =   "STARTUP.frx":0C6E
         ScaleHeight     =   3165
         ScaleWidth      =   4890
         TabIndex        =   4
         Top             =   324
         Width           =   4896
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Copyright. 1995-2014, V.Iyer"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   255
         TabIndex        =   8
         Top             =   3495
         Visible         =   0   'False
         Width           =   2430
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Loading . . .  "
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   5976
         TabIndex        =   7
         Top             =   3420
         Width           =   1488
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Version 5.0.0"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   5550
         TabIndex        =   5
         Top             =   1680
         Width           =   1800
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   600
      Left            =   156
      Top             =   3840
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "Command1"
      Height          =   396
      Left            =   6336
      TabIndex        =   3
      Top             =   420
      Width           =   972
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command3D1_Click()
Timer1.Enabled = False

Load MDIForm1
Unload Form7
MDIForm1.Show
End Sub

Private Sub Command3D1_GotFocus()
Command1.SetFocus
End Sub

Private Sub Form_Load()
 Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
'panel3d1.Left = form7.Left
'panel3d1.Top = form7.Top
Panel3D1.Height = Height
Panel3D1.Width = Width
 'form7.Visible = False

End Sub

Private Sub Panel3D1_Click()
MDIForm1.Show
End Sub

Private Sub Timer1_Timer()

For I% = 1 To 400
Gauge1.Value = I%
Next I%
'label2.Caption = "Done"
Load MDIForm1
MDIForm1.Show
Unload Form7
End Sub

