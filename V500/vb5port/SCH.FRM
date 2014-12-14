VERSION 5.00
Object = "{B16553C3-06DB-101B-85B2-0000C009BE81}#1.0#0"; "SPIN32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form form2 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "Schematic"
   ClientHeight    =   6135
   ClientLeft      =   2040
   ClientTop       =   2010
   ClientWidth     =   9105
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00C0C0C0&
   Icon            =   "SCH.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6135
   ScaleWidth      =   9105
   Begin VB.HScrollBar HScroll1 
      Height          =   276
      Left            =   360
      Max             =   5000
      SmallChange     =   50
      TabIndex        =   27
      Top             =   6600
      Visible         =   0   'False
      Width           =   2028
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6024
      Left            =   0
      ScaleHeight     =   6030
      ScaleWidth      =   12285
      TabIndex        =   28
      Top             =   -240
      Width           =   12288
      Begin VB.VScrollBar VScroll1 
         Height          =   2028
         Left            =   8640
         Max             =   5000
         SmallChange     =   50
         TabIndex        =   26
         Top             =   720
         Visible         =   0   'False
         Width           =   276
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4932
         Left            =   0
         ScaleHeight     =   4935
         ScaleWidth      =   8175
         TabIndex        =   0
         Top             =   240
         Width           =   8172
         Begin Threed.SSPanel Panel3D2 
            Height          =   252
            Left            =   3036
            TabIndex        =   1
            Top             =   432
            Width           =   948
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   15
            Caption         =   "Inches"
            ForeColor       =   4210752
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24.27
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
         End
         Begin Spin.SpinButton Spin5 
            Height          =   264
            Left            =   1044
            TabIndex        =   53
            Top             =   432
            Width           =   252
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   73
            ForeColor       =   4210752
            Delay           =   175
            ShadowBackColor =   12632256
            ShadowThickness =   1
         End
         Begin Spin.SpinButton Spin2 
            Height          =   264
            Index           =   4
            Left            =   2220
            TabIndex        =   54
            Top             =   1110
            Width           =   300
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   73
            ForeColor       =   4210752
            Delay           =   150
            ShadowBackColor =   12632256
            ShadowThickness =   1
         End
         Begin Threed.SSPanel Panel3D3 
            Height          =   252
            Index           =   4
            Left            =   1320
            TabIndex        =   2
            Top             =   1110
            Width           =   915
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   15
            ForeColor       =   4210752
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24.27
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
         End
         Begin Threed.SSPanel Panel3D3 
            Height          =   252
            Index           =   8
            Left            =   3750
            TabIndex        =   3
            Top             =   1110
            Width           =   915
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   15
            ForeColor       =   4210752
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24.27
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
         End
         Begin Spin.SpinButton Spin2 
            Height          =   264
            Index           =   8
            Left            =   4650
            TabIndex        =   55
            Top             =   1110
            Width           =   300
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   73
            ForeColor       =   4210752
            Delay           =   150
            ShadowBackColor =   12632256
            ShadowThickness =   1
         End
         Begin Spin.SpinButton Spin2 
            Height          =   270
            Index           =   10
            Left            =   5820
            TabIndex        =   56
            Top             =   1110
            Width           =   300
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   73
            ForeColor       =   4210752
            Delay           =   150
            ShadowBackColor =   12632256
            ShadowThickness =   1
         End
         Begin Threed.SSPanel Panel3D3 
            Height          =   252
            Index           =   10
            Left            =   4920
            TabIndex        =   4
            Top             =   1110
            Width           =   915
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   15
            ForeColor       =   4210752
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24.27
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
         End
         Begin Spin.SpinButton Spin3 
            Height          =   264
            Left            =   1140
            TabIndex        =   57
            Top             =   2760
            Width           =   300
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   73
            ForeColor       =   4210752
            Delay           =   150
            ShadowBackColor =   12632256
            ShadowThickness =   1
         End
         Begin Spin.SpinButton Spin4 
            Height          =   264
            Index           =   3
            Left            =   2310
            TabIndex        =   58
            Top             =   2520
            Width           =   300
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   73
            ForeColor       =   4210752
            Delay           =   150
            ShadowBackColor =   12632256
            ShadowThickness =   1
         End
         Begin Spin.SpinButton Spin4 
            Height          =   264
            Index           =   5
            Left            =   3510
            TabIndex        =   59
            Top             =   2520
            Width           =   300
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   73
            ForeColor       =   4210752
            Delay           =   150
            ShadowBackColor =   12632256
            ShadowThickness =   1
         End
         Begin Spin.SpinButton Spin4 
            Height          =   264
            Index           =   7
            Left            =   4710
            TabIndex        =   60
            Top             =   2520
            Width           =   300
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   73
            ForeColor       =   4210752
            Delay           =   150
            ShadowBackColor =   12632256
            ShadowThickness =   1
         End
         Begin Spin.SpinButton Spin4 
            Height          =   264
            Index           =   9
            Left            =   5910
            TabIndex        =   61
            Top             =   2520
            Width           =   300
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   73
            ForeColor       =   4210752
            Delay           =   150
            ShadowBackColor =   12632256
            ShadowThickness =   1
         End
         Begin Spin.SpinButton Spin1 
            Height          =   264
            Left            =   1140
            TabIndex        =   62
            Top             =   2520
            Width           =   300
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   73
            ForeColor       =   4210752
            Delay           =   150
            ShadowBackColor =   12632256
            ShadowThickness =   1
         End
         Begin Threed.SSPanel Panel3D4 
            Height          =   252
            Left            =   240
            TabIndex        =   5
            Top             =   2520
            Width           =   915
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   15
            Caption         =   "50.00"
            ForeColor       =   4210752
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24.27
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
         End
         Begin Threed.SSPanel Panel3D5 
            Height          =   252
            Left            =   240
            TabIndex        =   6
            Top             =   2760
            Width           =   915
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   15
            Caption         =   "0.00"
            ForeColor       =   4210752
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24.27
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
         End
         Begin Threed.SSPanel Panel3D6 
            Height          =   252
            Index           =   3
            Left            =   1410
            TabIndex        =   7
            Top             =   2520
            Width           =   915
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   15
            ForeColor       =   4210752
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24.27
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
         End
         Begin Threed.SSPanel Panel3D6 
            Height          =   252
            Index           =   5
            Left            =   2610
            TabIndex        =   8
            Top             =   2520
            Width           =   915
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   15
            ForeColor       =   4210752
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24.27
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
         End
         Begin Threed.SSPanel Panel3D6 
            Height          =   252
            Index           =   7
            Left            =   3810
            TabIndex        =   9
            Top             =   2520
            Width           =   915
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   15
            ForeColor       =   4210752
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24.27
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
         End
         Begin Threed.SSPanel Panel3D6 
            Height          =   252
            Index           =   9
            Left            =   5010
            TabIndex        =   10
            Top             =   2520
            Width           =   915
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   15
            ForeColor       =   4210752
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24.27
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
         End
         Begin Spin.SpinButton Spin2 
            Height          =   264
            Index           =   2
            Left            =   1140
            TabIndex        =   63
            Top             =   1110
            Width           =   300
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   73
            ForeColor       =   4210752
            Delay           =   150
            ShadowBackColor =   12632256
            ShadowThickness =   1
         End
         Begin Threed.SSPanel Panel3D10 
            Height          =   252
            Left            =   72
            TabIndex        =   11
            ToolTipText     =   "Use Component Tuner to set initial value"
            Top             =   432
            Width           =   972
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   15
            Caption         =   "100.00"
            ForeColor       =   4210752
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24.27
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
         End
         Begin Threed.SSPanel Panel3D12 
            Height          =   252
            Left            =   1656
            TabIndex        =   12
            Top             =   432
            Width           =   972
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   15
            ForeColor       =   4210752
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24.27
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
         End
         Begin Threed.SSPanel Panel3D3 
            Height          =   252
            Index           =   2
            Left            =   240
            TabIndex        =   13
            Top             =   1110
            Width           =   915
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   15
            ForeColor       =   4210752
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24.27
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
         End
         Begin Threed.SSPanel Panel3D3 
            Height          =   252
            Index           =   6
            Left            =   2520
            TabIndex        =   14
            Top             =   1110
            Width           =   915
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   15
            ForeColor       =   4210752
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24.27
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
         End
         Begin Spin.SpinButton Spin2 
            Height          =   264
            Index           =   6
            Left            =   3450
            TabIndex        =   64
            Top             =   1110
            Width           =   300
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   73
            ForeColor       =   4210752
            Delay           =   150
            ShadowBackColor =   12632256
            ShadowThickness =   1
         End
         Begin Threed.SSOption Option3D1 
            Height          =   372
            Left            =   2640
            TabIndex        =   15
            Top             =   36
            Width           =   1632
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Tune Length/C"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   16.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   -1  'True
         End
         Begin Spin.SpinButton Spin6 
            Height          =   264
            Left            =   3960
            TabIndex        =   65
            Top             =   432
            Width           =   252
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   73
            ForeColor       =   4210752
            Delay           =   150
            ShadowBackColor =   12632256
            ShadowThickness =   1
         End
         Begin Threed.SSPanel Panel3D8 
            Height          =   255
            Left            =   5268
            TabIndex        =   16
            Top             =   1728
            Width           =   1110
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   15
            ForeColor       =   4210752
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24.27
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
         End
         Begin Threed.SSPanel Panel3D9 
            Height          =   255
            Left            =   5268
            TabIndex        =   17
            Top             =   1980
            Width           =   1110
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   15
            ForeColor       =   4210752
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24.27
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
         End
         Begin Threed.SSCommand Command3D1 
            Height          =   465
            Left            =   6720
            TabIndex        =   18
            ToolTipText     =   "Smith Chart"
            Top             =   120
            Width           =   495
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            BevelWidth      =   1
            Outline         =   0   'False
            AutoSize        =   1
            Picture         =   "SCH.frx":030A
         End
         Begin Threed.SSCommand Command3D2 
            Height          =   465
            Left            =   6240
            TabIndex        =   19
            ToolTipText     =   "Amplifier Design"
            Top             =   120
            Width           =   495
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            BevelWidth      =   1
            Outline         =   0   'False
            Picture         =   "SCH.frx":0624
         End
         Begin Threed.SSCommand Command3D3 
            Height          =   465
            Left            =   7200
            TabIndex        =   20
            ToolTipText     =   "Sweep Graphs"
            Top             =   120
            Width           =   495
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "S21"
            ForeColor       =   16711680
            BevelWidth      =   1
            Outline         =   0   'False
         End
         Begin Threed.SSOption Option3D2 
            Height          =   375
            Left            =   4356
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   36
            Width           =   1665
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Tune Z0/ L/R"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   16.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label4 
            Caption         =   "Drag components to the white areas to build the circuit"
            Height          =   255
            Left            =   240
            TabIndex        =   66
            Top             =   4680
            Width           =   5175
         End
         Begin VB.Image Image16 
            Appearance      =   0  'Flat
            DragMode        =   1  'Automatic
            Height          =   286
            Left            =   60
            Picture         =   "SCH.frx":0A76
            Stretch         =   -1  'True
            Tag             =   "PRC"
            Top             =   3840
            Width           =   864
         End
         Begin VB.Image Image15 
            Appearance      =   0  'Flat
            DragMode        =   1  'Automatic
            Height          =   276
            Left            =   6540
            Picture         =   "SCH.frx":0B60
            Stretch         =   -1  'True
            Tag             =   "SRC"
            Top             =   3300
            Width           =   864
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "pf / nf / ohms"
            ForeColor       =   &H80000008&
            Height          =   252
            Index           =   13
            Left            =   5520
            TabIndex        =   51
            Top             =   3720
            Width           =   1296
         End
         Begin VB.Image Image14 
            Appearance      =   0  'Flat
            DragMode        =   1  'Automatic
            Height          =   276
            Left            =   4740
            Picture         =   "SCH.frx":0C4A
            Stretch         =   -1  'True
            Tag             =   "SLC"
            Top             =   3300
            Width           =   864
         End
         Begin VB.Image Image13 
            Appearance      =   0  'Flat
            DragMode        =   1  'Automatic
            Height          =   286
            Left            =   5640
            Picture         =   "SCH.frx":0D34
            Stretch         =   -1  'True
            Tag             =   "PLC"
            Top             =   3300
            Width           =   876
         End
         Begin VB.Line Line1 
            BorderWidth     =   2
            X1              =   6000
            X2              =   9900
            Y1              =   4560
            Y2              =   4920
         End
         Begin VB.Image Image1 
            Appearance      =   0  'Flat
            DragMode        =   1  'Automatic
            Height          =   252
            Left            =   1764
            Picture         =   "SCH.frx":0E1E
            Stretch         =   -1  'True
            Tag             =   "C"
            Top             =   3312
            Width           =   540
         End
         Begin VB.Image Image2 
            Appearance      =   0  'Flat
            DragMode        =   1  'Automatic
            Height          =   252
            Left            =   2340
            Picture         =   "SCH.frx":0F08
            Stretch         =   -1  'True
            Tag             =   "L"
            Top             =   3312
            Width           =   540
         End
         Begin VB.Image Image3 
            Appearance      =   0  'Flat
            DragMode        =   1  'Automatic
            Height          =   252
            Left            =   612
            Picture         =   "SCH.frx":0FF2
            Stretch         =   -1  'True
            Tag             =   "R"
            Top             =   3300
            Width           =   540
         End
         Begin VB.Image Image4 
            Appearance      =   0  'Flat
            DragMode        =   1  'Automatic
            Height          =   286
            Left            =   3492
            Picture         =   "SCH.frx":10DC
            Stretch         =   -1  'True
            Tag             =   "W"
            Top             =   3312
            Width           =   540
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "ohms"
            ForeColor       =   &H80000008&
            Height          =   252
            Index           =   5
            Left            =   576
            TabIndex        =   22
            Top             =   3576
            Width           =   552
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "nh"
            ForeColor       =   &H80000008&
            Height          =   252
            Index           =   8
            Left            =   2340
            TabIndex        =   23
            Top             =   3576
            Width           =   492
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "pf"
            ForeColor       =   &H80000008&
            Height          =   252
            Index           =   7
            Left            =   1800
            TabIndex        =   24
            Top             =   3576
            Width           =   492
         End
         Begin VB.Image Image5 
            Appearance      =   0  'Flat
            DragMode        =   1  'Automatic
            Height          =   252
            Left            =   1188
            Picture         =   "SCH.frx":1A1E
            Stretch         =   -1  'True
            Tag             =   "X"
            Top             =   3300
            Width           =   540
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "ohms"
            ForeColor       =   &H80000008&
            Height          =   252
            Index           =   6
            Left            =   1152
            TabIndex        =   25
            Top             =   3564
            Width           =   612
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Freq. [MHz]"
            ForeColor       =   &H80000008&
            Height          =   252
            Index           =   0
            Left            =   120
            TabIndex        =   29
            Top             =   120
            Width           =   1212
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "I.L. [dB]"
            ForeColor       =   &H80000008&
            Height          =   252
            Index           =   16
            Left            =   1656
            TabIndex        =   30
            Top             =   120
            Width           =   972
         End
         Begin VB.Image Image6 
            Appearance      =   0  'Flat
            DragMode        =   1  'Automatic
            Height          =   252
            Left            =   2916
            Picture         =   "SCH.frx":2360
            Stretch         =   -1  'True
            Tag             =   "T"
            Top             =   3312
            Width           =   540
         End
         Begin VB.Image Image8 
            Appearance      =   0  'Flat
            DragMode        =   1  'Automatic
            Height          =   912
            Left            =   4392
            Picture         =   "SCH.frx":244A
            Stretch         =   -1  'True
            Tag             =   "S"
            Top             =   3312
            Width           =   300
         End
         Begin VB.Image Image7 
            Appearance      =   0  'Flat
            DragMode        =   1  'Automatic
            Height          =   912
            Left            =   4068
            Picture         =   "SCH.frx":25AC
            Stretch         =   -1  'True
            Tag             =   "O"
            Top             =   3312
            Width           =   276
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "inch"
            ForeColor       =   &H80000008&
            Height          =   252
            Index           =   9
            Left            =   2880
            TabIndex        =   31
            Top             =   3564
            Width           =   552
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "inch"
            ForeColor       =   &H80000008&
            Height          =   252
            Index           =   10
            Left            =   3924
            TabIndex        =   32
            Top             =   4248
            Width           =   576
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "inch"
            ForeColor       =   &H80000008&
            Height          =   312
            Index           =   11
            Left            =   4356
            TabIndex        =   33
            Top             =   4248
            Width           =   552
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FF0000&
            BackStyle       =   0  'Transparent
            Caption         =   "W2"
            BeginProperty Font 
               Name            =   "System"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   192
            Index           =   2
            Left            =   264
            TabIndex        =   34
            Top             =   864
            Width           =   912
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "W6"
            BeginProperty Font 
               Name            =   "System"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   192
            Index           =   6
            Left            =   2496
            TabIndex        =   35
            Top             =   864
            Width           =   912
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "W8"
            BeginProperty Font 
               Name            =   "System"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   192
            Index           =   8
            Left            =   3684
            TabIndex        =   36
            Top             =   864
            Width           =   888
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "W10"
            BeginProperty Font 
               Name            =   "System"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   192
            Index           =   10
            Left            =   4908
            TabIndex        =   37
            Top             =   864
            Width           =   912
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "W3"
            BeginProperty Font 
               Name            =   "System"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   195
            Index           =   3
            Left            =   1440
            TabIndex        =   38
            Top             =   2820
            Width           =   855
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "W5"
            BeginProperty Font 
               Name            =   "System"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   195
            Index           =   5
            Left            =   2580
            TabIndex        =   39
            Top             =   2850
            Width           =   945
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "W7"
            BeginProperty Font 
               Name            =   "System"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   195
            Index           =   7
            Left            =   3840
            TabIndex        =   40
            Top             =   2850
            Width           =   915
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "W9"
            BeginProperty Font 
               Name            =   "System"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   195
            Index           =   9
            Left            =   5040
            TabIndex        =   41
            Top             =   2850
            Width           =   885
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "R"
            BeginProperty Font 
               Name            =   "System"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   252
            Index           =   1
            Left            =   36
            TabIndex        =   42
            Top             =   1728
            Width           =   192
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "System"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   252
            Index           =   2
            Left            =   36
            TabIndex        =   43
            Tag             =   "W"
            Top             =   1980
            Width           =   204
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "W4"
            BeginProperty Font 
               Name            =   "System"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   192
            Index           =   4
            Left            =   1452
            TabIndex        =   44
            Top             =   864
            Width           =   792
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Z-in [Ohms]"
            ForeColor       =   &H80000008&
            Height          =   252
            Index           =   15
            Left            =   5400
            TabIndex        =   45
            Top             =   1476
            Width           =   1092
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "M"
            ForeColor       =   &H000000FF&
            Height          =   192
            Index           =   3
            Left            =   576
            TabIndex        =   46
            Top             =   1860
            Visible         =   0   'False
            Width           =   180
         End
         Begin VB.Image Image9 
            Appearance      =   0  'Flat
            DragMode        =   1  'Automatic
            Height          =   252
            Left            =   36
            Picture         =   "SCH.frx":270E
            Stretch         =   -1  'True
            Tag             =   "G"
            Top             =   3312
            Width           =   540
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "gam"
            ForeColor       =   &H80000008&
            Height          =   252
            Index           =   4
            Left            =   0
            TabIndex        =   47
            Top             =   3576
            Width           =   552
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "System"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   18
            Left            =   4896
            TabIndex        =   48
            Tag             =   "W"
            Top             =   1980
            Width           =   372
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "R"
            BeginProperty Font 
               Name            =   "System"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   252
            Index           =   17
            Left            =   4920
            TabIndex        =   49
            Top             =   1728
            Width           =   372
         End
         Begin VB.Image Image10 
            Appearance      =   0  'Flat
            Height          =   1050
            Left            =   240
            Picture         =   "SCH.frx":27F8
            Top             =   1440
            Width           =   300
         End
         Begin VB.Image Image12 
            Appearance      =   0  'Flat
            DragMode        =   1  'Automatic
            Height          =   1050
            Index           =   3
            Left            =   1200
            Picture         =   "SCH.frx":295A
            Tag             =   "W"
            Top             =   1440
            Width           =   300
         End
         Begin VB.Image Image12 
            Appearance      =   0  'Flat
            DragMode        =   1  'Automatic
            Height          =   1050
            Index           =   5
            Left            =   2160
            Picture         =   "SCH.frx":2ABC
            Tag             =   "W"
            Top             =   1440
            Width           =   300
         End
         Begin VB.Image Image12 
            Appearance      =   0  'Flat
            DragMode        =   1  'Automatic
            Height          =   1050
            Index           =   7
            Left            =   3120
            Picture         =   "SCH.frx":2C1E
            Tag             =   "W"
            Top             =   1440
            Width           =   300
         End
         Begin VB.Image Image12 
            Appearance      =   0  'Flat
            DragMode        =   1  'Automatic
            Height          =   1050
            Index           =   9
            Left            =   4080
            Picture         =   "SCH.frx":2D80
            Tag             =   "W"
            ToolTipText     =   "Click component to change starting value"
            Top             =   1440
            Width           =   300
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Null"
            ForeColor       =   &H80000008&
            Height          =   252
            Index           =   12
            Left            =   3420
            TabIndex        =   50
            Top             =   3564
            Width           =   696
         End
         Begin VB.Image Image11 
            Appearance      =   0  'Flat
            DragMode        =   1  'Automatic
            Height          =   300
            Index           =   2
            Left            =   480
            Picture         =   "SCH.frx":2EE2
            Tag             =   "W"
            Top             =   1440
            Width           =   915
         End
         Begin VB.Image Image11 
            Appearance      =   0  'Flat
            DragMode        =   1  'Automatic
            Height          =   300
            Index           =   4
            Left            =   1440
            Picture         =   "SCH.frx":2FCC
            Tag             =   "W"
            Top             =   1440
            Width           =   915
         End
         Begin VB.Image Image11 
            Appearance      =   0  'Flat
            DragMode        =   1  'Automatic
            Height          =   300
            Index           =   6
            Left            =   2400
            Picture         =   "SCH.frx":30B6
            Tag             =   "W"
            Top             =   1440
            Width           =   915
         End
         Begin VB.Image Image11 
            Appearance      =   0  'Flat
            DragMode        =   1  'Automatic
            Height          =   300
            Index           =   8
            Left            =   3360
            Picture         =   "SCH.frx":31A0
            Tag             =   "W"
            Top             =   1440
            Width           =   915
         End
         Begin VB.Image Image11 
            Appearance      =   0  'Flat
            DragMode        =   1  'Automatic
            Height          =   300
            Index           =   10
            Left            =   4320
            Picture         =   "SCH.frx":328A
            Tag             =   "W"
            Top             =   1440
            Width           =   915
         End
      End
      Begin VB.Label Label6 
         Caption         =   "Click a component to the change value, double click value to change step size"
         Height          =   255
         Left            =   240
         TabIndex        =   68
         Top             =   5640
         Width           =   7095
      End
      Begin VB.Label Label5 
         Caption         =   "Drag the Null part on top of a component to remove it"
         Height          =   255
         Left            =   240
         TabIndex        =   67
         Top             =   5280
         Width           =   5175
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Print Label"
         ForeColor       =   &H80000008&
         Height          =   192
         Left            =   120
         TabIndex        =   52
         Top             =   4520
         Visible         =   0   'False
         Width           =   900
      End
   End
   Begin MSComDlg.CommonDialog CMDialog2 
      Left            =   3450
      Top             =   4830
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FontSize        =   0
      MaxFileSize     =   256
   End
   Begin VB.Menu FileMenu 
      Caption         =   "&File"
      Begin VB.Menu NewItem 
         Caption         =   "&New"
      End
      Begin VB.Menu OpenItem 
         Caption         =   "&Open Schematic"
      End
      Begin VB.Menu SaveAsMenu 
         Caption         =   "&Save Schematic As"
      End
      Begin VB.Menu PrSchItem 
         Caption         =   "&Print Schematic"
      End
      Begin VB.Menu EditItem 
         Caption         =   "&Edit Schematic File"
      End
      Begin VB.Menu MenuBar 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu ExitMenu 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu SweepMenu 
      Caption         =   "&Sweep Setup"
   End
   Begin VB.Menu ValueMenu 
      Caption         =   "&Assign Values"
      Begin VB.Menu FreqItem 
         Caption         =   "&Frequency"
         Shortcut        =   ^F
      End
      Begin VB.Menu VFItem 
         Caption         =   "&VelocityFactor"
         Shortcut        =   ^V
      End
      Begin VB.Menu CIItem 
         Caption         =   "&Char. Impedance"
         Shortcut        =   ^C
      End
      Begin VB.Menu TLineFreqItem 
         Caption         =   "&T-Line Design Freq."
         Shortcut        =   ^T
      End
      Begin VB.Menu EQItem 
         Caption         =   "Element &2: Q"
         Index           =   2
      End
      Begin VB.Menu EQItem 
         Caption         =   "Element &3: Q"
         Index           =   3
      End
      Begin VB.Menu EQItem 
         Caption         =   "Element &4: Q"
         Index           =   4
      End
      Begin VB.Menu EQItem 
         Caption         =   "Element &5: Q"
         Index           =   5
      End
      Begin VB.Menu EQItem 
         Caption         =   "Element &6 : Q"
         Index           =   6
      End
      Begin VB.Menu EQItem 
         Caption         =   "Element &7: Q"
         Index           =   7
      End
      Begin VB.Menu EQItem 
         Caption         =   "Element &8 : Q"
         Index           =   8
      End
      Begin VB.Menu EQItem 
         Caption         =   "Element &9: Q"
         Index           =   9
      End
      Begin VB.Menu EQItem 
         Caption         =   "Element &10: Q"
         Index           =   10
      End
      Begin VB.Menu bax 
         Caption         =   "-"
      End
      Begin VB.Menu DefItem 
         Caption         =   "&Set To Default"
      End
   End
   Begin VB.Menu TetminateMenu 
      Caption         =   "&Termination"
      Begin VB.Menu ImpeItem 
         Caption         =   "&Impedance"
         Checked         =   -1  'True
      End
      Begin VB.Menu GammaItem 
         Caption         =   "&Gamma"
      End
      Begin VB.Menu SepItem 
         Caption         =   "-"
      End
      Begin VB.Menu SingItem 
         Caption         =   "&Single"
      End
      Begin VB.Menu MulItem 
         Caption         =   "&Multiple"
         Begin VB.Menu editterItem 
            Caption         =   "&Edit "
         End
         Begin VB.Menu LinkItem 
            Caption         =   "&Link"
         End
      End
   End
   Begin VB.Menu ILMenu 
      Caption         =   "S&21"
      Begin VB.Menu SweepItem 
         Caption         =   "&Begin Sweep"
      End
   End
   Begin VB.Menu WindowMenu 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu Winitem 
         Caption         =   "&Cascade"
         Index           =   0
      End
      Begin VB.Menu Winitem 
         Caption         =   "Tile &Horizontal"
         Index           =   1
      End
      Begin VB.Menu Winitem 
         Caption         =   "Tile &Vertical"
         Index           =   2
      End
      Begin VB.Menu Winitem 
         Caption         =   "&Arrange Icons"
         Index           =   3
      End
      Begin VB.Menu Winitem 
         Caption         =   "&Normal"
         Index           =   4
      End
      Begin VB.Menu RedrawMenu 
         Caption         =   "&Refresh"
      End
      Begin VB.Menu baritem 
         Caption         =   "-"
      End
      Begin VB.Menu PlotShowItem 
         Caption         =   "Show &Graph Window"
      End
      Begin VB.Menu EditorItem 
         Caption         =   "Show &Editor"
      End
   End
   Begin VB.Menu Help2Item 
      Caption         =   "&Help"
      Begin VB.Menu contentsItem 
         Caption         =   "&Contents"
      End
      Begin VB.Menu ExamplesItem 
         Caption         =   "&Examples"
      End
      Begin VB.Menu AboutHelpItem 
         Caption         =   "About &Help"
      End
      Begin VB.Menu aboutbar 
         Caption         =   "-"
      End
      Begin VB.Menu AboutItem 
         Caption         =   "A&bout..."
      End
   End
End
Attribute VB_Name = "form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit

Private Sub AboutHelpItem_Click()
Dim dummy As String
 R = WinHelp(hwnd, dummy, HELP_HELPONHELP, CLng(0))

End Sub

Private Sub AboutItem_Click()
AboutBox.Show 1  'modal
End Sub

Private Sub ChartItem_Click()
form1.Show

End Sub

Private Sub CIItem_Click()
R$ = InputBox$("Enter Characteristic Impedance of the Chart", "QuickSmith", CStr(Z0))
If (IsNumeric(R$) = False) Then GoTo EndZ0
If (CDbl(R$) <= 0) Then GoTo EndZ0
Z0 = CDbl(R$)
If (Z0 = 0) Then Z0 = 0.00000000001 ' avoid divide by zero error
form1.Smith1.Z0 = Z0
Zcalsweep (0)
EndZ0:

End Sub

Private Sub Command3D1_Click()
form1.Show
form1.SetFocus
End Sub

Private Sub Command3D2_Click()
   Form6.Show
   MDIForm1.Hide

End Sub

Private Sub Command3D3_Click()
 Call sweepItem_click
End Sub

Private Sub ContentsItem_Click()
R = WinHelp(form1.hwnd, HelpFilePath, HELP_CONTEXT, CLng(1))
End Sub

Private Sub DefItem_Click()
Dim I%
VF = 1
Z0 = 50
TLfreq = 100
For I% = 2 To 10
Q(I%) = 1000000
z(I%) = 50
Next I%
 Zcalsweep (0)

MsgBox "Velocity factor Set to 1" & CRLF & "Chart/Line Characteristic Impedance Set to 50 Ohms" & CRLF & "T-Line Design Frequency Set to 100 MHz" & CRLF & "Element Q Factors Set to 1000000"
End Sub

Private Sub EditItem_Click()
 SCHl = 1
 Form5.Show

End Sub

Private Sub EditorItem_Click()
Form5.Show
End Sub

Private Sub editterItem_Click()
 SCHl = 0
 Form5.Show


End Sub

Private Sub EleItem_Click(index As Integer)

End Sub

Private Sub EQItem_Click(index As Integer)
temp$ = "Enter Quality Factor For Element " & index
R$ = InputBox$(temp$, "QuickSmith", CStr(Q(index)))
If (IsNumeric(R$) = False) Then GoTo EndQ
If (CDbl(R$) > 0) Then Q(index) = CDbl(R$)
Zcalsweep (0)
EndQ:

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

Private Sub Form_Activate()
For I% = 2 To 10 Step 2
form2.Panel3D3(I%) = form4.Panel3D3(I%)
'form2.Label1(i%).Caption = form4.Image11(i%).Tag & i%
Next I%
For I% = 3 To 9 Step 2
form2.Panel3D6(I%) = form4.Panel3D6(I%)
'form2.Label1(i%).Caption = form4.Image12(i%).Tag & i%
Next I%
'form2.Panel3D10 = form4.Panel3D10

End Sub

Private Sub Form_Load()

freq = 100     ' in MHz
StepFF = 1#
StepLR = 1
StepLI = 1
For I% = 0 To 14 Step 1
stepS(I%) = 1
stepP(I%) = 1
s(I%) = 0
p(I%) = 0
SwpStp(I%) = 0
Q(I%) = 1000000#
z(I%) = 50
Next
ZR(1) = 50 ' load
ZI(1) = 0
gammaM = 0
GammaA = 0
start1 = 1
stop1 = 100
step1 = 1
SwpStp(1) = step1
VF = 1
Z0 = 50
TL = 0
TLfreq = 100 ' in MHz
Element = 1
MULTITER = 0
Call Zcalsweep(0)
CurrentDir = App.Path ' CurDir$
'MsgBox CurrentDir
form2.Picture1.Move 0, 0 'ScaleWidth, HScroll1.Top
form2.Picture2.Move 0, 0

For I% = 2 To 8 Step 2
Panel3D3(I%).Width = Image11(I%).Width
Panel3D6(I% + 1).Width = Image11(I%).Width
Spin2(I%).Width = Image12(I% + 1).Width
Spin4(I% + 1).Width = Image12(I% + 1).Width
Next I%
Panel3D3(10).Width = Image11(10).Width
Spin2(10).Width = Image12(9).Width
Panel3D4.Width = Image11(2).Width
Panel3D5.Width = Image11(2).Width
Spin1.Width = Image10.Width
Spin3.Width = Spin1.Width

Image11(2).Left = Image10.Left + Image10.Width
For I% = 3 To 9 Step 2
Image12(I%).Left = Image11(I% - 1).Left + Image11(I% - 1).Width
Image11(I% + 1).Left = Image12(I%).Left + Image12(I%).Width
Next I%
Image11(10).Left = Image12(9).Left + Image12(9).Width
Line1.X1 = Image10.Left + Image10.Width / 2
Line1.Y1 = Image10.Top + Image10.Height
Line1.X2 = Image11(8).Left + Image11(8).Width + Image11(10).Width + Image12(9).Width
Line1.Y2 = Image11(10).Top + Image10.Height

Label1(17).Left = Image11(10).Left + Image11(10).Width * 0.8
Label1(18).Left = Image11(10).Left + Image11(10).Width * 0.8
Panel3D8.Left = Label1(17).Left + Label1(17).Width
Panel3D9.Left = Label1(18).Left + Label1(18).Width
Label1(15).Left = Panel3D8.Left

Panel3D3(2).Left = Image10.Left
Spin2(2).Left = Panel3D3(2).Left + Panel3D3(2).Width
Panel3D4.Left = Panel3D3(2).Left
Panel3D5.Left = Panel3D3(2).Left
Spin1.Left = Spin2(2).Left
Spin3.Left = Spin2(2).Left
Label2(2).Left = Panel3D3(2).Left
For I% = 3 To 9 Step 2
Panel3D6(I%).Left = Image12(I%).Left
Spin4(I%).Left = Panel3D6(I%).Left + Panel3D6(I%).Width
Panel3D3(I% + 1).Left = Image12(I%).Left
Spin2(I% + 1).Left = Panel3D3(I% + 1).Left + Panel3D3(I% + 1).Width
Label2(I% + 1).Left = Panel3D3(I% + 1).Left
Label2(I%).Left = Panel3D6(I%).Left
Next I%
Label5.Top = Label4.Top + 300
Label6.Top = Label5.Top + 300
End Sub

Private Sub Form_Resize()
Dim OldScaleMode As Integer
On Error Resume Next
 OldScaleMode = form2.ScaleMode
 form2.ScaleMode = 1 'Twips
  'MsgBox scaleheight
   
  Picture1.Move 0, 0, ScaleWidth, ScaleHeight

   ' Enable or Disable HScroll1 depending on new size of Form
  HScroll1.Value = 0
  HDiff = Picture1.ScaleWidth - Picture2.Width
  HScroll1.Visible = HDiff <= 0
  If HScroll1.Visible Then
    If (form2.VScroll1.Visible) Then
        form2.HScroll1.Move 0, form2.ScaleHeight - form2.HScroll1.Height, form2.ScaleWidth - form2.VScroll1.Width
    Else
        form2.HScroll1.Move 0, form2.ScaleHeight - form2.HScroll1.Height, form2.ScaleWidth
    End If
      
'      HScroll1.Move 0, ScaleHeight - HScroll1.Height, ScaleWidth
   HScroll1.max = Abs(HDiff)
   HScroll1.LargeChange = HScroll1.max \ 10
  End If
  VScroll1.Value = 0
  VDiff = Picture1.ScaleHeight - Picture2.Height
  VScroll1.Visible = VDiff <= 0
  
  If VScroll1.Visible Then
    If (form2.HScroll1.Visible) Then
     form2.VScroll1.Move form2.ScaleWidth - form2.VScroll1.Width, 0, form2.VScroll1.Width, form2.ScaleHeight - form2.HScroll1.Height
    Else
     form2.VScroll1.Move form2.ScaleWidth - form2.VScroll1.Width, 0, form2.VScroll1.Width, form2.ScaleHeight
    End If
    'VScroll1.Move form2.ScaleWidth - form2.VScroll1.Width, 0, form2.VScroll1.Width, form2.ScaleHeight - form2.HScroll1.Height
    VScroll1.max = Abs(VDiff)
    VScroll1.LargeChange = VScroll1.max \ 10
  End If
 form2.ScaleMode = OldScaleMode ' Restore defaults
End Sub

Private Sub FreqItem_Click()
If (freq <= 0.000000000001) Then Fr$ = "0" Else Fr$ = CStr(freq)
R$ = InputBox$("Enter Frequency in MHz", "QuickSmith", Fr$)
If (IsNumeric(R$) = False) Then GoTo Endfreq
If (CDbl(R$) < 0) Then GoTo Endfreq
If (CDbl(R$) > 0) Then freq = CDbl(R$) Else freq = 0.000000000001
Zcalsweep (0)
Panel3D10.Caption = Format$(freq, "###0.000")
If (freq > 9999.99) Then Panel3D10.Caption = Format$(freq, "##.#E+##")

Endfreq:

End Sub

Private Sub GammaItem_Click()
ImpeItem.Checked = False
GammaItem.Checked = True
Label1(1).Caption = "M"
Label1(2).Caption = "A"
gammaM = ZtoGammaM(CDbl(ZR(1)), CDbl(ZI(1)))
GammaA = ZtoGammaA(CDbl(ZR(1)), CDbl(ZI(1)))
Panel3D4.Caption = Format$(gammaM, "0.000")
Panel3D5.Caption = Format$(GammaA, "0.000")
End Sub

Private Sub HScroll1_Change()
Picture2.Left = -HScroll1.Value
End Sub

Private Sub Image10_DragDrop(Source As Control, X As Single, Y As Single)
  
  Select Case Source.Tag
         
        Case "R"

Y1R6:  If (ImpeItem.Checked) Then
          inp$ = InputBox$("Enter Value of Load Resistance in Ohms", "QuickSmith", "")
            If (IsNumeric(inp$)) Then
             ZR(1) = CDbl(inp$)
             If (Abs(CDbl(inp$)) > 99999.99) Then Panel3D4.Caption = Format$(CDbl(inp$), "##.#E+##") Else Panel3D4.Caption = Format$(CDbl(inp$), "0.000")
             Label1(1) = "R"
            End If
       
       Else: Beep
       End If
       
       Case "X"
           If (ImpeItem.Checked) Then
             inp$ = InputBox$("Enter Value of  Load Reactance in Ohms", "QuickSmith", "")
             If (IsNumeric(inp$)) Then
              ZI(1) = CDbl(inp$)
              If (Abs(CDbl(inp$)) > 99999.99) Then Panel3D5.Caption = Format$(CDbl(inp$), "##.#E+##") Else Panel3D5.Caption = Format$(CDbl(inp$), "0.000")
              Label1(2) = "X"
             End If
           
           Else: Beep
           End If


        Case "G"
        If (GammaItem.Checked) Then
            inp$ = InputBox$("Enter Value of Gamma Magnitude", "QuickSmith", "")
            If (inp$ = "") Then GoTo ENDGAMMA
            If (IsNumeric(inp$)) Then
            gammaM = CDbl(inp$)
           If (Abs(CDbl(inp$)) > 99999.99) Then Panel3D5.Caption = Format$(CDbl(inp$), "##.#E+##") Else Panel3D4.Caption = Format$(CDbl(inp$), "0.000")
            Label1(1) = "M"
            
            
            inp$ = InputBox$("Enter Value of Gamma Angle in Degrees", "QuickSmith", "")
              If (inp$ = "") Then GoTo ENDGAMMA
              If (IsNumeric(inp$)) Then
              GammaA = CDbl(inp$)
             If (Abs(CDbl(inp$)) > 99999.99) Then Panel3D5.Caption = Format$(CDbl(inp$), "##.#E+##") Else Panel3D5.Caption = Format$(CDbl(inp$), "0.000")
              Label1(2) = "A"
              End If
            End If
        
ENDGAMMA:
         Else: Beep
         End If
        Case "C", "L", "T", "S", "O"
            Beep

        Case "W"
        ZR(1) = 50
        ZI(1) = 0
        gammaM = 0
        gammaI = 0
        If (GammaItem.Checked) Then
           Label1(1) = "M"
           Label1(2) = "A"
           Panel3D4.Caption = "0.00"
           Panel3D5.Caption = "0.00"
        Else
           Label1(1) = "R"
           Label1(2) = "X"
           Panel3D4.Caption = "50.00"
           Panel3D5.Caption = "0.00"
        End If
        
        Image10.Tag = "W"
        Label1(2).Tag = "W"
        End Select
    Call Zcalsweep(0)

End Sub

Private Sub Image11_DragDrop(index As Integer, Source As Control, X As Single, Y As Single)
 On Error Resume Next
 Select Case Source.Tag
        
        Case "R"
            inp$ = InputBox$("Enter Value of Series Resistance in Ohms", "QuickSmith", "")
            If (IsNumeric(inp$)) Then
            s(index) = CDbl(inp$)  ' will be trapped in ZcalSweep(0)
            Call SeriesIO(CVar(Source.Tag), s(index), index)
            End If
        
        Case "C"
           inp$ = InputBox$("Enter Value of Series Capacitance in pf", "QuickSmith", "")
           If (IsNumeric(inp$)) Then
           s(index) = CDbl(inp$)  ' will be trapped in ZcalSweep(0)
           Call SeriesIO(CVar(Source.Tag), s(index), index)
           End If
        
        Case "L"
           inp$ = InputBox$("Enter Value of Series Inductance in nh", "QuickSmith", "")
           If (IsNumeric(inp$)) Then
           s(index) = CDbl(inp$)    ' will be trapped in ZcalSweep(0)
           Call SeriesIO(CVar(Source.Tag), s(index), index)
           End If
        
        Case "W"
           s(index) = 0
           Q(index) = 1000000
           Call SeriesIO(CVar(Source.Tag), s(index), index)
        Case "X"
            inp$ = InputBox$("Enter Value of Series Reactance in Ohms", "QuickSmith", "")
            If (IsNumeric(inp$)) Then
            s(index) = CDbl(inp$)
           Call SeriesIO(CVar(Source.Tag), s(index), index)
            End If
        
        Case "T"
           msg$ = "Enter Characteristic Impedance of the Line"
           in1$ = InputBox$(msg$, "QuickSmith", CStr(z(index)))
           If in1$ = "" Then GoTo EndSerSel
           If (IsNumeric(in1$) = False) Then GoTo EndSerSel
            Select Case TL
            Case 0
            msg$ = "Enter Slug Length in Inches" ' P(I%) in inches
            Case 1
              msg$ = "Enter Slug Length in mm" ' P(I%) in MM
            Case 2
              msg$ = "Enter Slug Length in Degrees" ' P(I%) in Degrees
            Case 3
              msg$ = "Enter Slug Length in Wavelengths" ' P(I%) in lamda
            Case 4
              msg$ = "Enter Slug Length in meters" ' P(I%) in meters
            
            Case Else
             msg$ = "Enter Slug Length in Inches" ' P(I%) in inches
            End Select
           
           inp$ = InputBox$(msg$, "QuickSmith", "")
           If (IsNumeric(inp$)) Then
           s(index) = CDbl(inp$)   ' will be trapped in ZcalSweep(0)
           z(index) = Abs(CDbl(in1$))   ' will be trapped in assign
           'If (Z(Index) <= 0) Then Z(Index) = .000000000001
           Call SeriesIO(CVar(Source.Tag), s(index), index)
           Call Option3D1_Click(True)
           Option3D1.SetFocus
           End If

        Case "SLC"
            msg$ = "Enter Value of the Capacitance [pf] in the Series LC Element"
            in1$ = InputBox$(msg$, "QuickSmith", "")
            If in1$ = "" Then GoTo EndSerSel
            If (IsNumeric(in1$) = False) Then GoTo EndSerSel
            msg$ = "Enter Value of the Inductance [nh] in the Series LC Element"
            inp$ = InputBox$(msg$, "QuickSmith", "")
            If (IsNumeric(inp$)) Then
              z(index) = CDbl(inp$)  ' will be trapped in ZcalSweep(0)
              s(index) = CDbl(in1$)   ' will be trapped in assign
              Call SeriesIO(CVar(Source.Tag), s(index), index)
              Call Option3D1_Click(True)
              Option3D1.SetFocus
            End If
        Case "PLC"
            msg$ = "Enter Value of the Capacitance [pf] in the Parallel LC Element"
            in1$ = InputBox$(msg$, "QuickSmith", "")
            If in1$ = "" Then GoTo EndSerSel
            If (IsNumeric(in1$) = False) Then GoTo EndSerSel
            msg$ = "Enter Value of the Inductance [nh] in the Parallel LC Element"
            inp$ = InputBox$(msg$, "QuickSmith", "")
            If (IsNumeric(inp$)) Then
              z(index) = CDbl(inp$)  ' will be trapped in ZcalSweep(0)
              s(index) = CDbl(in1$)   ' will be trapped in assign
              Call SeriesIO(CVar(Source.Tag), s(index), index)
              Call Option3D1_Click(True)
              Option3D1.SetFocus
            End If
        Case "SRC"
            msg$ = "Enter Value of the Capacitance [pf] in the Series RC Element"
            in1$ = InputBox$(msg$, "QuickSmith", "")
            If in1$ = "" Then GoTo EndSerSel
            If (IsNumeric(in1$) = False) Then GoTo EndSerSel
            msg$ = "Enter Value of the Resistance [Ohms] in the Series RC Element"
            inp$ = InputBox$(msg$, "QuickSmith", "")
            If (IsNumeric(inp$)) Then
              z(index) = CDbl(inp$)  ' will be trapped in ZcalSweep(0)  res
              s(index) = CDbl(in1$) ' will be trapped in assign   cap
              Call SeriesIO(CVar(Source.Tag), s(index), index)
              Call Option3D1_Click(True)
              Option3D1.SetFocus
            End If

        Case "PRC"
            msg$ = "Enter Value of the Capacitance [pf] in the Parallel RC Element"
            in1$ = InputBox$(msg$, "QuickSmith", "")
            If in1$ = "" Then GoTo EndSerSel
            If (IsNumeric(in1$) = False) Then GoTo EndSerSel
            msg$ = "Enter Value of the Resistance [Ohms] in the Parallel RC Element"
            inp$ = InputBox$(msg$, "QuickSmith", "")
            If (IsNumeric(inp$)) Then
              z(index) = CDbl(inp$)  ' will be trapped in ZcalSweep(0)  res
              s(index) = CDbl(in1$) ' will be trapped in assign   cap
              Call SeriesIO(CVar(Source.Tag), s(index), index)
              Call Option3D1_Click(True)
              Option3D1.SetFocus
            End If

        Case "S", "O", "G"
             Beep
EndSerSel:
     End Select
 
 
 'stepS(Index) = 1
 Call Zcalsweep(0)

End Sub

Private Sub Image12_DragDrop(index As Integer, Source As Control, X As Single, Y As Single)
        On Error Resume Next
        Select Case Source.Tag

        Case "R"
            inp$ = InputBox$("Enter Value of Parallel Resistance in Ohms", "QuickSmith", "")
            If (IsNumeric(inp$)) Then
            p(index) = CDbl(inp$)  ' will be trapped in ZcalSweep(0)
            Call ParallelIO(CVar(Source.Tag), p(index), index)
            End If
        
        Case "C"
            inp$ = InputBox$("Enter Value of Parallel Capacitance in pf", "QuickSmith", "")
            If (IsNumeric(inp$)) Then
            p(index) = CDbl(inp$)  ' will be trapped in ZcalSweep(0)
             Call ParallelIO(CVar(Source.Tag), p(index), index)
            End If
        
        Case "L"
            inp$ = InputBox$("Enter Value of Parallel Inductance in nh", "QuickSmith", "")
            If (IsNumeric(inp$)) Then
            p(index) = CDbl(inp$)  ' will be trapped in ZcalSweep(0)
              Call ParallelIO(CVar(Source.Tag), p(index), index)
            End If
        
        Case "W"
           Panel3D6(index).Caption = ""
           Q(index) = 1000000
           p(index) = 0
           Call ParallelIO(CVar(Source.Tag), p(index), index)
        
        Case "X"
           inp$ = InputBox$("Enter Value of Parallel Reactance in Ohms", "QuickSmith", "")
           If (IsNumeric(inp$)) Then
           p(index) = CDbl(inp$)  ' will be trapped in ZcalSweep(0)
           Call ParallelIO(CVar(Source.Tag), p(index), index)
           End If
        
        Case "O"
            msg$ = "Enter Characteristic Impedance of the Line"
            in1$ = InputBox$(msg$, "QuickSmith", CStr(z(index)))
            If in1$ = "" Then GoTo EndParSel
            If (IsNumeric(in1$) = False) Then GoTo EndParSel
            Select Case TL
            Case 0
            msg$ = "Enter Open-Circuited Stub Length in Inches" ' P(I%) in inches
            Case 1
             msg$ = "Enter Open-Circuited Stub Length in mm" ' P(I%) in MM
            Case 2
            msg$ = "Enter Open-Circuited Stub Length in Degrees" ' P(I%) in Degrees
            Case 3
            msg$ = "Enter Open-Circuited  Slug Length in Wavelengths" ' P(I%) in lamda
            Case 4
            msg$ = "Enter Open-Circuited Slug Length in meters" ' P(I%) in meterse
            
            Case Else
             msg$ = "Enter Open-Circuited Stub Length in Inches" ' P(I%) in inches
            End Select
           
           inp$ = InputBox$(msg$, "QuickSmith", "")
           If (IsNumeric(inp$)) Then
           p(index) = CDbl(inp$)  ' will be trapped in ZcalSweep(0)
           z(index) = Abs(CDbl(in1$))   ' will be trapped in assign
           'If (Z(Index) <= 0) Then Z(Index) = .000000000001
           Call ParallelIO(CVar(Source.Tag), p(index), index)
           Call Option3D1_Click(True)
           Option3D1.SetFocus
           End If
        
        Case "S"
            msg$ = "Enter Characteristic Impedance of the Line"
            in1$ = InputBox$(msg$, "QuickSmith", CStr(z(index)))
            If in1$ = "" Then GoTo EndParSel
            If (IsNumeric(in1$) = False) Then GoTo EndParSel
            Select Case TL
            Case 0
            msg$ = "Enter Short-Circuited Stub Length in Inches" ' P(I%) in inches
            Case 1
            msg$ = "Enter Short-Circuited Stub Length in mm" ' P(I%) in MM
            Case 2
            msg$ = "Enter Short-Circuited Stub Length in Degrees" ' P(I%) in Degrees
            Case 3
            msg$ = "Enter Short-Circuited  Slug Length in Wavelengths" ' P(I%) in lamda
            Case 4
            msg$ = "Enter Short-Circuited Slug Length in meters" ' P(I%) in meterse
           
            Case Else
            msg$ = "Enter Short-Circuited Stub Length in Inches" ' P(I%) in inches
            End Select
           
           inp$ = InputBox$(msg$, "QuickSmith", "")
           If (IsNumeric(inp$)) Then
           p(index) = CDbl(inp$)  ' will be trapped in ZcalSweep(0)
           z(index) = Abs(CDbl(in1$))   ' will be trapped in assign
           'If (Z(Index) <= 0) Then Z(Index) = .000000000001
           Call ParallelIO(CVar(Source.Tag), p(index), index)
           Call Option3D1_Click(True)
           Option3D1.SetFocus
           End If
        Case "SLC"
            msg$ = "Enter Value of the Capacitance [pf] in the Series LC Element"
            in1$ = InputBox$(msg$, "QuickSmith", "")
            If in1$ = "" Then GoTo EndParSel
            If (IsNumeric(in1$) = False) Then GoTo EndParSel
            msg$ = "Enter Value of the Inductance [nh] in the Series LC Element"
            inp$ = InputBox$(msg$, "QuickSmith", "")
            If (IsNumeric(inp$)) Then
              z(index) = CDbl(inp$)  ' will be trapped in ZcalSweep(0)
              p(index) = CDbl(in1$)   ' will be trapped in assign
              Call ParallelIO(CVar(Source.Tag), p(index), index)
              Call Option3D1_Click(True)
              Option3D1.SetFocus
            End If


        Case "PLC"
            msg$ = "Enter Value of the Capacitance [pf] in the Parallel LC Element"
            in1$ = InputBox$(msg$, "QuickSmith", "")
            If in1$ = "" Then GoTo EndParSel
            If (IsNumeric(in1$) = False) Then GoTo EndParSel
            msg$ = "Enter Value of the Inductance [nh] in the Parallel LC Element"
            inp$ = InputBox$(msg$, "QuickSmith", "")
            If (IsNumeric(inp$)) Then
              z(index) = CDbl(inp$)  ' will be trapped in ZcalSweep(0)
              p(index) = CDbl(in1$)   ' will be trapped in assign
              Call ParallelIO(CVar(Source.Tag), p(index), index)
              Call Option3D1_Click(True)
              Option3D1.SetFocus
             End If
        Case "SRC"
            msg$ = "Enter Value of the Capacitance [pf] in the Series RC Element"
            in1$ = InputBox$(msg$, "QuickSmith", "")
            If in1$ = "" Then GoTo EndParSel
            If (IsNumeric(in1$) = False) Then GoTo EndParSel
            msg$ = "Enter Value of the Resistance [Ohms] in the Series RC Element"
            inp$ = InputBox$(msg$, "QuickSmith", "")
            If (IsNumeric(inp$)) Then
              z(index) = CDbl(inp$)  ' will be trapped in ZcalSweep(0)
              p(index) = CDbl(in1$)   ' will be trapped in assign
              Call ParallelIO(CVar(Source.Tag), p(index), index)
              Call Option3D1_Click(True)
              Option3D1.SetFocus
            End If
        Case "PRC"
            msg$ = "Enter Value of the Capacitance [pf] in the Parallel RC Element"
            in1$ = InputBox$(msg$, "QuickSmith", "")
            If in1$ = "" Then GoTo EndParSel
            If (IsNumeric(in1$) = False) Then GoTo EndParSel
            msg$ = "Enter Value of the Resistance [Ohms] in the Parallel RC Element"
            inp$ = InputBox$(msg$, "QuickSmith", "")
            If (IsNumeric(inp$)) Then
              z(index) = CDbl(inp$)  ' will be trapped in ZcalSweep(0)res
              p(index) = CDbl(in1$)   ' will be trapped in assign cap
              Call ParallelIO(CVar(Source.Tag), p(index), index)
              Call Option3D1_Click(True)
              Option3D1.SetFocus
            End If


        Case "T", "G"
           Beep
EndParSel:
        End Select
      
      'stepP(Index) = 1
      Call Zcalsweep(0)

End Sub

Private Sub ImpeItem_Click()
ImpeItem.Checked = True
GammaItem.Checked = False
Label1(1).Caption = "R"
Label1(2).Caption = "X"
gammaM = GToZR(CDbl(gammaM), CDbl(GammaA))
GammaA = GToZI(CDbl(gammaM), CDbl(GammaA))
Panel3D4.Caption = Format$(ZR(1), "0.000")
Panel3D5.Caption = Format$(ZI(1), "0.000")

End Sub

Private Sub LinkItem_Click()
   Dim temp, phase, prevphase, offset
   Dim Filedata0 As Double
   Dim Filedata1 As Double
   Dim Filedata2 As Double
   Dim gamh As Integer ' used as glag to detect HZ file
   gamh = 0
   On Error Resume Next
   CMDialog2.CancelError = -1
   CMDialog2.DialogTitle = "Open File"
   CMDialog2.Filter = "Load Temination Files|*.Gam "
   CMDialog2.Flags = &H1
   CMDialog2.Filename = "*.Gam"
   CMDialog2.Action = 1
   If Err = 0 Then
    Open CMDialog2.Filename For Input As #1
    Input #1, temp
    If (temp <> "QuickSmith GAM") Then
        If (temp = "QuickSmith GAMH") Then
            gmah = 1
            GoTo label102
        End If
        Beep
        MsgBox "Bad Gamma File !" & CRLF & " Check QuickSmith Data File Format ", 48, "QuickSmith"
        Close #1
        Exit Sub
label102:
    End If
    Input #1, temp ' Z0
    n = 0
    prevphase = 0
    offset = 0

    Do While Not EOF(1)
         
         Input #1, Filedata0, Filedata1, Filedata2
            If (Filedata0 <> 0) Then
                n = n + 1
                If gmah = 1 Then Filedata0 = Filedata0 / 1000000#
                xatemp(n) = Filedata0   'freq
                y1atemp(n) = Filedata1 ' mag
                'y2atemp(n) = filedata2 ' phase
            
                ' code written to unwrap phase
                phase = Filedata2
                If (Abs(phase - prevphase) > 180) Then offset = offset - 360 * Sgn(phase - prevphase)
                prevphase = phase
                y2atemp(n) = phase + offset
            End If
            If (n > 1000) Then Exit Do
    Loop
   start1 = xatemp(1)
   stop1 = xatemp(n)

   MsgBox "Start Freq = " & start1 & " MHz" & CRLF & "Stop Freq = " & stop1 & " MHz" & CRLF & "Number of Points = " & n, , "Load / Frequency Sweep Settings"
   npoints = n
   step1 = Abs(stop1 - start1) / (npoints - 1)
   freq = start1
   Panel3D10.Caption = Format$(freq, "###0.000")

   Close #1
   If (CSng(Filedata0) <> 0) Then stop1 = CSng(Filedata0)
   Label1(3).Visible = True
   MULTITER = 1
   Call Zcalsweep(0)
   End If

End Sub

Private Sub NewItem_Click()
    Dim Value As Single
    VF = 1#
    Z0 = 50#
    freq = 100
    form2.Panel3D10.Caption = Format$(freq, "0.000")
    Value = 0
    ImpeItem.Checked = True
    GammaItem.Checked = False
    Label1(1).Caption = "R"
    Label1(2).Caption = "X"
    ZR(1) = 50  ' initialize load also
    ZI(1) = 0
    
    For I% = 0 To 14 Step 1
    stepS(I%) = 1
    stepP(I%) = 1
    s(I%) = 0
    p(I%) = 0
    SwpStp(I%) = 0
    Q(I%) = 1000000#
    Next I%
    
    I% = 2
    Do While I% <= 10
        Call SeriesIO("W", Value, I%)
        If (I% <= 9) Then Call ParallelIO("W", Value, I% + 1)
        I% = I% + 2
    Loop
    If (Option3D1.Value = False) Then Option3D1.Value = True
    form2.Caption = "Schematic"
    Zcalsweep (0)
End Sub

Private Sub OpenItem_Click()
Dim dat, dat1, dat2, dat3, dat4, points
Dim l_count As Integer
On Error Resume Next

CMDialog2.CancelError = -1
CMDialog2.DialogTitle = "Open Schematic"
CMDialog2.Filter = "Schematic Files | *.Sch"
CMDialog2.Flags = OFN_READONLY Or &H2000 ' Create if file doesnot exist
CMDialog2.Filename = "*.sch"
CMDialog2.Action = 1
If Err = 0 Then             ' Reset schematic
    For I% = 0 To 14 Step 1
    stepS(I%) = 1
    stepP(I%) = 1
    s(I%) = 0
    p(I%) = 0
    SwpStp(I%) = 0
    Q(I%) = 1000000#
    Next I%
    I% = 2
    Do While I% <= 10
        Call SeriesIO("W", Value, I%)
        If (I% <= 9) Then Call ParallelIO("W", Value, I% + 1)
        I% = I% + 2
    Loop
    If (Option3D1.Value = False) Then Option3D1.Value = True
    
    'Read file and assign
    Open CMDialog2.Filename For Input As #1
    Input #1, dat ' 1st line
    If (dat <> "QuickSmith SCH") Then
        Beep
        MsgBox "Bad Schematic File !" & CRLF & " Check QuickSmith Schematic File Format ", 48, "QuickSmith"
        Close #1
        GoTo endfileOpen
    End If
    Input #1, dat   '2nd line
    VF = dat
    If VF = 0 Then VF = 0.0000000001
    Input #1, dat   '3rd line
    Z0 = dat
    If (Z0 = 0) Then Z0 = 0.0000000001
    Input #1, dat   '4th Line
    freq = dat
    form2.Panel3D10.Caption = Format$(freq, "0.000")
    Input #1, dat   '5th Line
    TLfreq = dat
    
    Input #1, dat   '6th Line
    l_count = 1 'this was added break from infinite do loop
    Do Until TL = dat Or l_count = 10
    l_count = l_count + 1
    Spin6_SpinDown
    Loop
    TL = dat
    
    I% = 2   ' used for elements
   ' the code here was added to include load information and sweep details

    Input #1, dat   '7th Line
     If (CStr(dat) = "I" Or CStr(dat) = "G") Then
        Input #1, dat1 '  Load  resistance
        Input #1, dat2 '  Load reactance
        form2.Label1(3).Visible = False
        MULTITER = 0
        Call loadValue(dat, dat1, dat2)
        Input #1, dat1 '  Load  Sweep_Start
        start1 = dat1
        Input #1, dat1 '  Load  Sweep_Stop
        stop1 = dat1
        Input #1, dat1 ' Load Sweep_step
        step1 = dat1
        points = Abs((stop1 - start1) / step1) + 1
        If (points > 1000) Then step1 = (stop1 - start1) / 999


     Input #1, dat   'next line start with elements
     End If

    Do While Not EOF(1)
       
    'Input #1, dat   '7th Line
     Select Case CStr(dat)
     
     Case "T", "SLC", "PLC", "SRC", "PRC"
        Input #1, dat1 ' value
        Input #1, dat2 'Qvalue
        Input #1, dat3 'Zvalue
        Call SeriesIO(dat, dat1, I%) ' element,value,index
        If (dat2 > 0) Then Q(I%) = dat2 Else Q(I%) = 0.000000000001
        z(I%) = dat3
     Case "W"
     Case "R", "X"
        Input #1, dat1 ' value
        Call SeriesIO(dat, dat1, I%) ' element,value,index
     Case "C", "L"
        Input #1, dat1 ' value
        Input #1, dat2 'Qvalue
        Call SeriesIO(dat, dat1, I%) ' element,value,index
        If (dat2 > 0) Then Q(I%) = dat2 Else Q(I%) = 0.000000000001
     End Select
     
     If (Not EOF(1) And I% <= 9) Then
      Input #1, dat
      
      Select Case CStr(dat)
      Case "O", "S", "SLC", "PLC", "SRC", "PRC"
        Input #1, dat1 ' value
        Input #1, dat2 'Qvalue
        Input #1, dat3 'Zvalue
        Call ParallelIO(dat, dat1, I% + 1) ' element,value,index
        If (dat2 > 0) Then Q(I% + 1) = dat2 Else Q(I% + 1) = 0.000000000001
        z(I% + 1) = dat3
      Case "W"
      Case "R", "X"
         Input #1, dat1 ' value
         Call ParallelIO(dat, dat1, I% + 1) ' element,value,index
      Case "C", "L"
        Input #1, dat1 ' value
        Input #1, dat2 'Qvalue
        Call ParallelIO(dat, dat1, I% + 1) ' element,value,index
        If (dat2 > 0) Then Q(I% + 1) = dat2 Else Q(I% + 1) = 0.000000000001
      End Select
    End If
     
    I% = I% + 2
    If (I% > 10) Then Exit Do
    Input #1, dat   '7th Line
    Loop

Close #1
form2.Caption = CMDialog2.Filename   ' print file name in form2 caption
Zcalsweep (0)
endfileOpen:
End If

End Sub

Private Sub Option3D1_Click(Value As Integer)
For I% = 3 To 9 Step 2    ' parallel  ' inches to mm
      Select Case form2.Image12(I%).Tag
         Case "S", "O", "SLC", "PLC", "SRC", "PRC"
         'P(I%) = P(I%) * 25.4
         Panel3D6(I%).Caption = Format$(p(I%), "0.000")
         If (p(I%) > 99999.99 Or p(I%) < -99999.99) Then Panel3D6(I%).Caption = Format$(p(I%), "##.#E+##")
      End Select
    Next I%
   For I% = 2 To 10 Step 2            ' series
      Select Case form2.Image11(I%).Tag
         Case "T", "SLC", "PLC", "SRC", "PRC"
          'S(I%) = S(I%) * 25.4
         Panel3D3(I%).Caption = Format$(s(I%), "0.000")
         If (s(I%) > 99999.99 Or s(I%) < -99999.99) Then Panel3D3(I%).Caption = Format$(s(I%), "##.#E+##")
       End Select
   Next I%
        




End Sub

Private Sub Option3D2_Click(Value As Integer)
For I% = 3 To 9 Step 2    ' parallel  ' inches to mm
      Select Case form2.Image12(I%).Tag
         Case "S", "O", "SLC", "PLC", "SRC", "PRC"
         'P(I%) = P(I%) * 25.4
         Panel3D6(I%).Caption = Format$(z(I%), "0.000")
         If (z(I%) > 99999.99 Or z(I%) < -99999.99) Then Panel3D6(I%).Caption = Format$(z(I%), "##.#E+##")
      End Select
    Next I%
   For I% = 2 To 10 Step 2            ' series
      Select Case form2.Image11(I%).Tag
         Case "T", "SLC", "PLC", "SRC", "PRC"
          'S(I%) = S(I%) * 25.4
         Panel3D3(I%).Caption = Format$(z(I%), "0.000")
         If (z(I%) > 99999.99 Or z(I%) < -99999.99) Then Panel3D3(I%).Caption = Format$(z(I%), "##.#E+##")
       End Select
   Next I%
 'form4.Caption = "Tune Line Z0"
End Sub

Private Sub Panel3D10_DblClick()
stepsize$ = InputBox$("Enter Step Size ", "QuickSmith", CStr(StepFF))
If (IsNumeric(stepsize$)) Then StepFF = Abs(CDbl(stepsize$))
End Sub

Private Sub Panel3D3_DblClick(index As Integer)

stepsize$ = InputBox$("Enter Step Size ", "QuickSmith", CStr(stepS(index)))
If (IsNumeric(stepsize$)) Then stepS(index) = Abs(CDbl(stepsize$))

End Sub

Private Sub Panel3D3_DragDrop(index As Integer, Source As Control, X As Single, Y As Single)
Call Image11_DragDrop(index, Source, X, Y)
End Sub

Private Sub Panel3D4_DblClick()
stepsize$ = InputBox$("Enter Step Size", "QuickSmith", CStr(StepLR))
If (IsNumeric(stepsize$)) Then StepLR = Abs(CDbl(stepsize$))
End Sub

Private Sub Panel3D4_DragDrop(Source As Control, X As Single, Y As Single)
Call Image10_DragDrop(Source, X, Y)
End Sub

Private Sub Panel3D5_DblClick()
stepsize$ = InputBox$("Enter Step Size", "QuickSmith", CStr(StepLI))
If (IsNumeric(stepsize$)) Then StepLI = Abs(CDbl(stepsize$))
End Sub

Private Sub Panel3D5_DragDrop(Source As Control, X As Single, Y As Single)
Call Image10_DragDrop(Source, X, Y)
End Sub

Private Sub Panel3D6_DblClick(index As Integer)
stepsize$ = InputBox$("Enter Step Size", "QuickSmith", CStr(stepP(index)))
If (IsNumeric(stepsize$)) Then stepP(index) = Abs(CDbl(stepsize$))

End Sub

Private Sub Panel3D6_DragDrop(index As Integer, Source As Control, X As Single, Y As Single)
Call Image12_DragDrop(index, Source, X, Y)
End Sub

Private Sub PlotShowItem_Click()
Form3.Show
End Sub

Private Sub PrSchItem_Click()
Dim msg$, Filename$
Dim I%
Dim hscroll_flag As Integer
Dim Vscroll_flag As Integer

    hscroll_flag = False
    Vscroll_flag = False
    On Error GoTo FilePrintSetupError
    CMDialog2.CancelError = -1
    CMDialog2.Flags = &H40& 'PD_PRINTSETUP     ' Just enable print setup
    CMDialog2.Action = 5 ' DLG_PRINT        ' Show printer setup dialog box
   
    If Err = 0 Then
        Label3.Visible = True
        If (form2.Caption = "Schematic") Then Filename$ = "" Else Filename = form2.Caption
        Label3.Caption = "File Name: " & Filename$ & CRLF & "Char. Impedance: " & Z0 & CRLF & "Velocity Factor: " & VF & CRLF & "T-Line Design Frequency: " & TLfreq


        form2.BackColor = &HFFFFFF
        Panel3D4.BackColor = &HFFFFFF
        Panel3D5.BackColor = &HFFFFFF
        Panel3D4.BevelWidth = 0
        Panel3D5.BevelWidth = 0
    
        Command3D1.Visible = False
        Command3D2.Visible = False
        Command3D3.Visible = False
       ' panel3d8.Visible = False
       ' panel3d9.Visible = False
       ' Panel3D10.Visible = False
       ' panel3D12.Visible = False
        Panel3D2.Visible = False
        Spin6.Visible = False
        Option3D1.Visible = False
        Option3D2.Visible = False
        Spin1.Visible = False
        Spin3.Visible = False
        Spin5.Visible = False
        Image1.Visible = False
        Image2.Visible = False
        Image3.Visible = False
        Image4.Visible = False
        Image5.Visible = False
        Image6.Visible = False
        Image7.Visible = False
        Image8.Visible = False
        Image9.Visible = False
        Image13.Visible = False
        Image14.Visible = False
        'image15.Visible = False
        'image16.Visible = False
        For I% = 3 To 13 Step 1
            Label1(I%).Visible = False
        Next I%
        'label1(0).Visible = False
        For I% = 2 To 10 Step 2
            Spin2(I%).Visible = False
            Panel3D3(I%).BackColor = &HFFFFFF
            Panel3D3(I%).BevelWidth = 0
            If (I% <> 10) Then
                Spin4(I% + 1).Visible = False
                Panel3D6(I% + 1).BackColor = &HFFFFFF
                Panel3D6(I% + 1).BevelWidth = 0
            End If
        Next I%
    If (form2.HScroll1.Visible) Then
    hscroll_flag = True
    form2.HScroll1.Visible = False
    End If
    If (form2.VScroll1.Visible) Then
    Vscroll_flag = True
    form2.VScroll1.Visible = False
    End If

        Refresh
        'Printer.ScaleMode = 3
        'Printer.Print ""
    'modified code here to call new module print1.bas
    'instead of print.bas, to facilate high resolution printing

        Call PrintClient256(form2)
        
        Label3.Visible = False
        form2.BackColor = &HC0C0C0
        Panel3D4.BackColor = &HC0C0C0
        Panel3D5.BackColor = &HC0C0C0
        Panel3D4.BackColor = &HC0C0C0
        Panel3D4.BevelWidth = 1
        Panel3D5.BevelWidth = 1
     
         For I% = 3 To 13 Step 1
            Label1(I%).Visible = True
         Next I%
        'label1(0).Visible = True
        For I% = 2 To 10 Step 2
            Spin2(I%).Visible = True
            Panel3D3(I%).BackColor = &HC0C0C0
            Panel3D3(I%).BevelWidth = 1
            If (I% <> 10) Then
                Spin4(I% + 1).Visible = True
                Panel3D6(I% + 1).BackColor = &HC0C0C0
                Panel3D6(I% + 1).BevelWidth = 1
            End If
        Next I%
    
        Command3D1.Visible = True
        Command3D2.Visible = True
         Command3D3.Visible = True
        Panel3D2.Visible = True
        Spin6.Visible = True
        Option3D1.Visible = True
        Option3D2.Visible = True
        'panel3d8.Visible = True
        'panel3d9.Visible = True
        'Panel3D10.Visible = True
       ' panel3D12.Visible = True
        'panel3D13.Visible = True
        Spin1.Visible = True
        Spin3.Visible = True
        Spin5.Visible = True
        Image1.Visible = True
        Image2.Visible = True
        Image3.Visible = True
        Image4.Visible = True
        Image5.Visible = True
        Image6.Visible = True
        Image7.Visible = True
        Image8.Visible = True
        Image9.Visible = True
        Image13.Visible = True
        Image14.Visible = True
        'image15.Visible = True
        'image16.Visible = True
        If (hscroll_flag) Then form2.HScroll1.Visible = True
        If (Vscroll_flag) Then form2.VScroll1.Visible = True
        form2.Visible = True
        If (MULTITER) Then Label1(3).Visible = True Else Label1(3).Visible = False
    End If


PrintSetupExit:
    Exit Sub

FilePrintSetupError:
    If Err <> CDERR_CANCEL Then         ' If user didn't select CANCEL
    End If
    Resume PrintSetupExit

End Sub

Private Sub RedrawMenu_click()
If (form2.WindowState = 0) Then
h! = form2.Height
form2.Height = 0
form2.Height = h!
  If form2.HScroll1.Visible Then
    If (form2.VScroll1.Visible) Then
        form2.HScroll1.Move 0, form2.ScaleHeight - form2.HScroll1.Height, form2.ScaleWidth - form2.VScroll1.Width
    Else
        form2.HScroll1.Move 0, form2.ScaleHeight - form2.HScroll1.Height, form2.ScaleWidth
    End If
  End If
End If
End Sub

Private Sub SaveAsMenu_Click()
On Error Resume Next
CMDialog2.CancelError = -1
CMDialog2.DialogTitle = "Save File"
CMDialog2.Filter = "Schematic Files | *.sch"
CMDialog2.Flags = &H4    ' hide read only box
CMDialog2.Filename = "*.sch"
CMDialog2.Action = 2

If Err = 0 Then
    If (Dir(CMDialog2.Filename) <> "") Then 'file already exist
        response = MsgBox("Overwrite Existing File ?", 4 + 32 + 256, "QuickSmith")
        If response = 7 Then Exit Sub
    End If
   
   Open CMDialog2.Filename For Output As #1
   Print #1, "QuickSmith SCH"
   Print #1, VF
   Print #1, Z0
   Print #1, freq
   Print #1, TLfreq
   Print #1, TL
   If (form2.ImpeItem.Checked = True) Then
     Print #1, "I"
     Print #1, ZR(1)
     Print #1, ZI(1)
   Else
     Print #1, "G"
     Print #1, gammaM
     Print #1, GammaA
   End If
   Print #1, start1
   Print #1, stop1
   Print #1, step1

   I% = 2
   Do While I% <= 10
   Select Case Image11(I%).Tag
     Case "T"
       Print #1, Image11(I%).Tag & CRLF & Panel3D3(I%).Caption & CRLF & Format(Q(I%), "#00.00##") & CRLF & Format(z(I%), "##0.00##")
     Case "W"
       Print #1, Image11(I%).Tag
     Case "R", "X"
       Print #1, Image11(I%).Tag & CRLF & Panel3D3(I%).Caption
     Case "C", "L"
       Print #1, Image11(I%).Tag & CRLF & Panel3D3(I%).Caption & CRLF & Format(Q(I%), "#00.00##")
     Case "SLC", "PLC", "SRC", "PRC"
       Print #1, Image11(I%).Tag & CRLF & Panel3D3(I%).Caption & CRLF & Format(Q(I%), "#00.00##") & CRLF & Format(z(I%), "##0.00##")
     End Select
   If I% = 11 Then GoTo Skipthis
    Select Case Image12(I% + 1).Tag
      Case "S", "O"
        Print #1, Image12(I% + 1).Tag & CRLF & Panel3D6(I% + 1).Caption & CRLF & Format(Q(I% + 1), "#00.00##") & CRLF & Format(z(I% + 1), "##0.00##")
      Case "W"
        Print #1, Image12(I% + 1).Tag
      Case "R", "X"
         Print #1, Image12(I% + 1).Tag & CRLF & Panel3D6(I% + 1).Caption
      Case "C", "L"
         Print #1, Image12(I% + 1).Tag & CRLF & Panel3D6(I% + 1).Caption & CRLF & Format(Q(I% + 1), "#00.00##")
      Case "SLC", "PLC", "SRC", "PRC"
          Print #1, Image12(I% + 1).Tag & CRLF & Panel3D6(I% + 1).Caption & CRLF & Format(Q(I% + 1), "#00.00##") & CRLF & Format(z(I% + 1), "##0.00##")
   End Select
Skipthis:
   I% = I% + 2
   Loop
   
   Close #1
  form2.Caption = CMDialog2.Filename   ' print file name on caption

End If

End Sub

Private Sub SchematicItem_Click()
form2.Show

End Sub

Private Sub SingItem_Click()
Label1(3).Visible = False
MULTITER = 0
End Sub

Private Sub Spin1_SpinDown()
If (ImpeItem.Checked) Then
 If (Panel3D4.Caption <> "") Then
  If ((ZR(1) - StepLR) > 0) Then ZR(1) = ZR(1) - StepLR Else ZR(1) = 1E-25
  Panel3D4.Caption = Format$(ZR(1), "0.000")
  If (ZR(1) > 99999.99) Then Panel3D4.Caption = Format$(ZR(1), "##.#E+##")
 End If
Else
 If (Panel3D4.Caption <> "") Then
  gammaM = gammaM - StepLR
  Panel3D4.Caption = Format$(gammaM, "0.000")
  If (gammaM > 99999.99 Or gammaM < -99999.99) Then Panel3D4.Caption = Format$(gammaM, "##.#E+##")
 End If
End If
Call Zcalsweep(0)
 End Sub

Private Sub Spin1_SpinUp()
If (ImpeItem.Checked) Then
 If (Panel3D4.Caption <> "") Then
  ZR(1) = ZR(1) + StepLR
  Panel3D4.Caption = Format$(ZR(1), "0.000")
  If (ZR(1) > 99999.99 Or ZR(1) < -99999.99) Then Panel3D4.Caption = Format$(ZR(1), "##.#E+##")
 End If
Else
 If (Panel3D4.Caption <> "") Then
  gammaM = gammaM + StepLR
  Panel3D4.Caption = Format$(gammaM, "0.000")
  If (gammaM > 99999.99 Or gammaM < -99999.99) Then Panel3D4.Caption = Format$(gammaM, "##.#E+##")
 End If
End If
Call Zcalsweep(0)

End Sub

Private Sub Spin2_SpinDown(index As Integer)
    
     Select Case Image11(index).Tag
      Case "C", "L", "R", "X"
           s(index) = s(index) - stepS(index)
            Panel3D3(index).Caption = Format(s(index), "0.000")
            If (s(index) > 99999.99 Or s(index) < -99999.99) Then Panel3D3(index).Caption = Format$(s(index), "##.##E+##")
      
      Case "T", "SLC", "PLC", "SRC", "PRC"
            If (Option3D2.Value = True) Then
                z(index) = z(index) - stepS(index)
                If (z(index) <= 0) Then z(index) = 0.000000000001
                Panel3D3(index).Caption = Format(z(index), "0.000")
                If (z(index) > 99999.99 Or z(index) < -99999.99) Then Panel3D3(index).Caption = Format$(z(index), "##.##E+##")
            Else
                s(index) = s(index) - stepS(index)
                Panel3D3(index).Caption = Format(s(index), "0.000")
                If (s(index) > 99999.99 Or s(index) < -99999.99) Then Panel3D3(index).Caption = Format$(s(index), "##.##E+##")
            End If
      End Select

Call Zcalsweep(0)
End Sub

Private Sub Spin2_SpinUp(index As Integer)
     Select Case Image11(index).Tag
      Case "C", "L", "R", "X"
            s(index) = s(index) + stepS(index)
            Panel3D3(index).Caption = Format(s(index), "0.000")
            If (s(index) > 99999.99 Or s(index) < -99999.99) Then Panel3D3(index).Caption = Format$(s(index), "##.##E+##")
      
      Case "T", "SLC", "PLC", "SRC", "PRC"
            If (Option3D2.Value = True) Then
                z(index) = z(index) + stepS(index)
                If (z(index) <= 0) Then z(index) = 0.000000000001
                Panel3D3(index).Caption = Format(z(index), "0.000")
                If (z(index) > 99999.99 Or z(index) < -99999.99) Then Panel3D3(index).Caption = Format$(z(index), "##.##E+##")
            Else
                s(index) = s(index) + stepS(index)
                Panel3D3(index).Caption = Format(s(index), "0.000")
                If (s(index) > 99999.99 Or s(index) < -99999.99) Then Panel3D3(index).Caption = Format$(s(index), "##.##E+##")
            End If
      End Select

Call Zcalsweep(0)



End Sub

Private Sub Spin3_SpinDown()
If (ImpeItem.Checked) Then
 If (Panel3D5.Caption <> "") Then
  ZI(1) = ZI(1) - StepLI
  Panel3D5.Caption = Format$(ZI(1), "0.000")
  If (ZI(1) > 99999.99 Or ZI(1) < -99999.99) Then Panel3D4.Caption = Format$(ZI(1), "##.#E+##")
 End If
Else
 If (Panel3D5.Caption <> "") Then
  GammaA = GammaA - StepLI
  Panel3D5.Caption = Format$(GammaA, "0.000")
  If (GammaA > 99999.99 Or GammaA < -99999.99) Then Panel3D4.Caption = Format$(GammaA, "##.#E+##")
 End If


End If
Call Zcalsweep(0)
End Sub

Private Sub Spin3_SpinUp()
If (ImpeItem.Checked) Then
 If (Panel3D5.Caption <> "") Then
  ZI(1) = ZI(1) + StepLI
  Panel3D5.Caption = Format$(ZI(1), "0.000")
  If (ZI(1) > 99999.99 Or ZI(1) < -99999.99) Then Panel3D4.Caption = Format$(ZI(1), "##.#E+##")
 End If
Else
 If (Panel3D5.Caption <> "") Then
  GammaA = GammaA + StepLI
  Panel3D5.Caption = Format$(GammaA, "0.000")
  If (GammaA > 99999.99 Or GammaA < -99999.99) Then Panel3D4.Caption = Format$(GammaA, "##.#E+##")
 End If


End If
Call Zcalsweep(0)

End Sub

Private Sub Spin4_SpinDown(index As Integer)
   
 Select Case Image12(index).Tag
     Case "C", "L", "R", "X"
           p(index) = p(index) - stepP(index)
           Panel3D6(index).Caption = Format(p(index), "0.000")
           If (p(index) > 99999.99 Or p(index) < -99999.99) Then Panel3D6(index).Caption = Format$(p(index), "##.#E+##")
      Case "O", "S", "SLC", "PLC", "SRC", "PRC"
           If (Option3D2.Value = True) Then
                z(index) = z(index) - stepP(index)
                If (z(index) <= 0) Then z(index) = 0.000000000001
                Panel3D6(index).Caption = Format(z(index), "0.000")
                If (z(index) > 99999.99 Or z(index) < -99999.99) Then Panel3D6(index).Caption = Format$(z(index), "##.#E+##")
           Else
                p(index) = p(index) - stepP(index)
                Panel3D6(index).Caption = Format(p(index), "0.000")
                If (p(index) > 99999.99 Or p(index) < -99999.99) Then Panel3D6(index).Caption = Format$(p(index), "##.#E+##")
          End If
    End Select
Call Zcalsweep(0)
End Sub

Private Sub Spin4_SpinUp(index As Integer)
 Select Case Image12(index).Tag
     Case "C", "L", "R", "X"
           p(index) = p(index) + stepP(index)
           Panel3D6(index).Caption = Format(p(index), "0.000")
           If (p(index) > 99999.99 Or p(index) < -99999.99) Then Panel3D6(index).Caption = Format$(p(index), "##.#E+##")
      Case "O", "S", "SLC", "PLC", "SRC", "PRC"
           If (Option3D2.Value = True) Then
                z(index) = z(index) + stepP(index)
                If (z(index) <= 0) Then z(index) = 0.000000000001
                Panel3D6(index).Caption = Format(z(index), "0.000")
                If (z(index) > 99999.99 Or z(index) < -99999.99) Then Panel3D6(index).Caption = Format$(z(index), "##.#E+##")
           Else
                p(index) = p(index) + stepP(index)
                Panel3D6(index).Caption = Format(p(index), "0.000")
                If (p(index) > 99999.99 Or p(index) < -99999.99) Then Panel3D6(index).Caption = Format$(p(index), "##.#E+##")
          End If
    End Select
Call Zcalsweep(0)
        
End Sub

Private Sub Spin5_SpinDown()
If (Panel3D10.Caption <> "") Then
  If ((freq - StepFF) > 0) Then freq = freq - StepFF Else freq = 0.000000000001
  Panel3D10.Caption = Format$(freq, "###0.000")
  If (freq > 9999.99) Then Panel3D10.Caption = Format$(freq, "##.#E+##")
End If

Call Zcalsweep(0)
  
End Sub

Private Sub Spin5_SpinUp()
If (Panel3D10.Caption <> "") Then
  If (freq > 0) Then freq = freq + StepFF Else freq = 0.000000000001
  Panel3D10.Caption = Format$(freq, "###0.000")
  If (freq > 9999.99) Then Panel3D10.Caption = Format$(freq, "##.#E+##")
End If

Call Zcalsweep(0)

End Sub

Private Sub Spin6_SpinDown()
  ' 0=Inch => 1=mm => 2=degree => 3=lamda => 4 =>meters
 If (Option3D2 = False) Then
 Select Case TL
  Case 0    ' old   inches
        Panel3D2.Caption = "mm"
        Label1(9).Caption = "mm"
        Label1(10).Caption = "mm"
        Label1(11).Caption = "mm"
        TL = 1 ' new  mm
    For I% = 3 To 9 Step 2    ' parallel  ' inches to mm
      Select Case form2.Image12(I%).Tag
         Case "S", "O"
         p(I%) = p(I%) * 25.4
         Panel3D6(I%).Caption = Format$(p(I%), "0.000")
         If (p(I%) > 99999.99 Or p(I%) < -99999.99) Then Panel3D6(I%).Caption = Format$(p(I%), "##.#E+##")
      End Select
    Next I%
   For I% = 2 To 10 Step 2            ' series
      Select Case form2.Image11(I%).Tag
         Case "T"
          s(I%) = s(I%) * 25.4
         Panel3D3(I%).Caption = Format$(s(I%), "0.000")
         If (s(I%) > 99999.99 Or s(I%) < -99999.99) Then Panel3D3(I%).Caption = Format$(s(I%), "##.#E+##")
       End Select
   Next I%
        

  Case 2 'old  degree
        Panel3D2.Caption = "WaveLen."
        Label1(9).Caption = "lam."
        Label1(10).Caption = "lam."
        Label1(11).Caption = "lam."
        TL = 3  ' lamda
    For I% = 3 To 9 Step 2    ' parallel  ' degree to lamda
      Select Case form2.Image12(I%).Tag
         Case "S", "O"
         p(I%) = p(I%) / 360
         Panel3D6(I%).Caption = Format$(p(I%), "0.000")
         If (p(I%) > 99999.99 Or p(I%) < -99999.99) Then Panel3D6(I%).Caption = Format$(p(I%), "##.#E+##")
      End Select
    Next I%
   For I% = 2 To 10 Step 2            ' series
      Select Case form2.Image11(I%).Tag
         Case "T"
         s(I%) = s(I%) / 360
         Panel3D3(I%).Caption = Format$(s(I%), "0.000")
         If (s(I%) > 99999.99 Or s(I%) < -99999.99) Then Panel3D3(I%).Caption = Format$(s(I%), "##.#E+##")
       End Select
   Next I%
  
  
  Case 1    ' from mm
        Panel3D2.Caption = "Degree"
        Label1(9).Caption = "Deg."
        Label1(10).Caption = "Deg."
        Label1(11).Caption = "Deg."
        TL = 2   ' to degree
    For I% = 3 To 9 Step 2    ' parallel  ' mm to degree
      Select Case form2.Image12(I%).Tag   '
         Case "S", "O"
         p(I%) = p(I%) * TLfreq * 1000000 * 360 / (VF * C * 1000)
         Panel3D6(I%).Caption = Format$(p(I%), "0.000")
         If (p(I%) > 99999.99 Or p(I%) < -99999.99) Then Panel3D6(I%).Caption = Format$(p(I%), "##.#E+##")
      End Select
    Next I%
   For I% = 2 To 10 Step 2            ' series
      Select Case form2.Image11(I%).Tag
         Case "T"
          s(I%) = s(I%) * TLfreq * 360 * 1000000 / (VF * C * 1000)
         Panel3D3(I%).Caption = Format$(s(I%), "0.000")
         If (s(I%) > 99999.99 Or s(I%) < -99999.99) Then Panel3D3(I%).Caption = Format$(s(I%), "##.#E+##")
       End Select
   Next I%
  Case 3    ' from lamda
        Panel3D2.Caption = "Meters"
        Label1(9).Caption = "m"
        Label1(10).Caption = "m"
        Label1(11).Caption = "m"
        TL = 4   ' to meters
    For I% = 3 To 9 Step 2    ' parallel  ' degree to meters
      Select Case form2.Image12(I%).Tag   '
         Case "S", "O"
         p(I%) = p(I%) * VF * C / (TLfreq * 1000000)
         Panel3D6(I%).Caption = Format$(p(I%), "0.000")
         If (p(I%) > 99999.99 Or p(I%) < -99999.99) Then Panel3D6(I%).Caption = Format$(p(I%), "##.#E+##")
      End Select
    Next I%
   For I% = 2 To 10 Step 2            ' series
      Select Case form2.Image11(I%).Tag
         Case "T"
         s(I%) = s(I%) * VF * C / (TLfreq * 1000000)
         Panel3D3(I%).Caption = Format$(s(I%), "0.000")
         If (s(I%) > 99999.99 Or s(I%) < -99999.99) Then Panel3D3(I%).Caption = Format$(s(I%), "##.#E+##")
       End Select
   Next I%
  Case 4    ' from Meters
        Panel3D2.Caption = "Inches"
        Label1(9).Caption = "inch"
        Label1(10).Caption = "inch"
        Label1(11).Caption = "inch"
        TL = 0   ' to inches
    For I% = 3 To 9 Step 2    ' parallel  ' degree to meters
      Select Case form2.Image12(I%).Tag   '
         Case "S", "O"
         p(I%) = p(I%) / 0.0254
         Panel3D6(I%).Caption = Format$(p(I%), "0.000")
         If (p(I%) > 99999.99 Or p(I%) < -99999.99) Then Panel3D6(I%).Caption = Format$(p(I%), "##.#E+##")
      End Select
    Next I%
   For I% = 2 To 10 Step 2            ' series
      Select Case form2.Image11(I%).Tag
         Case "T"
         s(I%) = s(I%) / 0.0254
         Panel3D3(I%).Caption = Format$(s(I%), "0.000")
         If (s(I%) > 99999.99 Or s(I%) < -99999.99) Then Panel3D3(I%).Caption = Format$(s(I%), "##.#E+##")
       End Select
   Next I%



End Select
End If
End Sub

Private Sub Spin6_SpinUp()
        
Spin6_SpinDown


End Sub

Private Sub SpotItem_Click()
Call Ilcal(0)
End Sub

Private Sub StartFreqMenu_Click()
Form8.Show
'R$ = InputBox$("Enter Start Value ", "QuickSmith", CStr(start1))
'If (IsNumeric(R$) = False) Then GoTo startfreqEnd
'If (CDbl(R$) = 0) Then start1 = .000000000001 Else start1 = CDbl(R$)
'startfreqEnd:

End Sub

Private Sub StepItem_Click()
Dim points As Long
R$ = InputBox$("Enter Step Value", "QuickSmith", CStr(step1))
If (IsNumeric(R$) = False) Then GoTo StepEnd

If (CDbl(R$) <= 0) Then GoTo StepEnd ' zero condition
step1 = CDbl(R$)
points = Abs((stop1 - start1) / step1) + 1
If (points > 1000) Then
step1 = (stop1 - start1) / 999
MsgBox ("Number of Points exceeds 1000. Step Size Set to " & step1)
End If
StepEnd:

End Sub

Private Sub StopFreqMenu_Click()
R$ = InputBox$("Enter Stop Value", "QuickSmith", CStr(stop1))
If (IsNumeric(R$) = False) Then GoTo stopfreqEnd
If (CDbl(R$) = 0) Then stop1 = 0.000000000001 Else stop1 = CDbl(R$)
stopfreqEnd:

End Sub

Private Sub sweepItem_click()
On Error GoTo SweepSchErr
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

SweepSchErr:
If (Err = 11) Then
    response% = MsgBox("' A Division By Zero' Error Was Trapped", 49, "Error")
    If response = 1 Then Resume Else Exit Sub
End If


End Sub

Private Sub SweepMenu_Click()
Form8.Show
End Sub

Private Sub SweepStepItem_Click()
'SW:        Title$ = "QuickSmith"
 '           Prompt$ = "Enter  Sweep Step Size in MHz"
   '         Default$ = ""
  '          inp$ = InputBox$(Prompt$, Title$, Default$)
            'If (inp$ = "") Then inp$ = "0"
    '        If (IsNumeric(inp$)) Then
     '       If (CDbl(inp$) > 0) Then SS = CDbl(inp$) Else GoTo SW
            'Panel3d10.Caption = Format$(CDbl(inp$), "0.000")
            'If (CDbl(inp$) > 99999.99 Or CDbl(inp$) < -99999.99) Then Panel3d10.Caption = Format$(CDbl(inp$), "##.#E+##")
      '      End If

End Sub

Private Sub TLineFreqItem_Click()
R$ = InputBox$("Enter T-Line Design Frequency in MHz", "QuickSmith", CStr(TLfreq))
If (IsNumeric(R$) = False) Then GoTo EndTLfreq
If (CDbl(R$) < 0) Then GoTo EndTLfreq
If (CDbl(R$) > 0) Then TLfreq = CDbl(R$) Else freq = 0.000000000001
Call Zcalsweep(0)

EndTLfreq:


End Sub

Private Sub VFItem_Click()
R$ = InputBox$("Enter Velocity Factor of The Line", "QuickSmith", CStr(VF))
If (IsNumeric(R$) = False) Then GoTo EndVF
If (CDbl(R$) < 0) Then GoTo EndVF
VF = CDbl(R$)
If (VF = 0) Then VF = 0.0000000001
Zcalsweep (0)
EndVF:

End Sub

Private Sub VScroll1_Change()
Picture2.Top = -VScroll1.Value
End Sub

Private Sub WinItem_Click(index As Integer)
    On Error GoTo form2WindowError
    Select Case index
      Case 0:
             form1.ShowToolItem.Checked = True
             WindowVertical = False
             WindowHorizontal = False
             form4.Show 'makes toolbox visible
             mdiArrange WM_MDICASCADE
      Case 1:
            form4.Hide 'makes toolbox visible
            form1.ShowToolItem.Checked = False
            WindowVertical = False
            WindowHorizontal = True
            mdiArrange MDITILE_HORIZONTAL
      Case 2:
             form4.Hide 'makes toolbox visible
             form1.ShowToolItem.Checked = False
             WindowVertical = True
             WindowHorizontal = False
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
            
            form2.Move 400, 0, 0.84 * MDIForm1.Width, 0.75 * MDIForm1.Height
            form1.Move 0, 0, 0.84 * MDIForm1.Width, 0.85 * MDIForm1.Height
            form1.ShowToolItem.Checked = True
            form4.Show 'makes toolbox visible
            
    End Select
    form2.Picture1.Move 0, 0, ScaleWidth, ScaleHeight 'HScroll1.Top
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
    form1.Frame3D1.Left = -form1.HScroll1.Value
    form1.Frame3D1.Width = form1.Width + form1.HScroll1.Value
    form2.Picture2.Left = -form2.HScroll1.Value
Else
    form2.HScroll1.Visible = False
    form1.HScroll1.Visible = False
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
Call Form_Resize  ' check if scroll bars are needed
Exit Sub
form2WindowError:
If (Err) Then
response% = MsgBox("QuickSmith Detected An Error", 49, "Error")
If response = 1 Then Resume Else Exit Sub
End If


End Sub

