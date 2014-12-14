VERSION 5.00
Object = "{B16553C3-06DB-101B-85B2-0000C009BE81}#1.0#0"; "SPIN32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form form4 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   " Component Tuner"
   ClientHeight    =   3660
   ClientLeft      =   4890
   ClientTop       =   2490
   ClientWidth     =   1620
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3660
   ScaleWidth      =   1620
   Begin Spin.SpinButton Spin4 
      Height          =   265
      Index           =   3
      Left            =   1332
      TabIndex        =   20
      Top             =   852
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
   Begin Threed.SSPanel Panel3D6 
      Height          =   252
      Index           =   3
      Left            =   360
      TabIndex        =   5
      Top             =   852
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
   End
   Begin Spin.SpinButton Spin5 
      Height          =   265
      Index           =   0
      Left            =   1344
      TabIndex        =   21
      Top             =   120
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
   Begin Spin.SpinButton Spin4 
      Height          =   265
      Index           =   9
      Left            =   1332
      TabIndex        =   22
      Top             =   2988
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
   Begin Spin.SpinButton Spin4 
      Height          =   265
      Index           =   7
      Left            =   1332
      TabIndex        =   23
      Top             =   2268
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
   Begin Spin.SpinButton Spin4 
      Height          =   265
      Index           =   5
      Left            =   1332
      TabIndex        =   24
      Top             =   1548
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
   Begin Threed.SSPanel Panel3D6 
      Height          =   252
      Index           =   9
      Left            =   360
      TabIndex        =   8
      Top             =   2988
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
   End
   Begin Threed.SSPanel Panel3D6 
      Height          =   252
      Index           =   7
      Left            =   360
      TabIndex        =   7
      Top             =   2268
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
   End
   Begin Threed.SSPanel Panel3D6 
      Height          =   252
      Index           =   5
      Left            =   360
      TabIndex        =   6
      Top             =   1548
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
   End
   Begin Spin.SpinButton Spin2 
      Height          =   265
      Index           =   10
      Left            =   1332
      TabIndex        =   25
      Top             =   3348
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
   Begin Threed.SSPanel Panel3D3 
      Height          =   252
      Index           =   10
      Left            =   360
      TabIndex        =   4
      Top             =   3348
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
   End
   Begin Spin.SpinButton Spin2 
      Height          =   265
      Index           =   8
      Left            =   1332
      TabIndex        =   26
      Top             =   2628
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
   Begin Threed.SSPanel Panel3D3 
      Height          =   252
      Index           =   8
      Left            =   360
      TabIndex        =   3
      Top             =   2628
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
   End
   Begin Spin.SpinButton Spin2 
      Height          =   265
      Index           =   4
      Left            =   1332
      TabIndex        =   27
      Top             =   1188
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
   Begin Spin.SpinButton Spin2 
      Height          =   276
      Index           =   6
      Left            =   1332
      TabIndex        =   28
      Top             =   1908
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
   Begin Threed.SSPanel Panel3D3 
      Height          =   252
      Index           =   4
      Left            =   360
      TabIndex        =   1
      Top             =   1188
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
   End
   Begin Spin.SpinButton Spin2 
      Height          =   265
      Index           =   2
      Left            =   1332
      TabIndex        =   29
      Top             =   504
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
   Begin Threed.SSPanel Panel3D3 
      Height          =   252
      Index           =   2
      Left            =   360
      TabIndex        =   0
      Top             =   504
      Width           =   972
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   15
      ForeColor       =   4210752
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   25.78
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
   End
   Begin Threed.SSPanel Panel3D3 
      Height          =   252
      Index           =   6
      Left            =   360
      TabIndex        =   2
      Top             =   1908
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
   End
   Begin Threed.SSPanel Panel3D10 
      Height          =   252
      Left            =   504
      TabIndex        =   18
      Top             =   120
      Width           =   852
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      Caption         =   "Freq."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   192
      Index           =   0
      Left            =   24
      TabIndex        =   19
      Top             =   120
      Width           =   468
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   10
      Left            =   48
      TabIndex        =   9
      Top             =   3360
      Width           =   300
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   264
      Index           =   9
      Left            =   96
      TabIndex        =   11
      Top             =   3000
      Width           =   252
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   252
      Index           =   3
      Left            =   60
      TabIndex        =   17
      Top             =   840
      Width           =   252
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   312
      Index           =   5
      Left            =   84
      TabIndex        =   16
      Top             =   1560
      Width           =   252
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   252
      Index           =   6
      Left            =   96
      TabIndex        =   15
      Top             =   1944
      Width           =   252
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   192
      Index           =   7
      Left            =   96
      TabIndex        =   14
      Top             =   2280
      Width           =   252
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   192
      Index           =   8
      Left            =   84
      TabIndex        =   13
      Top             =   2652
      Width           =   252
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   252
      Index           =   4
      Left            =   72
      TabIndex        =   12
      Top             =   1200
      Width           =   252
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   264
      Index           =   2
      Left            =   84
      TabIndex        =   10
      Top             =   480
      Width           =   240
   End
End
Attribute VB_Name = "form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Panel3D10_DblClick()
stepsize$ = InputBox$("Enter Step Size ", "Smith Chart", CStr(StepFF))
If (IsNumeric(stepsize$)) Then StepFF = Abs(CDbl(stepsize$))

End Sub

Private Sub Panel3D3_DblClick(index As Integer)

stepsize$ = InputBox$("Enter Step Size ", "Smith Chart", CStr(stepS(index)))
If (IsNumeric(stepsize$)) Then stepS(index) = Abs(CDbl(stepsize$))


End Sub

Private Sub Panel3D6_DblClick(index As Integer)
stepsize$ = InputBox$("Enter Step Size", "Smith Chart", CStr(stepP(index)))
If (IsNumeric(stepsize$)) Then stepP(index) = Abs(CDbl(stepsize$))

End Sub

Private Sub Spin2_SpinDown(index As Integer)
     Select Case form2.Image11(index).Tag
      Case "C", "L", "R", "X"
           s(index) = s(index) - stepS(index)
            Panel3D3(index).Caption = Format(s(index), "0.000")
            If (s(index) > 99999.99 Or s(index) < -99999.99) Then Panel3D3(index).Caption = Format$(s(index), "##.##E+##")
      
      Case "T", "SLC", "PLC"
            If (form2.Option3D2.Value = True) Then
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
     Select Case form2.Image11(index).Tag
      Case "C", "L", "R", "X"
            s(index) = s(index) + stepS(index)
            Panel3D3(index).Caption = Format(s(index), "0.000")
            If (s(index) > 99999.99 Or s(index) < -99999.99) Then Panel3D3(index).Caption = Format$(s(index), "##.##E+##")
      
      Case "T", "SLC", "PLC"
            If (form2.Option3D2.Value = True) Then
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

Private Sub Spin3_SpinDown(index As Integer)
If (Panel3D10.Caption <> "") Then
  If ((freq - StepFF) > 0) Then freq = freq - StepFF Else freq = 0.000000000001
  Panel3D10.Caption = Format$(freq, "###0.000")
  If (freq > 9999.99) Then Panel3D10.Caption = Format$(freq, "##.#E+##")
End If

Call Zcalsweep(0)


End Sub

Private Sub Spin3_SpinUp(index As Integer)
If (Panel3D10.Caption <> "") Then
  If (freq > 0) Then freq = freq + StepFF Else freq = 0.000000000001
  Panel3D10.Caption = Format$(freq, "###0.000")
  If (freq > 9999.99) Then Panel3D10.Caption = Format$(freq, "##.#E+##")
End If

Call Zcalsweep(0)

End Sub

Private Sub Spin4_SpinDown(index As Integer)
   
 Select Case form2.Image12(index).Tag
     Case "C", "L", "R", "X"
           p(index) = p(index) - stepP(index)
           Panel3D6(index).Caption = Format(p(index), "0.000")
           If (p(index) > 99999.99 Or p(index) < -99999.99) Then Panel3D6(index).Caption = Format$(p(index), "##.#E+##")
      Case "O", "S", "SLC", "PLC"
           If (form2.Option3D2.Value = True) Then
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
 Select Case form2.Image12(index).Tag
     Case "C", "L", "R", "X"
           p(index) = p(index) + stepP(index)
           Panel3D6(index).Caption = Format(p(index), "0.000")
           If (p(index) > 99999.99 Or p(index) < -99999.99) Then form2.Panel3D6(index).Caption = Format$(p(index), "##.#E+##")
      Case "O", "S", "SLC", "PLC"
           If (form2.Option3D2.Value = True) Then
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

Private Sub Spin5_SpinDown(index As Integer)
If (Panel3D10.Caption <> "") Then
  If ((freq - StepFF) > 0) Then freq = freq - StepFF Else freq = 0.000000000001
  Panel3D10.Caption = Format$(freq, "###0.000")
  If (freq > 9999.99) Then Panel3D10.Caption = Format$(freq, "##.#E+##")
End If

Call Zcalsweep(0)

End Sub

Private Sub Spin5_SpinUp(index As Integer)
If (Panel3D10.Caption <> "") Then
  If ((freq + StepFF) > 0) Then freq = freq + StepFF Else freq = 0.000000000001
  Panel3D10.Caption = Format$(freq, "###0.000")
  If (freq > 9999.99) Then Panel3D10.Caption = Format$(freq, "##.#E+##")
End If

Call Zcalsweep(0)

End Sub

