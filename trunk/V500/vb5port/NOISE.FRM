VERSION 5.00
Begin VB.Form Form10 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Enter Noise Parameters"
   ClientHeight    =   2865
   ClientLeft      =   7710
   ClientTop       =   3285
   ClientWidth     =   4545
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2865
   ScaleWidth      =   4545
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "O K"
      Height          =   372
      Left            =   2748
      TabIndex        =   8
      Top             =   924
      Width           =   1212
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      Height          =   288
      Left            =   1200
      TabIndex        =   7
      Text            =   "9.42"
      Top             =   1800
      Width           =   972
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   288
      Left            =   1212
      TabIndex        =   6
      Text            =   "141"
      Top             =   1320
      Width           =   972
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   288
      Left            =   1200
      TabIndex        =   5
      Text            =   "0.542"
      Top             =   840
      Width           =   972
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   288
      Left            =   1188
      TabIndex        =   4
      Text            =   "2.9"
      Top             =   360
      Width           =   972
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "RN (Ohms) "
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   144
      TabIndex        =   3
      Top             =   1800
      Width           =   1068
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "G0 (Deg.)"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   180
      TabIndex        =   2
      Top             =   1332
      Width           =   1032
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "G0 (Mag.)"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   180
      TabIndex        =   1
      Top             =   864
      Width           =   1032
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Fmin [dB]"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   156
      TabIndex        =   0
      Top             =   396
      Width           =   1044
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
If (IsNumeric(Text1.Text) And IsNumeric(text2.Text) And IsNumeric(text3.Text) And IsNumeric(text4.Text)) Then
Fmin = Form10.Text1.Text
Form6.NoiseShowItem.Enabled = True
Form10.Hide
End If

End Sub

Private Sub Form_Load()
 Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
 NOISEOPENFLAG = True
 Form10.Show
 Text1.SetFocus
 
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

