VERSION 5.00
Begin VB.Form AboutBox 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About QuickSmith"
   ClientHeight    =   2865
   ClientLeft      =   1650
   ClientTop       =   2190
   ClientWidth     =   5865
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "System"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00C0C0C0&
   LinkMode        =   1  'Source
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2865
   ScaleWidth      =   5865
   Begin VB.CommandButton Cmd_OK 
      Appearance      =   0  'Flat
      Caption         =   "OK"
      Height          =   360
      Left            =   4560
      TabIndex        =   5
      Top             =   2040
      Width           =   1035
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   " Note: Smith Chart is a registered trademark of Analog Instruments Company"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   12
      TabIndex        =   7
      Top             =   2616
      Width           =   5856
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Questions/comments to niyer@sbcglobal.net"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   1296
      TabIndex        =   6
      Top             =   708
      Width           =   4272
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   552
      Left            =   684
      Picture         =   "ABOUTBOX.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   588
   End
   Begin VB.Line lin_HorizontalLine1 
      BorderWidth     =   2
      X1              =   516
      X2              =   4551
      Y1              =   1668
      Y2              =   1668
   End
   Begin VB.Label Lbl_IconWorks 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "QuickSmith"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   336
      Left            =   1356
      TabIndex        =   1
      Top             =   156
      Width           =   1716
   End
   Begin VB.Label Lbl_Version 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Version 5.0.0"
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
      Height          =   288
      Left            =   3540
      TabIndex        =   2
      Top             =   240
      Width           =   1908
   End
   Begin VB.Label Lbl_Microsoft 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "For Microsoft  32/64 bit Windows"
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
      Height          =   240
      Left            =   480
      TabIndex        =   3
      Top             =   1440
      Width           =   4368
   End
   Begin VB.Label Lbl_Info 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
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
      Height          =   240
      Left            =   645
      TabIndex        =   4
      Top             =   2145
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.Label Lbl_InfoValues 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
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
      Height          =   195
      Left            =   480
      TabIndex        =   0
      Top             =   1800
      Width           =   4050
   End
End
Attribute VB_Name = "AboutBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefInt A-Z
Const WF_STANDARD = &H10
Const WF_ENHANCED = &H20
Const WF_80x87 = &H400

Dim CRLF               As String * 2

Private Sub Cmd_OK_Click()
    Unload AboutBox
End Sub

Private Sub Form_Load()
Dim WinFlags As Long
Dim Mode As String, Processor As String

Mem.dwLength = Len(Mem)
GlobalMemoryStatus Mem

    CRLF = Chr$(13) + Chr$(10)
    ' Dialog Boxes should only have Move and Close items
    ' in their System menus', so remove the others.
    'Remove_Items_From_Sysmenu AboutBox
    ' Center the AboutBox on the screen
     Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    ' Get current Windows configuration
' Function/Sub GETWINFLAGS parameters should be checked against the calling convention on the next line.
' -> Unknown WinAPI/Third-Party DLL Call -> GETWINFLAGS  Check Calling convention for compatibility with BSTR/UNICODE/ANSI Strings. See VB4DLL.TXT in VB directory.
'      WinFlags = GetWinFlags()
'      If WinFlags And WF_ENHANCED Then Mode = "386 Enhanced Mode" Else Mode = "Standard Mode"
'      Lbl_Info.Caption = Mode + CRLF + "Free Memory:" + CRLF + "Math Co-processor:"
'      If WinFlags And WF_80x87 Then Processor = "Present" Else Processor = "None"
'Function/Sub GETFREESPACE parameters should be checked against the calling convention on the next line.
' -> Unknown WinAPI/Third-Party DLL Call -> GETFREESPACE  Check Calling convention for compatibility with BSTR/UNICODE/ANSI Strings. See VB4DLL.TXT in VB directory.

' need to use GlobalMemoryStatus

Lbl_InfoValues.Caption = "Free Physical Memory is " + CStr((Mem.dwAvailPhys / 1024)) + " KB"

End Sub


