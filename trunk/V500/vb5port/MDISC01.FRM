VERSION 5.00
Begin VB.MDIForm MDIForm1 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000C&
   Caption         =   "QuickSmith"
   ClientHeight    =   5775
   ClientLeft      =   1050
   ClientTop       =   2850
   ClientWidth     =   9600
   Icon            =   "MDISC01.frx":0000
   LinkTopic       =   "MDIForm1"
   Begin VB.Menu FileMenu 
      Caption         =   "File"
      Begin VB.Menu OpencItem 
         Caption         =   "Open Chart"
      End
      Begin VB.Menu OpenSItem 
         Caption         =   "Open Schematic"
      End
   End
   Begin VB.Menu ExitMenu 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ExitMenu_Click()
If (EdopenFlag = True) Then Unload Form5
 If (PltopenFlag = True) Then Unload Form3
If (SPAROPENFLAG = True) Then Unload Form9
If (NOISEOPENFLAG = True) Then Unload Form10
If (AMPOPENFLAG = True) Then Unload Form6
If (SweepopenFlag = True) Then Unload Form8
Unload Me
End Sub

Private Sub MDIForm_Load()
 
   'Call checkdate


    ' Check video mode.  If less than EGA, terminate QuickSmith
    '
    If Screen.Height < EGA_HEIGHT Then
        MsgBox "QuickSmith requires EGA or Better.", 16, "QuickSmith"
        End
    Else
        ' Since you cannot assign values like TAB, CR, and LF to string
        ' constants, the values of TAB and CRLF which are used frequently
        ' thoughout QuickSmith when displaying messages, these values are
        ' are assigned to the global string values of A_TAB and CRLF
        '
        'A_TAB = Chr$(9)
        CRLF = Chr$(13) + Chr$(10)
    End If
        If Not Help_File_In_Path() Then
            Text = "QS.HLP not found in your path." + CRLF + CRLF
            Text = Text + "Windows searches your PATH environment variable for help files, "
            Text = Text + "so you need to copy QS.HLP to a directory included in your "
            Text = Text + "PATH if you wish to obtain help while running QuickSmith."
            MsgBox Text, 48, "QuickSmith help not available"
        End If

Height = Screen.Height
Width = Screen.Width
Left = 0
Top = 0

Load form1
Load form2
If Screen.Height < 7500 Then
    form1.Height = 0.85 * Height
    form1.Width = 0.44 * Width
    form2.Height = 0.85 * Height
    form2.Width = 0.45 * Width
End If
form2.Left = form1.Width + 10
form2.Top = form1.Top

'form2.Show
'form1.Show

mdiArrange MDITILE_VERTICAL   ' Tile vertical by default added ver4.2

 
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
 Dim dummy As String
 R = WinHelp(hwnd, dummy, HELP_QUIT, CLng(0))
 End
End Sub

Private Sub OpencItem_Click()
Load form1

End Sub

Private Sub OpenItem_Click()
form1.Show

End Sub

Private Sub OpenSItem_Click()
Load form2


End Sub

Private Sub WindowMenu_Click()
form1.Show
End Sub

