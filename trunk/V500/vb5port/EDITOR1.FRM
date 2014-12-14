VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form5 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "QuickSmith Editor"
   ClientHeight    =   5700
   ClientLeft      =   7755
   ClientTop       =   2085
   ClientWidth     =   3855
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
   Icon            =   "EDITOR1.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5700
   ScaleWidth      =   3855
   Begin MSComDlg.CommonDialog CMDialog1 
      Left            =   0
      Top             =   5280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FontSize        =   0
      MaxFileSize     =   256
   End
   Begin VB.TextBox PadText 
      Appearance      =   0  'Flat
      Height          =   5412
      Left            =   156
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   168
      Width           =   3612
   End
   Begin VB.Menu FileItem 
      Caption         =   "&File"
      Begin VB.Menu NewItem 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu LoadItem 
         Caption         =   "&Open File..."
         HelpContextID   =   1101
         Shortcut        =   ^L
      End
      Begin VB.Menu SaveItem 
         Caption         =   "&Save File..."
         HelpContextID   =   1102
         Shortcut        =   ^S
      End
      Begin VB.Menu CloseItem 
         Caption         =   "&Close"
         Enabled         =   0   'False
         Shortcut        =   ^Z
      End
      Begin VB.Menu PrintItem 
         Caption         =   "&Print File..."
         Shortcut        =   ^P
      End
      Begin VB.Menu Separator 
         Caption         =   "-"
      End
      Begin VB.Menu ExitItem 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
      Begin VB.Menu FileArray 
         Caption         =   "-"
         Index           =   0
         Visible         =   0   'False
      End
   End
   Begin VB.Menu EditMenu 
      Caption         =   "&Edit"
      Begin VB.Menu CutItem 
         Caption         =   "&Cut"
         Enabled         =   0   'False
         Shortcut        =   ^C
      End
      Begin VB.Menu PasteItem 
         Caption         =   "&Paste"
         Enabled         =   0   'False
         Shortcut        =   ^V
      End
      Begin VB.Menu ClearItem 
         Caption         =   "Clear &All"
         Enabled         =   0   'False
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu OptionMenu 
      Caption         =   "&Options"
      Begin VB.Menu FFF 
         Caption         =   "Font"
      End
      Begin VB.Menu ForeColorItem 
         Caption         =   "Fore Color"
      End
      Begin VB.Menu BackcolorItem 
         Caption         =   "BackGround Color"
      End
   End
   Begin VB.Menu AboutMenu 
      Caption         =   "&About..."
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CutText As String

Private Sub AboutMenu_Click()
AboutBox.Show 1 'modal
End Sub

Private Sub BackcolorItem_Click()

On Error Resume Next
CMDialog1.CancelError = -1
CMDialog1.DialogTitle = "Background Color"
CMDialog1.Flags = &H1 ' show initial value
CMDialog1.Action = 3

If Err = 0 Then PadText.BackColor = CMDialog1.Color

End Sub

Private Sub ClearItem_Click()
    PadText.Text = ""
    PadText.SetFocus
End Sub

Private Sub CloseItem_Click()
    'If (Dir(CMDialog1.Filename) <> "") Then 'file already exist
     '   response = MsgBox("Overwrite Existing File ?", 4 + 32 + 256, "QuickSmith")
     '  If response = 7 Then
       Form5.Caption = "Text Editor: Untitled"
       PadText.Text = ""
       CloseItem.Enabled = False
      ' Exit Sub
      ' End If
    
    'End If
   ' Open CMDialog1.Filename For Output As #1
   ' Print #1, padText.Text
   ' Close #1
   ' form5.Caption = "Text Editor: Untitled"
   ' padText.Text = ""
   ' closeItem.Enabled = False
End Sub

Private Sub CutItem_Click()
      CutText = PadText.SelText
      PadText.SelText = ""
      PasteItem.Enabled = True
      PadText.SetFocus
End Sub

Private Sub ExitItem_Click()
Form5.Hide
Unload Me
End Sub

Private Sub FFF_Click()
   On Error Resume Next
   CMDialog1.Flags = &H3 Or &H100
   CMDialog1.CancelError = -1
   CMDialog1.DialogTitle = "Select Font"
   CMDialog1.FontName = PadText.FontName
   CMDialog1.FontItalic = PadText.FontItalic
   CMDialog1.FontSize = PadText.FontSize
   CMDialog1.Color = PadText.ForeColor

   CMDialog1.Action = 4
   If Err = 0 Then
   CMDialog1.FontName = PadText.FontName = CMDialog1.FontName
    PadText.FontItalic = CMDialog1.FontItalic
    PadText.FontSize = CMDialog1.FontSize
  PadText.ForeColor = CMDialog1.Color
    
   End If

End Sub

Private Sub filearray_click(index As Integer)
 If index > 0 Then
    Open FileArray(index).Caption For Input As #1
    PadText.Text = Input$(LOF(1), #1)
    Close #1
    CloseItem.Enabled = True
    Form5.Caption = "Text Editor: " & FileArray(index).Caption
    ' UPDATE OPEN FILE HISTORY
    
    Form5.FileArray(0).Visible = True            ' Make the initial element visible / display separator bar.
    ArrayNum = ArrayNum + 1                             ' Increment index property of control array.
    ' Check to see if Filename is already on menu list.
    For I = 0 To ArrayNum - 1
        If Form5.FileArray(I).Caption = FileArray(index).Caption Then
            ArrayNum = ArrayNum - 1
            Exit Sub
        End If
    Next I
    
    ' If filename is not on menu list, add menu item.
    Load Form5.FileArray(ArrayNum)               ' Create a new menu control.
    Form5.FileArray(ArrayNum).Caption = FileArray(index).Caption  ' Set the caption of the new menu item.
    Form5.FileArray(ArrayNum).Visible = True     ' Make the new menu item visible.

End If

End Sub

Private Sub ForeColorItem_Click()
On Error Resume Next
CMDialog1.CancelError = -1
CMDialog1.DialogTitle = "Fore Color"
CMDialog1.Flags = &H1 ' show initial value
CMDialog1.Action = 3
If Err = 0 Then PadText.ForeColor = CMDialog1.Color


End Sub

Private Sub Form_Load()
  Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
   ChDir App.Path
   ChDrive App.Path
   Form5.Show
   ArrayNum = 0
   EdopenFlag = True
    LoadItem_Click

End Sub

Private Sub Form_Unload(Cancel As Integer)
EdopenFlag = False
End Sub

Private Sub LoadItem_Click()
 
   On Error Resume Next
   CMDialog1.CancelError = -1
   CMDialog1.DialogTitle = "EDITOR : Open File "
   CMDialog1.Filter = "Gamma|*.Gam |Schematic Files | *.Sch |S21|*.S21 |Load Temination Files|*.Ter "
   If (SCHl = 0) Then CMDialog1.FilterIndex = 1 Else CMDialog1.FilterIndex = 2
   If (SCHl = 3) Then CMDialog1.FilterIndex = 3
   If (SCHl = 4) Then CMDialog1.FilterIndex = 4
   CMDialog1.Flags = OFN_READONLY
   CMDialog1.Action = 1
   'If ("Text Editor: " + CmDialog1.Filename) Then Exit Sub ' File already loaded
   If Err = 0 Then
    Open CMDialog1.Filename For Input As #1
    PadText.Text = Input$(LOF(1), #1)
    Close #1
    CloseItem.Enabled = True
    Form5.Caption = "Text Editor: " & CMDialog1.Filename
    ' UPDATE OPEN FILE HISTORY
    
    Form5.FileArray(0).Visible = True            ' Make the initial element visible / display separator bar.
    ArrayNum = ArrayNum + 1                             ' Increment index property of control array.
    ' Check to see if Filename is already on menu list.
    For I = 0 To ArrayNum - 1
        If Form5.FileArray(I).Caption = CMDialog1.Filename Then
            ArrayNum = ArrayNum - 1
            Exit Sub
        End If
    Next I
    
    ' If filename is not on menu list, add menu item.
    Load Form5.FileArray(ArrayNum)               ' Create a new menu control.
    Form5.FileArray(ArrayNum).Caption = CMDialog1.Filename  ' Set the caption of the new menu item.
    Form5.FileArray(ArrayNum).Visible = True     ' Make the new menu item visible.

   End If

End Sub

Private Sub NewItem_Click()
PadText.Text = ""
Form5.Caption = "Text Editor: Untitled"
CloseItem.Enabled = True
End Sub

Private Sub PadText_Change()
    CutItem.Enabled = True
    ClearItem.Enabled = True
End Sub

Private Sub PasteItem_Click()
    PadText.SelText = CutText
End Sub

Private Sub PrintItem_Click()
   On Error Resume Next
  
   CMDialog1.CancelError = -1
    CMDialog1.Flags = &H0 Or &H10
    CMDialog1.DialogTitle = "Print File"
   CMDialog1.FromPage = 1
   CMDialog1.ToPage = 1
   CMDialog1.Action = 5
  If Err = 0 Then
  Printer.Print PadText.Text
  Printer.EndDoc
   End If

End Sub

Private Sub SaveItem_Click()
   Dim response
   On Error Resume Next
   CMDialog1.CancelError = -1
   CMDialog1.DialogTitle = "Save File"
 '  Cmdialog1.Filter = "Gamma|*.Gam |Schematic Files | *.Sch |S21|*.S21"
   CMDialog1.Filter = "Gamma|*.Gam |Schematic Files | *.Sch |S21|*.S21 |Load Temination Files|*.Ter "
   If (SCHl = 0) Then CMDialog1.FilterIndex = 1 Else CMDialog1.FilterIndex = 2
   If (SCHl = 3) Then CMDialog1.FilterIndex = 3
   If (SCHl = 4) Then CMDialog1.FilterIndex = 4
   CMDialog1.Flags = &H4    ' hide read only box
   CMDialog1.Action = 2
   If Err = 0 Then
    If (Dir(CMDialog1.Filename) <> "") Then 'file already exist
    response = MsgBox("Overwrite Existing File ?", 4 + 32 + 256, "QuickSmith")
    If response = 7 Then Exit Sub
    End If
    Open CMDialog1.Filename For Output As #1
    Print #1, PadText.Text
    Close #1
   Form5.Caption = "Text Editor: " & CMDialog1.Filename
   End If
  
End Sub

