VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "ForNote"
   ClientHeight    =   6975
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   9885
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6975
   ScaleWidth      =   9885
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2640
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      Height          =   6975
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   9855
   End
   Begin VB.Menu File 
      Caption         =   "&File"
      Begin VB.Menu New 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu Open 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu BtSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu BtSaveAs 
         Caption         =   "&Save As"
      End
      Begin VB.Menu Spcestrie 
         Caption         =   "-"
      End
      Begin VB.Menu Print 
         Caption         =   "&Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu Stripe 
         Caption         =   "-"
      End
      Begin VB.Menu Exit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu Edit 
      Caption         =   "&Edit"
      Begin VB.Menu Undo 
         Caption         =   "&Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu Space 
         Caption         =   "-"
      End
      Begin VB.Menu Cut 
         Caption         =   "&Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu Copy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu Paste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu Del 
         Caption         =   "&Delete"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu sss 
         Caption         =   "-"
      End
      Begin VB.Menu BtSellect 
         Caption         =   "&Select All"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu About 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private TempText As String
Private FilePath As String
Private Function ShowSave()
    Dim result As Boolean
    result = False
    CommonDialog1.Filter = "All Files|*.*"
    CommonDialog1.ShowSave
    If (CommonDialog1.filename <> "") Then
        result = True
        SaveFile CommonDialog1.filename
    End If
    ShowSave = result
End Function

Private Function SaveAs()
    Dim result As Boolean
    If (ShowSave) Then
        FilePath = CommonDialog1.filename
        TempText = Text1.Text
        result = True
    Else
        result = False
    End If
    SaveAs = result
End Function

Private Function Save()
    Dim result As Boolean
    result = True
    If (FilePath = "") Then
        If (Not SaveAs) Then
            result = False
        End If
    Else
        SaveFile FilePath
        TempText = Text1.Text
    End If
    Save = result
End Function



Private Sub About_Click()
frmAbout.Show
End Sub

Private Sub BtSellect_Click()
Text1.SelStart = 0
Text1.SelLength = Len(ActiveControl.Text)
End Sub

Private Sub Copy_Click()
Clipboard.Clear
Clipboard.SetText Text1.SelText
End Sub

Private Sub Cut_Click()
Clipboard.Clear
Clipboard.SetText Text1.SelText
Text1.SelText = ""
End Sub

Private Sub Del_Click()
Text1.SelText = ""
End Sub

Private Sub Exit_Click()
On Error GoTo ErrorHandler
Dim Msg, Style, Title, Response, MyString
Msg = "Are you sure you want to exit ?"
Style = vbYesNo + vbQuestion + vbDefaultButton1
Title = "Warning"
Response = MsgBox(Msg, Style, Title)
If Response = vbYes Then
MyString = "Yes"
End
End If
ErrorHandler:
End Sub


Private Sub Form_Resize()
  Text1.Width = Me.ScaleWidth - (Text1.Left * 2)
  Text1.Height = Me.ScaleHeight - (Text1.Top * 2)
End Sub

Private Sub New_Click()
NewFile = MsgBox("Save File?", vbYesNoCancel + vbQuestion, "New")
If NewFile = vbYes Then
Save
Text1.Text = ""
Else
If NewFile = vbNo Then
Text1.Text = ""
End If
End If
End Sub

Private Sub Open_Click()
CommonDialog1.Filter = "All Files (*.*)|*.*|Text Files (*.txt)|*.txt"
CommonDialog1.FilterIndex = 2
CommonDialog1.ShowOpen
Dim LoadFileToTB As Boolean
Dim TxtBox As Object
Dim FilePath As String
Dim Append As Boolean
Dim iFile As Integer
Dim s As String
If Dir(FilePath) = "" Then Exit Sub
On Error GoTo ErrorHandler:
s = Text1.Text
iFile = FreeFile
Open CommonDialog1.filename For Input As #iFile
s = Input(LOF(iFile), #iFile)
If Append Then
Text1.Text = Text1.Text & s
Else
Text1.Text = s
End If
LoadFileToTB = True
ErrorHandler:
If iFile > 0 Then Close #iFile
End Sub

Private Sub Paste_Click()
Text1.SelText = Clipboard.GetText()
End Sub

Private Sub Print_Click()
On Error GoTo ErrHandler
Dim BeginPage, EndPage, NumCopies, i
CommonDialog1.CancelError = True
CommonDialog1.ShowPrinter
BeginPage = CommonDialog1.FromPage
EndPage = CommonDialog1.ToPage
NumCopies = CommonDialog1.Copies
For i = 1 To NumCopies
Printer.Print Text1.Text
Next i
Exit Sub
ErrHandler:
Exit Sub
End Sub
Private Sub SaveFile(filename As String)
    Open filename For Output As #1
        Print #1, Text1.Text
    Close #1
End Sub
Private Sub BtSave_Click()
   Save
End Sub
Private Sub BtSaveAs_Click()
On Error GoTo ErrorHandler
CommonDialog1.Filter = "All Files (*.*)|*.*|Text Files (*.txt)|*.txt"
CommonDialog1.FilterIndex = 2
CommonDialog1.ShowSave
CommonDialog1.filename = CommonDialog1.filename
Dim iFile As Integer
Dim SaveFileFromTB As Boolean
Dim TxtBox As Object
Dim FilePath As String
Dim Append As Boolean
iFile = FreeFile
If Append Then
Open CommonDialog1.filename For Append As #iFile
Else
Open CommonDialog1.filename For Output As #iFile
End If
Print #iFile, Text1.Text
SaveFileFromTB = True
ErrorHandler:
Close #iFile
End Sub

Private Sub Undo_Click()
    Text1.SetFocus
    SendKeys "^Z", 1
End Sub
Private Sub Edit_Click()

If Text1.SelLength > 0 Then
Cut.Enabled = True
Copy.Enabled = True
Del.Enabled = True
Else
Cut.Enabled = False
Copy.Enabled = False
Del.Enabled = False
End If
End Sub
