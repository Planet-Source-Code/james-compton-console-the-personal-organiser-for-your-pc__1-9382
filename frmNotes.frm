VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmNotes 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Console-NOTES"
   ClientHeight    =   4365
   ClientLeft      =   2805
   ClientTop       =   1125
   ClientWidth     =   4620
   LinkTopic       =   "Form1"
   ScaleHeight     =   4365
   ScaleWidth      =   4620
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog Dialog1 
      Left            =   2940
      Top             =   4080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Notes 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   3915
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   60
      Width           =   4455
   End
   Begin VB.Label mnuNotesClear 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Federation"
         Size            =   15.75
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   330
      Left            =   3420
      TabIndex        =   3
      Top             =   4020
      Width           =   1110
   End
   Begin VB.Label mnuNotesLoad 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Load"
      BeginProperty Font 
         Name            =   "Federation"
         Size            =   15.75
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   330
      Left            =   1740
      TabIndex        =   2
      Top             =   4020
      Width           =   1050
   End
   Begin VB.Label mnuNotesSave 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Federation"
         Size            =   15.75
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   330
      Left            =   60
      TabIndex        =   1
      Top             =   4020
      Width           =   1050
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      Height          =   4035
      Left            =   0
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "frmNotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub setoriginalcolour()
mnuNotesLoad.Forecolor = &H80FF&
mnuNotesSave.Forecolor = &H80FF&
mnuNotesClear.Forecolor = &H80FF&
End Sub
Private Sub Label1_Click()

End Sub

Private Sub Form_Load()
zz = SetParent(Me.hwnd, Main.hwnd)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
setoriginalcolour
End Sub

Private Sub mnuNotesClear_Click()
message = MsgBox("Are you sure you want to clear the notepad?", 36, "Sure?")
If message = vbYes Then
Notes.Text = ""
Dialog1.filename = ""
Playsound "Beep.wav"
Else
End If
End Sub

Private Sub mnuNotesClear_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
setoriginalcolour
mnuNotesClear.Forecolor = &HFFFF&
End Sub

Private Sub mnuNotesLoad_Click()
On Error GoTo Error

AlwaysOnTop Main, False 'allow it to be ontop of console

'Open dialog
Dialog1.CancelError = True
Dialog1.DialogTitle = "Load NotePad text"
Dialog1.Filter = "Text Files *.txt|*.txt|Batch Files *.bat|*.bat|INI Files *.ini|*.ini|All Files *.*|*.*"
'Dialog1.filename = ""
Dialog1.ShowOpen
RestoreWindows
If Main.ontop.Value = 1 Then AlwaysOnTop Main, True
If Main.ontop.Value = 0 Then AlwaysOnTop Main, False
If Dialog1.filename = "" Then Exit Sub

'Check size
If FileLen(Dialog1.filename) > 62000 Then
MsgBox "The file is too large to open.", vbCritical, "Error opening file"
Exit Sub
Else
'Load it
Screen.MousePointer = 11
Open Dialog1.filename For Input As #1
        If Err Then Close #1: GoTo Error
        Notes.Text = ""
        'Do Until EOF(1)
            'Line Input #1, a
            'If Err Then Exit Do
            'Notes.Text = Notes.Text & a & Chr(13) & Chr(10)
        'Loop
        Notes.Text = Input(LOF(1), 1) ' This way is quicker
    Close #1
    Screen.MousePointer = 0
    Playsound "Beep.wav"
    Exit Sub
Notes.SetFocus

Error:
If Err.Number <> 32755 Then
MsgBox "Error loading file - " & Err.Number & " : " & Err.Description
Screen.MousePointer = 0
Close #1
RestoreWindows
If Main.ontop.Value = 1 Then AlwaysOnTop Main, True
If Main.ontop.Value = 0 Then AlwaysOnTop Main, False
Exit Sub
Else
End If
End If
RestoreWindows
If Main.ontop.Value = 1 Then AlwaysOnTop Main, True
If Main.ontop.Value = 0 Then AlwaysOnTop Main, False
End Sub

Private Sub mnuNotesLoad_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
setoriginalcolour
mnuNotesLoad.Forecolor = &HFFFF&
End Sub

Private Sub mnuNotesSave_Click()
AlwaysOnTop Main, False
On Error GoTo Error
'Open Dialog
Dialog1.CancelError = True
Dialog1.DialogTitle = "Save NotePad text"
Dialog1.Filter = "Text Files *.txt|*.txt|All Files *.*|*.*"
'Dialog1.filename = "Notes.txt"
Dialog1.ShowSave
RestoreWindows
If Main.ontop.Value = 1 Then AlwaysOnTop Main, True
If Main.ontop.Value = 0 Then AlwaysOnTop Main, False
'Save it
Open Dialog1.filename For Output As #1
Print #1, Notes.Text
Close #1
Playsound "Beep.wav"
Exit Sub

Error:
If Err.Number <> 32755 Then
MsgBox "Error saving file - " & Err.Number & " : " & Err.Description
RestoreWindows
If Main.ontop.Value = 1 Then AlwaysOnTop Main, True
If Main.ontop.Value = 0 Then AlwaysOnTop Main, False
Exit Sub
Else
End If
RestoreWindows
If Main.ontop.Value = 1 Then AlwaysOnTop Main, True
If Main.ontop.Value = 0 Then AlwaysOnTop Main, False
End Sub

Private Sub mnuNotesSave_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
setoriginalcolour
mnuNotesSave.Forecolor = &HFFFF&
End Sub

Private Sub Notes_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
setoriginalcolour
End Sub
