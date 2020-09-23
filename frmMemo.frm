VERSION 5.00
Begin VB.Form frmMemo 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Console-MEMO"
   ClientHeight    =   4365
   ClientLeft      =   2805
   ClientTop       =   1125
   ClientWidth     =   4620
   LinkTopic       =   "Form1"
   ScaleHeight     =   4365
   ScaleWidth      =   4620
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Memo 
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
      Height          =   2595
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   1680
      Width           =   4395
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Federation"
         Size            =   9.75
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   1380
      Left            =   120
      Pattern         =   "*.mem"
      TabIndex        =   0
      Top             =   120
      Width           =   2475
   End
   Begin VB.Label Loading 
      Caption         =   "0"
      Height          =   195
      Left            =   2640
      TabIndex        =   4
      Top             =   60
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Label memDelete 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "Federation"
         Size            =   14.25
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   330
      Left            =   2760
      TabIndex        =   3
      Top             =   900
      Width           =   1770
   End
   Begin VB.Label memCreate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Create"
      BeginProperty Font 
         Name            =   "Federation"
         Size            =   14.25
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   285
      Left            =   2760
      TabIndex        =   2
      Top             =   300
      Width           =   1770
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00808080&
      Height          =   2715
      Left            =   60
      Top             =   1620
      Width           =   4515
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      Height          =   1515
      Left            =   60
      Top             =   60
      Width           =   2595
   End
End
Attribute VB_Name = "frmMemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub setoriginalcolour()
memCreate.Forecolor = &H80FF&
memDelete.Forecolor = &H80FF&
End Sub
Private Function ProgPath()
' If dragged file is in the root, append filename.
If Mid(App.Path, Len(App.Path)) = "\" Then
ProgPath = App.Path
' If dragged file is not in root, append "\" and filename.
Else
ProgPath = App.Path & "\"
End If
End Function

Private Sub File1_Click()
Dim A As String
On Error GoTo Error
'Load it
Loading.Caption = 1
Call ProgPath
Open ProgPath & "Memo\" & File1.filename For Input As #1
        If Err Then Close #1: GoTo Error
        Memo.Text = ""
        'Do Until EOF(1)
         '   Line Input #1, A
          '  If Err Then Exit Do
           ' Memo.Text = Memo.Text & A & Chr(13) & Chr(10)
           Memo.Text = Input(LOF(1), 1)
        'Loop
    Close #1
Loading.Caption = 0
    Exit Sub
Memo.SetFocus

Error:
If Err.Number <> 32755 Then
Loading.Caption = 0
MsgBox "Error opening memo - " & Err.Number & " : " & Err.Description
Close #1
Exit Sub
Else
End If
Me.Show
Loading.Caption = 0
End Sub

Private Sub File1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
setoriginalcolour
End Sub

Private Sub Form_Load()
zz = SetParent(Me.hwnd, Main.hwnd)
Call ProgPath
On Error GoTo Error
File1.Path = ProgPath & "Memo"
Exit Sub
Error:
MsgBox "Error loading memo list - " & Err.Number & " : " & Err.Description
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
setoriginalcolour
End Sub

Private Sub memCreate_Click()
Dim MemoName As String
Dim LSTNO As Integer
On Error GoTo Error
Loading.Caption = 1
MemoName = InputBox("Enter the subject title", "Subject")
If MemoName = "" Then Exit Sub
Open ProgPath & "Memo\" & MemoName & ".mem" For Output As #1
        If Err Then Close #1: GoTo Error
        Memo.Text = "Memo for " & Date & " at " & Format(Time, "HH:MM:SS AMPM") & Chr(13) & Chr(10)
        Print #1, "Memo for " & Date & " at " & Format(Time, "HH:MM:SS AMPM") & Chr(13) & Chr(10)
    Close #1
    File1.Refresh
    LSTNO = FindExactInCombo(File1, MemoName & ".mem")
    File1.ListIndex = LSTNO
    Loading.Caption = 0
    Me.Show
    Memo.SetFocus
    
'Memo message count
If frmMemo.File1.ListCount <> 1 Then
Main.MemoMess.Caption = " -------------------------" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "You have " & frmMemo.File1.ListCount & " messages in your memo."
Else
Main.MemoMess.Caption = " -------------------------" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "You have " & frmMemo.File1.ListCount & " message in your memo."
End If
    Playsound "Beep.wav"
    Exit Sub
    
Error:
Loading.Caption = 0
MsgBox "Error creating memo - " & Err.Number & " : " & Err.Description
Loading.Caption = 0

End Sub

Private Sub memCreate_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
setoriginalcolour
memCreate.Forecolor = &HFFFF&
End Sub

Private Sub memDelete_Click()
On Error GoTo Error
If File1.ListIndex <> -1 Then
Loading.Caption = 1
message = MsgBox("Are you sure you want to delete " & File1.filename & "?", 36, "Delete memo?")
If message = 6 Then

Call ProgPath
Kill ProgPath & "Memo\" & File1.filename
File1.Refresh
Memo.Text = ""
Loading.Caption = 0
'Memo message count
If frmMemo.File1.ListCount <> 1 Then
Main.MemoMess.Caption = " -------------------------" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "You have " & frmMemo.File1.ListCount & " messages in your memo."
Else
Main.MemoMess.Caption = " -------------------------" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "You have " & frmMemo.File1.ListCount & " message in your memo."
End If
Playsound "Beep.wav"
Exit Sub
Error:
MsgBox "Error deleting memo - " & Err.Number & " : " & Err.Description
Else
End If
Else
End If
End Sub

Private Sub memDelete_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
setoriginalcolour
memDelete.Forecolor = &HFFFF&
End Sub

Private Sub Memo_Change()
On Error GoTo Error
'Save it
If Loading.Caption = 0 Then
Call ProgPath
Open ProgPath & "\Memo\" & File1.filename For Output As #1
Print #1, Memo.Text
Close #1
Exit Sub

Error:
If Err.Number <> 32755 Then
MsgBox "Error saving memo - " & Err.Number & " : " & Err.Description
Exit Sub
Else
End If
Else
End If
End Sub

Private Sub Memo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
setoriginalcolour
End Sub
