VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPlaylist 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Console-Playlist Editor"
   ClientHeight    =   4365
   ClientLeft      =   2805
   ClientTop       =   1125
   ClientWidth     =   4620
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   4620
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog Dialog1 
      Left            =   1860
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ListBox plyLIST 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00C0C0C0&
      Height          =   1785
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "Playlist - Double click to see details, right click to bring up menu."
      Top             =   2220
      Width           =   2655
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00C0C0C0&
      Height          =   1980
      Left            =   2400
      MultiSelect     =   2  'Extended
      Pattern         =   $"frmPlaylist.frx":0000
      TabIndex        =   2
      ToolTipText     =   "Double click a file to add it to the playlist"
      Top             =   120
      Width           =   2115
   End
   Begin VB.DirListBox Dir1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00C0C0C0&
      Height          =   1665
      Left            =   120
      TabIndex        =   1
      Top             =   420
      Width           =   2175
   End
   Begin VB.DriveListBox Drive1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00C0C0C0&
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   2295
   End
   Begin VB.Label filenametext 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Federation"
         Size            =   8.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   4140
      Width           =   2715
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Federation"
         Size            =   9.75
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   225
      Left            =   2880
      TabIndex        =   11
      Top             =   2640
      Width           =   1710
   End
   Begin VB.Label mnuTRANS 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Transfer"
      BeginProperty Font 
         Name            =   "Federation"
         Size            =   9.75
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   225
      Left            =   2880
      TabIndex        =   10
      Top             =   4140
      Width           =   1710
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00808080&
      Height          =   2115
      Left            =   60
      Top             =   60
      Width           =   2295
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00808080&
      Height          =   2115
      Left            =   2340
      Top             =   60
      Width           =   2235
   End
   Begin VB.Label mnuLOAD 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Load"
      BeginProperty Font 
         Name            =   "Federation"
         Size            =   9.75
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   225
      Left            =   2880
      TabIndex        =   9
      Top             =   3660
      Width           =   1710
   End
   Begin VB.Label mnuSave 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Federation"
         Size            =   9.75
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   225
      Left            =   2880
      TabIndex        =   8
      Top             =   3900
      Width           =   1710
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Move Down"
      BeginProperty Font 
         Name            =   "Federation"
         Size            =   9.75
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   225
      Left            =   2880
      TabIndex        =   7
      Top             =   3300
      Width           =   1710
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Move Up"
      BeginProperty Font 
         Name            =   "Federation"
         Size            =   9.75
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   225
      Left            =   2880
      TabIndex        =   6
      Top             =   3060
      Width           =   1710
   End
   Begin VB.Label mnuREMOVE 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Remove"
      BeginProperty Font 
         Name            =   "Federation"
         Size            =   9.75
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   225
      Left            =   2880
      TabIndex        =   5
      Top             =   2400
      Width           =   1710
   End
   Begin VB.Label mnuADD 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "Federation"
         Size            =   9.75
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   225
      Left            =   2880
      TabIndex        =   4
      Top             =   2160
      Width           =   1710
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      Height          =   1935
      Left            =   60
      Top             =   2160
      Width           =   2775
   End
End
Attribute VB_Name = "frmPlaylist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private lngLstIH As Long, lngLstWid As Long, lngLstMax As Long, lngScrollW As Long
Private Sub setoriginalcolour()
mnuADD.Forecolor = &H80FF&
mnuREMOVE.Forecolor = &H80FF&
Label1.Forecolor = &H80FF&
Label2.Forecolor = &H80FF&
mnuSave.Forecolor = &H80FF&
mnuLOAD.Forecolor = &H80FF&
mnuTRANS.Forecolor = &H80FF&
Label3.Forecolor = &H80FF&

End Sub

Private Function PathCheck()
' If dragged file is in the root, append filename.
If Mid(Dir1.Path, Len(Dir1.Path)) = "\" Then
PathCheck = Dir1.Path
' If dragged file is not in root, append "\" and filename.
Else
PathCheck = Dir1.Path & "\"
End If
End Function
Private Sub Dir1_Change()
On Error GoTo Error
File1.Path = Dir1.Path
Exit Sub

Error:
MsgBox "Error selecting directory - " & Err.Number & " : " & Err.Description
End Sub

Private Sub Dir1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
setoriginalcolour
End Sub

Private Sub Drive1_Change()
On Error GoTo Error
Dir1.Path = Drive1.Drive
Exit Sub

Error:
MsgBox "Error selecting drive - " & Err.Number & " : " & Err.Description
End Sub

Private Sub File1_DblClick()
If File1.ListIndex = -1 Then Exit Sub
For X = 0 To File1.ListCount - 1
    If File1.Selected(X) = True Then 'Found a selected item
        plyLIST.AddItem PathCheck & File1.List(X) 'Add to second list
    End If
DoEvents
Next X
'plyLIST.AddItem PathCheck & File1.filename
End Sub

Private Sub File1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
setoriginalcolour
Dim idx As Integer, strItem As String

    'calculate index position by diving y (twips) by the height of a list item
    'by the height of a single item (converted to twips) and adding it to
    'TopIndex - the index of the value at the top of the list
    idx = File1.TopIndex + (Y \ (lngLstIH * Screen.TwipsPerPixelY))

    If idx > File1.ListCount - 1 Then Exit Sub 'pointing at blank space under items!
    
    strItem = File1.List(idx)
    'listboxes don't have TextWidth methods, but forms do.
    'make sure Font Properties are same for Form as for listbox
    'if they _need_ to be different, you can use the TextWidth method of a picturebox
    '(which can be hidden if you don't need it otherwise)
    
    'the following statement tests the width of the text (in Twips)
    'against the width of the inside of the listbox
    'The IIF statement returns the width of a scroll bar if one is showing
    'by texting to see if the listcount is greater than the max lines showing in
    'the listbox (or returns 0 if it is not).  That value is subtracted from the
    'width to give the 'visible' width.  That value is then converted to Twips.
    
    If Me.TextWidth(strItem) >= (lngLstWid - _
        IIf(File1.ListCount > lngLstMax, lngScrollW, 0)) _
        * Screen.TwipsPerPixelY Then
        'If true, then item pointed at is too wide to see
        File1.ToolTipText = strItem
    Else
        File1.ToolTipText = "" 'Turn it off if it was on!
    End If
End Sub

Private Sub Form_Load()
zz = SetParent(Me.hwnd, Main.hwnd)
'LISTBOX STUFF
Dim lReturn As Long, recLst As RECT
    'Height in pixels of an item on the list
    lngLstIH = SendMessageByNum(File1.hwnd, LB_GETITEMHEIGHT, 0, 0)
    'get dimensions of listbox in pixels
    'lReturn is a dummy var.  The real data goes into recLst
    lReturn = GetClientRect(File1.hwnd, recLst)
    'Number of lines on the list
    lngLstMax = (recLst.Bottom - recLst.Top) \ lngLstIH
    'Width of inside of listbox (not including borders) in pixels
    lngLstWid = (recLst.Right - recLst.Left)
    'Width of a vertical ScrollBar in pixels
    lngScrollW = GetSystemMetrics(SM_CYVSCROLL)
End Sub

Private Sub Label1_Click()
On Error Resume Next
Dim tmpLIST As String
Dim tmpNO As Integer
Dim tmpLIST2 As String

tmpLIST = plyLIST.Text
tmpNO = plyLIST.ListIndex
plyLIST.ListIndex = plyLIST.ListIndex - 1
tmpLIST2 = plyLIST.Text

plyLIST.List(plyLIST.ListIndex) = tmpLIST

plyLIST.ListIndex = tmpNO
plyLIST.List(plyLIST.ListIndex) = tmpLIST2
plyLIST.ListIndex = plyLIST.ListIndex - 1
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
setoriginalcolour
Label1.Forecolor = &HFFFF&
End Sub

Private Sub Label2_Click()
On Error Resume Next
Dim tmpLIST As String
Dim tmpNO As Integer
Dim tmpLIST2 As String

tmpLIST = plyLIST.Text
tmpNO = plyLIST.ListIndex
plyLIST.ListIndex = plyLIST.ListIndex + 1
tmpLIST2 = plyLIST.Text

plyLIST.List(plyLIST.ListIndex) = tmpLIST

plyLIST.ListIndex = tmpNO
plyLIST.List(plyLIST.ListIndex) = tmpLIST2
plyLIST.ListIndex = plyLIST.ListIndex + 1
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
setoriginalcolour
Label2.Forecolor = &HFFFF&
End Sub

Private Sub Label3_Click()
message = MsgBox("Are you sure you want to clear the playlist?", 36, "Sure?")
If message = vbYes Then
plyLIST.Clear
Playsound "Beep.wav"
Exit Sub
Else
End If
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
setoriginalcolour
Label3.Forecolor = &HFFFF&
End Sub

Private Sub mnuADD_Click()
If File1.ListIndex = -1 Then Exit Sub
If File1.ListIndex = -1 Then Exit Sub
For X = 0 To File1.ListCount - 1
    If File1.Selected(X) = True Then 'Found a selected item
        plyLIST.AddItem PathCheck & File1.List(X) 'Add to second list
    End If
DoEvents
Next X
End Sub

Private Sub mnuADD_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
setoriginalcolour
mnuADD.Forecolor = &HFFFF&
End Sub

Private Sub mnuLOAD_Click()
Dim TempString As String
AlwaysOnTop Main, False
'Save list to a file - simple!
On Error GoTo Error
Dialog1.CancelError = True
Dialog1.DialogTitle = "Load audio/video playlist"
Dialog1.Filter = "Audio/Video playlists *.avp|*.avp"
Dialog1.filename = ""
Dialog1.ShowOpen
RestoreWindows
If Main.ontop.Value = 1 Then AlwaysOnTop Main, True
If Main.ontop.Value = 0 Then AlwaysOnTop Main, False

Open Dialog1.filename For Input As #1
Screen.MousePointer = 11
plyLIST.Clear
Do Until EOF(1)
    Line Input #1, TempString
    plyLIST.AddItem TempString
Loop
Close #1
Screen.MousePointer = 0
filenametext.Caption = Mid(Dialog1.filename, InStrRevVB5(Dialog1.filename, "\") + 1, Len(Dialog1.filename))
Playsound "Beep.wav"
Exit Sub

Error:
Close #1
Screen.MousePointer = 0
If Err.Number <> 32755 Then
MsgBox "Error loading file - " & Err.Number & " : " & Err.Description
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

Private Sub mnuLOAD_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
setoriginalcolour
mnuLOAD.Forecolor = &HFFFF&
End Sub

Private Sub mnuREMOVE_Click()
On Error Resume Next
plyLIST.RemoveItem plyLIST.ListIndex
End Sub

Private Sub mnuREMOVE_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
setoriginalcolour
mnuREMOVE.Forecolor = &HFFFF&
End Sub

Private Sub mnuSave_Click()
'Save list to a file - simple!
AlwaysOnTop Main, False
On Error GoTo Error
Dialog1.CancelError = True
Dialog1.DialogTitle = "Save audio/video playlist"
Dialog1.Filter = "Audio/Video playlists *.avp|*.avp"
Dialog1.filename = filenametext.Caption
Dialog1.ShowSave

RestoreWindows
If Main.ontop.Value = 1 Then AlwaysOnTop Main, True
If Main.ontop.Value = 0 Then AlwaysOnTop Main, False

plyLIST.Visible = False
Open Dialog1.filename For Output As #1
Screen.MousePointer = 11
For X = 0 To plyLIST.ListCount - 1
    plyLIST.ListIndex = X
    Print #1, plyLIST.Text
Next
Close #1
plyLIST.Visible = True ' bring it back!
Screen.MousePointer = 0
filenametext.Caption = Mid(Dialog1.filename, InStrRevVB5(Dialog1.filename, "\") + 1, Len(Dialog1.filename))
Playsound "Beep.wav"
Exit Sub

Error:
Close #1
Screen.MousePointer = 0
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

Private Sub mnuSave_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
setoriginalcolour
mnuSave.Forecolor = &HFFFF&
End Sub

Private Sub mnuTRANS_Click()
If plyLIST.ListCount = 0 Then
MsgBox "There are no files to transfer!", vbCritical
Exit Sub
Else
frmMedia.playlistfile.Caption = 1
frmMedia.plyLIST.Clear
frmMedia.Media1.AutoStart = True
frmMedia.Media1.Visible = True
plyLIST.Visible = False ' to stop scrolling
For X = 0 To plyLIST.ListCount - 1
DoEvents
    plyLIST.ListIndex = X
    frmMedia.plyLIST.AddItem (plyLIST.Text)
Next
plyLIST.Visible = True ' bring it back!
frmMedia.plyLIST.ListIndex = 0

On Error Resume Next
If frmMedia.randomtrk.Value = 1 Then GoTo Random
If frmMedia.plyLIST.Text <> "" Then
frmMedia.Media1.filename = frmMedia.plyLIST.Text
frmMedia.Media1.Play
frmMedia.Dialog1.filename = ""
Playsound "Beep.wav"
Exit Sub

Random:
Randomize
frmMedia.plyLIST.ListIndex = Int(Rnd * frmMedia.plyLIST.ListCount)
frmMedia.Media1.filename = frmMedia.plyLIST.Text
frmMedia.Dialog1.filename = ""
Playsound "Beep.wav"
End If
Exit Sub
End If

End Sub

Private Sub mnuTRANS_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
setoriginalcolour
mnuTRANS.Forecolor = &HFFFF&
End Sub

Private Sub plyLIST_DblClick()
MsgBox "File : " & plyLIST.Text & Chr(13) & Chr(10) & "Track No : " & plyLIST.ListIndex + 1 & Chr(13) & Chr(10) & "Total : " & plyLIST.ListCount, vbInformation, "File Information"
End Sub

Private Sub plyLIST_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
setoriginalcolour
End Sub

Private Sub plyLIST_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
frmmen.mediaplayerplaytrack.Caption = "Play " & Mid(frmPlaylist.plyLIST.Text, InStrRevVB5(frmPlaylist.plyLIST.Text, "\") + 1, Len(frmPlaylist.plyLIST.Text))
PopupMenu frmmen.frmpop
Else
End If
End Sub
