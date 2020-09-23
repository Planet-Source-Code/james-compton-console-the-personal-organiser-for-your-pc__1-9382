VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Begin VB.Form frmMedia 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Console-MEDIA"
   ClientHeight    =   4365
   ClientLeft      =   2805
   ClientTop       =   1125
   ClientWidth     =   4620
   LinkTopic       =   "Form1"
   ScaleHeight     =   4365
   ScaleWidth      =   4620
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Anitime 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   1485
      Top             =   3825
   End
   Begin VB.PictureBox Ani 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   3300
      Left            =   45
      ScaleHeight     =   3300
      ScaleWidth      =   4515
      TabIndex        =   20
      Top             =   0
      Width           =   4515
   End
   Begin VB.CheckBox randomtrk 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Federation"
         Size            =   6
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   1080
      TabIndex        =   18
      ToolTipText     =   "Play a random track in a AVP playlist."
      Top             =   3360
      Width           =   255
   End
   Begin VB.ListBox plyLIST 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Federation"
         Size            =   6
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   165
      Left            =   120
      TabIndex        =   16
      Top             =   3900
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.PictureBox posper 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   120
      ScaleHeight     =   105
      ScaleWidth      =   1245
      TabIndex        =   12
      Top             =   3900
      Width           =   1275
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   480
      Top             =   2760
   End
   Begin MSComDlg.CommonDialog Dialog1 
      Left            =   0
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MediaPlayerCtl.MediaPlayer Media1 
      Height          =   3285
      Left            =   45
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   4530
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   0   'False
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   0   'False
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   0   'False
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   -1  'True
      VideoBorderWidth=   0
      VideoBorderColor=   8421504
      VideoBorder3D   =   0   'False
      Volume          =   -60
      WindowlessVideo =   0   'False
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Rnd."
      BeginProperty Font 
         Name            =   "Federation"
         Size            =   8.25
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   1080
      TabIndex        =   19
      Top             =   3600
      Width           =   555
   End
   Begin VB.Label playlistfile 
      Caption         =   "0"
      Height          =   255
      Left            =   4080
      TabIndex        =   17
      Top             =   3900
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "List"
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
      Left            =   60
      TabIndex        =   15
      Top             =   3600
      Width           =   990
   End
   Begin VB.Label medNEXT 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   ">|"
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
      Left            =   4020
      TabIndex        =   14
      Top             =   3660
      Width           =   570
   End
   Begin VB.Label medPREV 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "|<"
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
      Left            =   1620
      TabIndex        =   13
      Top             =   3660
      Width           =   570
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00808080&
      Height          =   255
      Left            =   60
      Top             =   3840
      Width           =   1395
   End
   Begin VB.Label mediafile 
      BackStyle       =   0  'Transparent
      Caption         =   "NO FILE"
      BeginProperty Font 
         Name            =   "Federation"
         Size            =   6
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   1560
      TabIndex        =   11
      Top             =   3960
      Width           =   2895
   End
   Begin VB.Label medpl 
      Caption         =   "0"
      Height          =   255
      Left            =   4380
      TabIndex        =   10
      Top             =   4080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label MEDlp 
      Caption         =   "0"
      Height          =   195
      Left            =   4200
      TabIndex        =   9
      Top             =   4080
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label medLoop 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Loop"
      BeginProperty Font 
         Name            =   "Federation"
         Size            =   11.25
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   270
      Left            =   2640
      TabIndex        =   8
      Top             =   3360
      Width           =   930
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      Height          =   255
      Left            =   60
      Top             =   4080
      Width           =   1395
   End
   Begin VB.Label medTIME 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Federation"
         Size            =   9
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   60
      TabIndex        =   7
      Top             =   4095
      Width           =   1395
   End
   Begin VB.Label medRW 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "<<"
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
      Left            =   2220
      TabIndex        =   6
      Top             =   3660
      Width           =   630
   End
   Begin VB.Label medFF 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   ">>"
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
      Left            =   3360
      TabIndex        =   5
      Top             =   3660
      Width           =   630
   End
   Begin VB.Label medPause 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "||"
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
      Top             =   3660
      Width           =   450
   End
   Begin VB.Label medStop 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Stop"
      BeginProperty Font 
         Name            =   "Federation"
         Size            =   11.25
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   270
      Left            =   3660
      TabIndex        =   3
      Top             =   3360
      Width           =   930
   End
   Begin VB.Label medPlay 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Play"
      BeginProperty Font 
         Name            =   "Federation"
         Size            =   11.25
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   270
      Left            =   1620
      TabIndex        =   2
      Top             =   3360
      Width           =   930
   End
   Begin VB.Label medLoad 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
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
      Left            =   60
      TabIndex        =   1
      Top             =   3360
      Width           =   990
   End
End
Attribute VB_Name = "frmMedia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CheckForAudioPlaylist()
If (UCase(Right(plyLIST.Text, 4)) = ".MP3" Or UCase(Right(plyLIST.Text, 4)) = ".WAV" Or UCase(Right(plyLIST.Text, 4)) = ".MID" Or UCase(Right(plyLIST.Text, 4)) = ".AIF" Or UCase(Right(plyLIST.Text, 5)) = ".AIFC" Or UCase(Right(plyLIST.Text, 5)) = ".AIFF" Or UCase(Right(plyLIST.Text, 4)) = ".WMA" Or UCase(Right(plyLIST.Text, 3)) = ".AU" Or UCase(Right(plyLIST.Text, 4)) = ".SND" Or UCase(Right(plyLIST.Text, 4)) = ".M3U") And Media1.PlayState = mpPlaying Then
Ani.Visible = True
Anitime.Enabled = True
Media1.Visible = False
Else

If (UCase(Right(plyLIST.Text, 4)) = ".MP3" Or UCase(Right(plyLIST.Text, 4)) = ".WAV" Or UCase(Right(plyLIST.Text, 4)) = ".MID" Or UCase(Right(plyLIST.Text, 4)) = ".AIF" Or UCase(Right(plyLIST.Text, 5)) = ".AIFC" Or UCase(Right(plyLIST.Text, 5)) = ".AIFF" Or UCase(Right(plyLIST.Text, 4)) = ".WMA" Or UCase(Right(plyLIST.Text, 3)) = ".AU" Or UCase(Right(plyLIST.Text, 4)) = ".SND" Or UCase(Right(plyLIST.Text, 4)) = ".M3U") And Media1.PlayState = mpPaused Then
Ani.Visible = True
Anitime.Enabled = False
Media1.Visible = False
Else
If (UCase(Right(plyLIST.Text, 4)) = ".MP3" Or UCase(Right(plyLIST.Text, 4)) = ".WAV" Or UCase(Right(plyLIST.Text, 4)) = ".MID" Or UCase(Right(plyLIST.Text, 4)) = ".AIF" Or UCase(Right(plyLIST.Text, 5)) = ".AIFC" Or UCase(Right(plyLIST.Text, 5)) = ".AIFF" Or UCase(Right(plyLIST.Text, 4)) = ".WMA" Or UCase(Right(plyLIST.Text, 3)) = ".AU" Or UCase(Right(plyLIST.Text, 4)) = ".SND" Or UCase(Right(plyLIST.Text, 4)) = ".M3U") And Media1.PlayState = mpStopped Then
Ani.Visible = True
Anitime.Enabled = False
Media1.Visible = False
Ani.Cls
Else
Ani.Visible = False
Anitime.Enabled = False
Media1.Visible = True
End If
End If
End If

End Sub
Private Sub CheckForAudio()
If playlistfile.Caption = 1 Then
CheckForAudioPlaylist
Exit Sub
Else
End If
If (UCase(Right(mediafile.Caption, 4)) = ".MP3" Or UCase(Right(mediafile.Caption, 4)) = ".WAV" Or UCase(Right(mediafile.Caption, 4)) = ".MID" Or UCase(Right(mediafile.Caption, 4)) = ".AIF" Or UCase(Right(mediafile.Caption, 5)) = ".AIFC" Or UCase(Right(mediafile.Caption, 5)) = ".AIFF" Or UCase(Right(mediafile.Caption, 4)) = ".WMA" Or UCase(Right(mediafile.Caption, 3)) = ".AU" Or UCase(Right(mediafile.Caption, 4)) = ".SND" Or UCase(Right(mediafile.Caption, 4)) = ".M3U") And Media1.PlayState = mpPlaying Then
Ani.Visible = True
Anitime.Enabled = True
Media1.Visible = False
Else

If (UCase(Right(mediafile.Caption, 4)) = ".MP3" Or UCase(Right(mediafile.Caption, 4)) = ".WAV" Or UCase(Right(mediafile.Caption, 4)) = ".MID" Or UCase(Right(mediafile.Caption, 4)) = ".AIF" Or UCase(Right(mediafile.Caption, 5)) = ".AIFC" Or UCase(Right(mediafile.Caption, 5)) = ".AIFF" Or UCase(Right(mediafile.Caption, 4)) = ".WMA" Or UCase(Right(mediafile.Caption, 3)) = ".AU" Or UCase(Right(mediafile.Caption, 4)) = ".SND" Or UCase(Right(mediafile.Caption, 4)) = ".M3U") And Media1.PlayState = mpPaused Then
Ani.Visible = True
Anitime.Enabled = False
Media1.Visible = False
Else
If (UCase(Right(mediafile.Caption, 4)) = ".MP3" Or UCase(Right(mediafile.Caption, 4)) = ".WAV" Or UCase(Right(mediafile.Caption, 4)) = ".MID" Or UCase(Right(mediafile.Caption, 4)) = ".AIF" Or UCase(Right(mediafile.Caption, 5)) = ".AIFC" Or UCase(Right(mediafile.Caption, 5)) = ".AIFF" Or UCase(Right(mediafile.Caption, 4)) = ".WMA" Or UCase(Right(mediafile.Caption, 3)) = ".AU" Or UCase(Right(mediafile.Caption, 4)) = ".SND" Or UCase(Right(mediafile.Caption, 4)) = ".M3U") And Media1.PlayState = mpStopped Then
Ani.Visible = True
Anitime.Enabled = False
Media1.Visible = False
Ani.Cls
Else
Ani.Visible = False
Anitime.Enabled = False
If mediafile.Caption <> "NO FILE" Then Media1.Visible = True
End If
End If
End If

End Sub
Private Sub setoriginalcolour()
medPlay.Forecolor = &H80FF&
medStop.Forecolor = &H80FF&
medLoop.Forecolor = &H80FF&
medPause.Forecolor = &H80FF&
medFF.Forecolor = &H80FF&
medRW.Forecolor = &H80FF&
medLoad.Forecolor = &H80FF&
medNEXT.Forecolor = &H80FF&
medPREV.Forecolor = &H80FF&
Label1.Forecolor = &H80FF&

End Sub



Private Sub Ani_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
setoriginalcolour
End Sub

Private Sub Form_Load()
zz = SetParent(Me.hwnd, Main.hwnd)
Media1.Visible = False
Ani.Visible = False
DelLines = True ' Delete trails?
DrawWidth = 1 ' Line width
RST = 10 ' Trail size
Pts = 4 ' Number of sides
Calc (1)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
setoriginalcolour
End Sub

Private Sub Label1_Click()
On Error GoTo Error
AlwaysOnTop Main, False
'Open dialog
Media1.AutoStart = True
Dialog1.CancelError = True
Dialog1.DialogTitle = "Load audio/video playlist"
Dialog1.Filter = "Audio/Video playlists *.avp|*.avp"
Dialog1.filename = ""
Dialog1.ShowOpen
RestoreWindows
If Main.ontop.Value = 1 Then AlwaysOnTop Main, True
If Main.ontop.Value = 0 Then AlwaysOnTop Main, False
Open Dialog1.filename For Input As #1
plyLIST.Clear
Do Until EOF(1)
    Line Input #1, TempString
    plyLIST.AddItem TempString
Loop
Close #1
Media1.Visible = True
mediafile.Caption = Mid(Dialog1.filename, InStrRevVB5(Dialog1.filename, "\") + 1, Len(Dialog1.filename))
playlistfile.Caption = 1
'plyLIST.Visible = True
plyLIST.ListIndex = 0
If plyLIST.ListCount = 0 Then
MsgBox "There are no files to play!", vbCritical
Exit Sub
End If
If randomtrk.Value = 1 Then GoTo Random
If plyLIST.Text <> "" Then
Media1.filename = plyLIST.Text
Media1.Play
CheckForAudio
Exit Sub
End If
Exit Sub

Random:
Randomize
plyLIST.ListIndex = Int(Rnd * plyLIST.ListCount)
Media1.filename = plyLIST.Text
Exit Sub

Error:
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


Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
setoriginalcolour
Label1.Forecolor = &HFFFF&
End Sub

Private Sub medFF_Click()
On Error Resume Next
Media1.CurrentPosition = Media1.CurrentPosition + 5
End Sub


Private Sub medFF_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
setoriginalcolour
medFF.Forecolor = &HFFFF&
End Sub

Private Sub Media1_EndOfStream(ByVal Result As Long)
If playlistfile.Caption = 1 Then
    
On Error Resume Next
If MEDlp.Caption = 1 And Media1.PlayState = mpStopped Then GoTo LOOPME
If randomtrk.Value = 1 Then GoTo Random
If plyLIST.ListIndex < (plyLIST.ListCount - 1) Then
plyLIST.ListIndex = plyLIST.ListIndex + 1
Media1.filename = plyLIST.Text
Media1.Play
Exit Sub

LOOPME:
Media1.Play
Exit Sub

Random:
Randomize
plyLIST.ListIndex = Int(Rnd * plyLIST.ListCount)
Media1.filename = plyLIST.Text

Else
End If
End If
End Sub

Private Sub Media1_MouseMove(Button As Integer, ShiftState As Integer, X As Single, Y As Single)
setoriginalcolour
End Sub

Private Sub mediafile_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
setoriginalcolour
End Sub

Private Sub medLoad_Click()
On Error GoTo Error
AlwaysOnTop Main, False
'Open dialog
Media1.AutoStart = True
Dialog1.CancelError = True
Dialog1.DialogTitle = "Load Media"
Dialog1.Filter = "All supported files |*.dat;*.avi;*.asf;*.asx;*.wav;*.wma;*.wax;*.mpg;*.mpeg;*.m1v;*.mp2;*.mp3;*.mpa;*.mpe;*.mid;*.rmi;*.qt;*.aif;*.aifc;*.aiff;*.mov;*.au;*.snd;*.m3u|MP3 Files *.mp3|*.mp3|Avi Files *.avi|*.avi|MPEG Movies *.mpg|*.mpg|Wave Files *.wav|*.wav|Winamp Playlists *.m3u|*.m3u"
Dialog1.filename = ""
Dialog1.ShowOpen
RestoreWindows
If Main.ontop.Value = 1 Then AlwaysOnTop Main, True
If Main.ontop.Value = 0 Then AlwaysOnTop Main, False
Media1.Open Dialog1.filename
Media1.Visible = True
mediafile.Caption = Mid(Dialog1.filename, InStrRevVB5(Dialog1.filename, "\") + 1, Len(Dialog1.filename))
playlistfile.Caption = 0
'plyLIST.Visible = False
CheckForAudio
Exit Sub

Error:
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

Private Sub medLoad_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
setoriginalcolour
medLoad.Forecolor = &HFFFF&
End Sub

Private Sub medLoop_Click()

On Error GoTo Error
If MEDlp.Caption = 0 Then
MEDlp.Caption = 1
medLoop.BackColor = &H80&
Else
MEDlp.Caption = 0
medLoop.BackColor = &H404040
Exit Sub

Error:
MsgBox "Error while looping file - " & Err.Number & " : " & Err.Description
End If
End Sub

Private Sub medLoop_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
setoriginalcolour
medLoop.Forecolor = &HFFFF&
End Sub



Private Sub medPause_Click()
On Error Resume Next
Media1.Pause
End Sub

Private Sub medPause_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
setoriginalcolour
medPause.Forecolor = &HFFFF&
End Sub

Private Sub medPlay_Click()
On Error GoTo Error
If playlistfile.Caption = 1 Then GoTo Playlist
Media1.Play
medpl.Caption = 1
Exit Sub

Playlist:
If plyLIST.ListCount = 0 Then
MsgBox "There are no files to play!", vbCritical
Exit Sub
End If
If plyLIST.Text <> "" Then
If Media1.PlayState = mpPaused Then
Media1.Play
Exit Sub
Else
If randomtrk.Value = 1 Then GoTo Random
Media1.filename = plyLIST.Text
Media1.Play
Exit Sub
End If
End If
Exit Sub

Random:
Randomize
plyLIST.ListIndex = Int(Rnd * plyLIST.ListCount)
Media1.filename = plyLIST.Text
Media1.Play
Exit Sub

Error:
MsgBox "Error playing file - " & Err.Number & " : " & Err.Description
End Sub

Private Sub medPlay_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
setoriginalcolour
medPlay.Forecolor = &HFFFF&
End Sub

Private Sub medRW_Click()
On Error Resume Next
Media1.CurrentPosition = Media1.CurrentPosition - 5

End Sub

Private Sub medRW_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
setoriginalcolour
medRW.Forecolor = &HFFFF&
End Sub

Private Sub medStop_Click()
On Error GoTo Error
If mediafile.Caption <> "NO FILE" Then
Media1.Stop
Media1.CurrentPosition = 0
medpl.Caption = 0
Exit Sub

Error:
MsgBox "Error stopping file - " & Err.Number & " : " & Err.Description
Else
End If
End Sub

Private Sub medStop_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
setoriginalcolour
medStop.Forecolor = &HFFFF&
End Sub

Private Sub medTIME_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
setoriginalcolour
End Sub

Private Sub medNEXT_Click()
On Error Resume Next
If playlistfile.Caption = 0 Then
Media1.Previous
Exit Sub
Else
If randomtrk.Value = 1 Then GoTo Random
plyLIST.ListIndex = plyLIST.ListIndex + 1
Media1.filename = plyLIST.Text
Exit Sub
Random:
Randomize
plyLIST.ListIndex = Int(Rnd * plyLIST.ListCount)
Media1.filename = plyLIST.Text
End If
End Sub

Private Sub medNEXT_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
setoriginalcolour
medNEXT.Forecolor = &HFFFF&
End Sub

Private Sub medPREV_Click()
On Error Resume Next
If playlistfile.Caption = 0 Then
Media1.Next
Exit Sub
Else
If randomtrk.Value = 1 Then GoTo Random
plyLIST.ListIndex = plyLIST.ListIndex - 1
Media1.filename = plyLIST.Text
Exit Sub
Random:
Randomize
plyLIST.ListIndex = Int(Rnd * plyLIST.ListCount)
Media1.filename = plyLIST.Text
End If
End Sub

Private Sub medPREV_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
setoriginalcolour
medPREV.Forecolor = &HFFFF&
End Sub

Private Sub plyLIST_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
setoriginalcolour
End Sub

Private Sub Timer1_Timer()
'If Me.Visible = True Then  'to stop the focus being grabbed by the procedure.
CheckForAudio
'Else
'End If
Dim tmplength As Long
'Timer
On Error Resume Next
If Media1.CurrentPosition > 0 Then
tmplength = Int(Media1.CurrentPosition) ' this is to stop the stange 'OVERFLOW' message when changing tracks on a m3u file.
medTIME.Caption = TimeSerial(0, 0, tmplength)

On Error Resume Next
Call PercentBar(posper, Media1.CurrentPosition, Media1.Duration)
Else

End If
If MEDlp.Caption = 1 And Media1.PlayState = mpStopped And medpl.Caption = 1 Then
Media1.Play
End If

'If playing, make button red
If Media1.PlayState = mpPlaying Then
medPlay.BackColor = &H80&
Else
medPlay.BackColor = &H404040
End If

'If paused, make button red
If Media1.PlayState = mpPaused Then
medPause.BackColor = &H80&
Else
medPause.BackColor = &H404040
End If

'If stopped, make button red
If Media1.PlayState = mpStopped Then
medStop.BackColor = &H80&
Else
medStop.BackColor = &H404040
End If

If playlistfile.Caption = 1 And Dialog1.filename = "" Then mediafile.Caption = "TRANSFERED PLAYLIST - " & Mid(plyLIST.Text, InStrRevVB5(plyLIST.Text, "\") + 1, Len(plyLIST.Text)) & " (" & plyLIST.ListIndex + 1 & " of " & plyLIST.ListCount & ")"
If playlistfile.Caption = 1 And Dialog1.filename <> "" Then mediafile.Caption = Mid(Dialog1.filename, InStrRevVB5(Dialog1.filename, "\") + 1, Len(Dialog1.filename)) & " - " & Mid(plyLIST.Text, InStrRevVB5(plyLIST.Text, "\") + 1, Len(plyLIST.Text)) & " (" & plyLIST.ListIndex + 1 & " of " & plyLIST.ListCount & ")"
End Sub

Private Sub Anitime_Timer()
'Draw the mistify polygons
    Calc (0)
End Sub

Sub Dpoly(P() As Integer, C As Long)
    Dim I As Integer
'Draw the lines that delimit the polygon
    For I = 0 To (Pts - 2) * 2 Step 2
        Ani.Line (P(I), P(I + 1))-(P(I + 2), P(I + 3)), C
    Next I
    Ani.Line (P(I), P(I + 1))-(P(0), P(1)), C
End Sub
Sub Calc(Op As Integer)
'********************************************************
'
'This is where all happens
'
'********************************************************
'Declare local static variabled required for the program
    Static X(7) As Integer, Y(7) As Integer
    Static T1(7) As Integer, T2(7) As Integer
    Static Maxx As Long, Maxy As Long
    Static C(3) As Integer, CP(3) As Integer
'Get the Forms Size and take 5 twips
    Maxx = Ani.Width - 5
    Maxy = Ani.Height - 5
    
'1st time. Set the variables values
    If Op = 1 Then
'Start the random number generator
        Randomize
'Update the "coordinate points" variables
        T1(0) = Int(Rnd * Maxx)
        T1(1) = Int(Rnd * Maxy)
        T1(2) = Int(Rnd * Maxx)
        T1(3) = Int(Rnd * Maxy)
        T1(4) = Int(Rnd * Maxx)
        T1(5) = Int(Rnd * Maxy)
        T1(6) = Int(Rnd * Maxx)
        T1(7) = Int(Rnd * Maxy)
'Update the "delocation factor" variables
        X(0) = 60
        Y(0) = 80
        X(1) = 80
        Y(1) = 100
        X(2) = 100
        Y(2) = 60
        X(3) = 90
        Y(3) = 90
'Update the "gradient color" variables
        C(1) = Int(Rnd * 256)
        C(2) = Int(Rnd * 256)
        C(3) = Int(Rnd * 256)
'Update the "gradient factor" variables
        CP(1) = 2
        CP(2) = 3
        CP(3) = 4
'Do the next code
        Op = 2
    End If
    If Op = 2 Then
'Update the "points coordinates" variables
        T2(0) = T1(0)
        T2(1) = T1(1)
        T2(2) = T1(2)
        T2(3) = T1(3)
        T2(4) = T1(4)
        T2(5) = T1(5)
        T2(6) = T1(6)
        T2(7) = T1(7)
        X(4) = X(0)
        Y(4) = Y(0)
        X(5) = X(1)
        Y(5) = Y(1)
        X(6) = X(2)
        Y(6) = Y(2)
        X(7) = X(3)
        Y(7) = Y(3)
'Set the "left trails" variable to 0
        A = 0
'Clear the screen
        Ani.Cls
    End If
'Check if the "gradient factor" variables are between 0
'and 255 and update the "gradient color" variables
    For I = 1 To 3
        If C(I) + CP(I) < 0 Or C(I) + CP(I) > 255 Then CP(I) = -CP(I)
        C(I) = C(I) + CP(I)
    Next I
'Draw the Polygon
    Dpoly T1, RGB(C(1), C(2), C(3))
'Check if the points have colided with the border of the
'screen. If so update the "delocation factor" variables
    If T1(0) < 1 Or T1(0) > Maxx Then X(0) = -X(0)
    If T1(2) < 1 Or T1(2) > Maxx Then X(1) = -X(1)
    If T1(4) < 1 Or T1(4) > Maxx Then X(2) = -X(2)
    If T1(6) < 1 Or T1(6) > Maxx Then X(3) = -X(3)
    If T1(1) < 1 Or T1(1) > Maxy Then Y(0) = -Y(0)
    If T1(3) < 1 Or T1(3) > Maxy Then Y(1) = -Y(1)
    If T1(5) < 1 Or T1(5) > Maxy Then Y(2) = -Y(2)
    If T1(7) < 1 Or T1(7) > Maxy Then Y(3) = -Y(3)
'Check if the "delete trails" variable is true
    If DelLines = True Then
'If so, check if the "lefted trails" variable is bigger
'than the "trails size" variable
        If A > RST Then
'If so, delete the old trails, by drawing a black polygon
'over them
            Dpoly T2, 0
'Update the 2nd poly "coordinate points" variables as
'with the 1st
            If T2(0) < 1 Or T2(0) > Maxx Then X(4) = -X(4)
            If T2(2) < 1 Or T2(2) > Maxx Then X(5) = -X(5)
            If T2(4) < 1 Or T2(4) > Maxx Then X(6) = -X(6)
            If T2(6) < 1 Or T2(6) > Maxx Then X(7) = -X(7)
            If T2(1) < 1 Or T2(1) > Maxy Then Y(4) = -Y(4)
            If T2(3) < 1 Or T2(3) > Maxy Then Y(5) = -Y(5)
            If T2(5) < 1 Or T2(5) > Maxy Then Y(6) = -Y(6)
            If T2(7) < 1 Or T2(7) > Maxy Then Y(7) = -Y(7)
            T2(0) = T2(0) + X(4)
            T2(1) = T2(1) + Y(4)
            T2(2) = T2(2) + X(5)
            T2(3) = T2(3) + Y(5)
            T2(4) = T2(4) + X(6)
            T2(5) = T2(5) + Y(6)
            T2(6) = T2(6) + X(7)
            T2(7) = T2(7) + Y(7)
        Else
'Update the "left trails" variable
            A = A + 1
        End If
    End If
'Update the 1st polygon "coordinate points" variables
    T1(0) = T1(0) + X(0)
    T1(1) = T1(1) + Y(0)
    T1(2) = T1(2) + X(1)
    T1(3) = T1(3) + Y(1)
    T1(4) = T1(4) + X(2)
    T1(5) = T1(5) + Y(2)
    T1(6) = T1(6) + X(3)
    T1(7) = T1(7) + Y(3)
'Do the cached events
    DoEvents
End Sub
