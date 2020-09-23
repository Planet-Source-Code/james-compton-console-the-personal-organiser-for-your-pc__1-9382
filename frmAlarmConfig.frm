VERSION 5.00
Begin VB.Form frmAlarmConfig 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Console-Alarm Config"
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
   Begin VB.CheckBox AlarmShut 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Shut down computer"
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
      Height          =   495
      Left            =   180
      TabIndex        =   8
      Top             =   3255
      Width           =   4155
   End
   Begin VB.CheckBox AlarmOrig 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Display notification and play sound"
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
      Height          =   495
      Left            =   180
      TabIndex        =   6
      Top             =   2535
      Width           =   4155
   End
   Begin VB.TextBox AlarmHOUR 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Federation"
         Size            =   14.25
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   315
      Left            =   1080
      MaxLength       =   2
      TabIndex        =   3
      Text            =   "12"
      ToolTipText     =   "Alarm Hours"
      Top             =   1155
      Width           =   735
   End
   Begin VB.TextBox AlarmMIN 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Federation"
         Size            =   14.25
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   315
      Left            =   1860
      MaxLength       =   2
      TabIndex        =   2
      Text            =   "00"
      ToolTipText     =   "Alarm Minutes"
      Top             =   1155
      Width           =   795
   End
   Begin VB.TextBox AMPM 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Federation"
         Size            =   14.25
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   315
      Left            =   2460
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   1
      Text            =   "AM"
      ToolTipText     =   "Alarm AM or PM"
      Top             =   1155
      Width           =   975
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Alarm On/Off"
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
      Height          =   195
      Left            =   1200
      TabIndex        =   0
      ToolTipText     =   "Alarm on or off?"
      Top             =   360
      Width           =   2175
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   120
      Top             =   180
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00808080&
      Height          =   300
      Left            =   1395
      Top             =   1530
      Visible         =   0   'False
      Width           =   1890
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "When alarm occurs:"
      BeginProperty Font 
         Name            =   "Federation"
         Size            =   12
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   315
      Left            =   240
      TabIndex        =   7
      Top             =   1935
      Width           =   4155
   End
   Begin VB.Label AlarmCon 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00:00"
      BeginProperty Font 
         Name            =   "Federation"
         Size            =   12
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   1440
      TabIndex        =   5
      Top             =   1530
      Visible         =   0   'False
      Width           =   1785
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Federation"
         Size            =   14.25
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   315
      Left            =   1800
      TabIndex        =   4
      Top             =   1095
      Width           =   135
   End
End
Attribute VB_Name = "frmAlarmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub RestoreWindows()
'Restore hidden windows:
Select Case Main.WhichButton.Caption
'Notepad
Case 1
frmNotes.Show

'Programs
Case 2
frmProgs.Show

'Memo
Case 3
frmMemo.Show

'Favs
Case 4
frmFavs.Show

'Media
Case 5
frmMedia.Show

'Alarm Config
Case 6
frmAlarmConfig.Show

End Select
End Sub
Private Sub CHECK_ALARM()
If AlarmOrig.Value = 1 Then
Main.Show
AlwaysOnTop Main, True ' bring it to the front
AlwaysOnTop Main, False
frmAlarm.Show vbModal
If Main.ontop.Value = 1 Then AlwaysOnTop Main, True
If Main.ontop.Value = 0 Then AlwaysOnTop Main, False
RestoreWindows

Exit Sub
Else
If AlarmShut.Value = 1 Then ShutDownWindows
Exit Sub
End If
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

Private Sub AlarmCon_Change()
Call ProgPath
On Error GoTo Error
Open ProgPath & "Alarm.dat" For Output As #1
Print #1, AlarmHOUR.Text
Print #1, AlarmMIN.Text
Print #1, AMPM.Text
Print #1, Check1.Value
Print #1, AlarmOrig.Value
Print #1, AlarmShut.Value
Close #1
Exit Sub

Error:
MsgBox "Error saving alarm data - " & Err.Number & " : " & Err.Description

End Sub

Private Sub AlarmHOUR_LostFocus()
isit = IsNumeric(AlarmHOUR.Text)
If isit = False Then
MsgBox "Error - Enter a numeric value between 1 and 12.", vbExclamation, "Error setting alarm"
AlarmHOUR.SetFocus
Exit Sub
Else
If AlarmHOUR.Text < 1 Then
MsgBox "Error - Enter a numeric value between 1 and 12.", vbExclamation, "Error setting alarm"
AlarmHOUR.SetFocus
Exit Sub
Else
If AlarmHOUR.Text > 12 Then
MsgBox "Error - Enter a numeric value between 1 and 12.", vbExclamation, "Error setting alarm"
AlarmHOUR.SetFocus
Exit Sub
Else
End If
End If
End If
'AlarmHOUR.Text = Int(AlarmHOUR.Text)
End Sub

Private Sub AlarmMIN_LostFocus()
isit = IsNumeric(AlarmMIN.Text)
If isit = False Then
MsgBox "Error - Enter a numeric value between 0 and 59.", vbExclamation, "Error setting alarm"
AlarmMIN.SetFocus
Exit Sub
Else
If AlarmMIN.Text < 0 Then
MsgBox "Error - Enter a numeric value between 0 and 59.", vbExclamation, "Error setting alarm"
AlarmMIN.SetFocus
Exit Sub
Else
If AlarmMIN.Text > 59 Then
MsgBox "Error - Enter a numeric value between 0 and 59.", vbExclamation, "Error setting alarm"
AlarmMIN.SetFocus
Exit Sub
Else
End If
End If
End If
'AlarmMIN.Text = Int(AlarmMIN.Text)
End Sub

Private Sub alarmorig_Click()
If AlarmOrig.Value = 1 Then AlarmShut.Value = 0
Call ProgPath
On Error GoTo Error
Open ProgPath & "Alarm.dat" For Output As #1
Print #1, AlarmHOUR.Text
Print #1, AlarmMIN.Text
Print #1, AMPM.Text
Print #1, Check1.Value
Print #1, AlarmOrig.Value
Print #1, AlarmShut.Value
Close #1
Exit Sub

Error:
MsgBox "Error saving alarm data - " & Err.Number & " : " & Err.Description

End Sub

Private Sub alarmshut_Click()
If AlarmShut.Value = 1 Then AlarmOrig.Value = 0
Call ProgPath
On Error GoTo Error
Open ProgPath & "Alarm.dat" For Output As #1
Print #1, AlarmHOUR.Text
Print #1, AlarmMIN.Text
Print #1, AMPM.Text
Print #1, Check1.Value
Print #1, AlarmOrig.Value
Print #1, AlarmShut.Value
Close #1
Exit Sub

Error:
MsgBox "Error saving alarm data - " & Err.Number & " : " & Err.Description

End Sub

Private Sub AMPM_Click()
If AMPM.Text = "AM" Then
AMPM.Text = "PM"
Else
AMPM.Text = "AM"
End If

Call ProgPath
On Error GoTo Error
Open ProgPath & "Alarm.dat" For Output As #1
Print #1, AlarmHOUR.Text
Print #1, AlarmMIN.Text
Print #1, AMPM.Text
Print #1, Check1.Value
Print #1, AlarmOrig.Value
Print #1, AlarmShut.Value
Close #1
Exit Sub

Error:
MsgBox "Error saving alarm data - " & Err.Number & " : " & Err.Description

End Sub

Private Sub Check1_Click()
Call ProgPath
On Error GoTo Error
Open ProgPath & "Alarm.dat" For Output As #1
Print #1, AlarmHOUR.Text
Print #1, AlarmMIN.Text
Print #1, AMPM.Text
Print #1, Check1.Value
Print #1, AlarmOrig.Value
Print #1, AlarmShut.Value
Close #1
Exit Sub

Error:
MsgBox "Error saving alarm data - " & Err.Number & " : " & Err.Description

End Sub



Private Sub Form_Load()
zz = SetParent(Me.hwnd, Main.hwnd)
End Sub

Private Sub Timer1_Timer()
Dim CONVERTEDTIME As String
Dim alarmtime As Integer

'Alarm
'On Error Resume Next
'Hours
On Error Resume Next
        AMPM.Text = UCase(AMPM.Text)
        If AMPM.Text = "AM" And AlarmHOUR.Text <> "12" Then CONVERTEDTIME = AlarmHOUR.Text
        If AMPM.Text = "AM" And AlarmHOUR.Text = "12" Then CONVERTEDTIME = 0
        If AMPM.Text = "PM" And AlarmHOUR.Text <> "12" Then CONVERTEDTIME = AlarmHOUR.Text + 12
        If AMPM.Text = "PM" And AlarmHOUR.Text = "12" Then CONVERTEDTIME = AlarmHOUR.Text

        'Mins
        CONVERTEDTIME = CONVERTEDTIME & ":" & AlarmMIN.Text & ":00"
        AlarmCon.Caption = Format(CONVERTEDTIME, "HH:MM:SS")

        'Check alarm
        If Check1.Value = 1 And AlarmCon.Caption = Time Then CHECK_ALARM
End Sub
