VERSION 5.00
Begin VB.Form frmAlarm 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Alarm Notification"
   ClientHeight    =   2115
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer2 
      Interval        =   5
      Left            =   60
      Top             =   1260
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   60
      Top             =   840
   End
   Begin VB.Label alarmtime 
      Caption         =   "Label2"
      Height          =   195
      Left            =   2880
      TabIndex        =   3
      Top             =   1260
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label OKBUT 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "OK"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Federation"
         Size            =   18
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   390
      Left            =   1380
      TabIndex        =   2
      Top             =   1500
      Width           =   1710
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Alarm for "
      BeginProperty Font 
         Name            =   "Federation"
         Size            =   12
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   270
      Left            =   120
      TabIndex        =   1
      Top             =   900
      Width           =   4410
   End
   Begin VB.Label Alarmnotif 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Alarm Notification"
      BeginProperty Font 
         Name            =   "Federation"
         Size            =   18
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   450
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   4410
   End
End
Attribute VB_Name = "frmAlarm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub setoriginalcolour()
OKBUT.Forecolor = &H80FF&
End Sub

Private Sub Alarmnotif_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
setoriginalcolour
End Sub

Private Sub Form_Load()
Label1.Caption = "Alarm for " & Format(frmAlarmConfig.AlarmCon.Caption, "HH:MM:SS AMPM")
alarmtime.Caption = 0
'Main.Check1.Value = 0
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
setoriginalcolour
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
setoriginalcolour
End Sub

Private Sub OKBUT_Click()
Unload Me
End Sub

Private Sub OKBUT_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
OKBUT.Forecolor = &HFFFF&
End Sub

Private Sub Timer1_Timer()
If alarmtime.Caption > 3 Then Alarmnotif.Forecolor = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
End Sub

Private Sub Timer2_Timer()

If alarmtime.Caption > 3 Then
'Main.Check1.Value = 1
OKBUT.Enabled = True
Exit Sub
Else
'sound
Do Until alarmtime.Caption > 3
Playsound "Alarm.wav"
alarmtime.Caption = alarmtime.Caption + 1
Loop
End If
End Sub
