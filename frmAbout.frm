VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About"
   ClientHeight    =   4635
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4155
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   4155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   4485
      Left            =   60
      Picture         =   "frmAbout.frx":0000
      ScaleHeight     =   4455
      ScaleWidth      =   3990
      TabIndex        =   0
      Top             =   60
      Width           =   4020
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Additional credit"
         BeginProperty Font 
            Name            =   "Federation"
            Size            =   6
            Charset         =   0
            Weight          =   800
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   1155
         Left            =   120
         TabIndex        =   6
         Top             =   2700
         Width           =   3795
      End
      Begin VB.Label lExit 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "OK"
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
         Height          =   450
         Left            =   60
         TabIndex        =   5
         ToolTipText     =   "NICE!"
         Top             =   4020
         Width           =   3810
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Special thanks to Jan-Alexander Mock for moveable form and RJ Soft for the transparant form and systray stuff."
         BeginProperty Font 
            Name            =   "Federation"
            Size            =   6
            Charset         =   0
            Weight          =   800
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   495
         Left            =   60
         TabIndex        =   4
         Top             =   2160
         Width           =   3735
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   $"frmAbout.frx":2618A
         BeginProperty Font 
            Name            =   "Federation"
            Size            =   8.25
            Charset         =   0
            Weight          =   800
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   1035
         Left            =   60
         TabIndex        =   3
         Top             =   1020
         Width           =   3795
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "Federation"
            Size            =   9.75
            Charset         =   0
            Weight          =   800
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   420
         TabIndex        =   2
         Top             =   480
         Width           =   3015
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Console"
         BeginProperty Font 
            Name            =   "Federation"
            Size            =   15.75
            Charset         =   0
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   360
         Left            =   1140
         TabIndex        =   1
         Top             =   120
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub setoriginalcolour()
lExit.Forecolor = &H80FF&
End Sub
Private Sub Form_Load()
Label2.Caption = "V" & App.Major & "." & App.Minor & App.Revision
Label5.Caption = "Additional credit to :" & Chr(13) & Chr(10) & "Simon Gill for the idea of a visual alarm." & Chr(13) & Chr(10) & "Andrew Cartwright for the idea if a playlist editor." & Chr(13) & Chr(10) & "Lastly, all the other code such as listbox tooltips etc. Thanks!"
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
setoriginalcolour
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
setoriginalcolour
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
setoriginalcolour
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
setoriginalcolour
End Sub

Private Sub lExit_Click()
Unload Me
End Sub

Private Sub lExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lExit.Forecolor = &HFFFF&
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
setoriginalcolour
End Sub
