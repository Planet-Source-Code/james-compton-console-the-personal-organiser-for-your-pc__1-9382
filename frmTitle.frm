VERSION 5.00
Begin VB.Form frmTitle 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Title"
   ClientHeight    =   3615
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmTitle.frx":0000
   ScaleHeight     =   3615
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   675
      Top             =   3015
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Loading..."
      BeginProperty Font 
         Name            =   "Federation"
         Size            =   15.75
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   45
      TabIndex        =   3
      ToolTipText     =   "Click and drag to move window"
      Top             =   3060
      Width           =   5310
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "(c) James Compton 2000"
      BeginProperty Font 
         Name            =   "Federation"
         Size            =   15.75
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   735
      Left            =   45
      TabIndex        =   2
      ToolTipText     =   "Click and drag to move window"
      Top             =   1845
      Width           =   5310
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "Federation"
         Size            =   15.75
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   330
      Left            =   315
      TabIndex        =   1
      ToolTipText     =   "Click and drag to move window"
      Top             =   945
      Width           =   4635
   End
   Begin VB.Label MainCaption 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Console"
      BeginProperty Font 
         Name            =   "Federation"
         Size            =   26.25
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   600
      Left            =   315
      TabIndex        =   0
      ToolTipText     =   "Click and drag to move window"
      Top             =   360
      Width           =   4635
   End
End
Attribute VB_Name = "frmTitle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Label1.Caption = "Version " & App.Major & "." & App.Minor & App.Revision
End Sub

Private Sub Timer1_Timer()
If Timer1.Interval = 1000 Then
Main.Show
Unload Me
Else
End If
End Sub
