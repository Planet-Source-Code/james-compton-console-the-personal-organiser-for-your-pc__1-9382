VERSION 5.00
Begin VB.Form Main 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Console"
   ClientHeight    =   6045
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7920
   FillColor       =   &H00FFFFFF&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":030A
   ScaleHeight     =   6045
   ScaleWidth      =   7920
   Tag             =   "Console - By J Compton"
   Begin VB.PictureBox Blank 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   4275
      Left            =   2805
      ScaleHeight     =   4245
      ScaleWidth      =   4605
      TabIndex        =   18
      Top             =   1140
      Visible         =   0   'False
      Width           =   4635
   End
   Begin VB.OptionButton searchtype 
      BackColor       =   &H00000000&
      Caption         =   "Yahoo!"
      BeginProperty Font 
         Name            =   "Federation"
         Size            =   8.25
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   240
      Index           =   0
      Left            =   2880
      TabIndex        =   39
      Top             =   4230
      Width           =   1590
   End
   Begin VB.OptionButton searchtype 
      BackColor       =   &H00000000&
      Caption         =   "Crack search"
      BeginProperty Font 
         Name            =   "Federation"
         Size            =   8.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   240
      Index           =   1
      Left            =   2880
      TabIndex        =   38
      Top             =   4590
      Width           =   2085
   End
   Begin VB.TextBox netSearch 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
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
      Height          =   240
      Left            =   2880
      TabIndex        =   37
      Top             =   3825
      Width           =   2265
   End
   Begin VB.CheckBox snd 
      Caption         =   "Check1"
      Height          =   195
      Left            =   2970
      TabIndex        =   36
      ToolTipText     =   "Mute sound from media player?"
      Top             =   5580
      Width           =   195
   End
   Begin VB.HScrollBar VolumeCtrl 
      Height          =   195
      LargeChange     =   40
      Left            =   3150
      Max             =   5000
      SmallChange     =   10
      TabIndex        =   33
      Top             =   5580
      Value           =   2350
      Width           =   3525
   End
   Begin VB.CheckBox ontop 
      Caption         =   "Ontop?"
      Height          =   195
      Left            =   780
      TabIndex        =   30
      Top             =   1380
      Width           =   195
   End
   Begin VB.Timer AlarmTimer 
      Interval        =   10
      Left            =   60
      Top             =   1035
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   510
      Left            =   2685
      Picture         =   "Form1.frx":9C0F4
      ScaleHeight     =   510
      ScaleWidth      =   5010
      TabIndex        =   0
      ToolTipText     =   "Click and drag to move window"
      Top             =   225
      Width           =   5010
      Begin VB.Label MainCaption 
         BackStyle       =   0  'Transparent
         Caption         =   "Console - Command"
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
         Left            =   165
         TabIndex        =   1
         ToolTipText     =   "Click and drag to move window"
         Top             =   60
         Width           =   4635
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   3375
      Left            =   2340
      ScaleHeight     =   3345
      ScaleWidth      =   225
      TabIndex        =   16
      Top             =   1560
      Width           =   255
      Begin VB.Label comCom 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "C O M M A N D"
         BeginProperty Font 
            Name            =   "Federation"
            Size            =   8.25
            Charset         =   0
            Weight          =   800
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   2355
         Left            =   0
         TabIndex        =   17
         Top             =   540
         Width           =   255
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Shape Shape9 
      BorderColor     =   &H00808080&
      Height          =   330
      Left            =   2835
      Top             =   3780
      Width           =   2355
   End
   Begin VB.Label Escaped 
      Height          =   255
      Left            =   360
      TabIndex        =   42
      Top             =   5265
      Visible         =   0   'False
      Width           =   2265
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Net Search"
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
      Height          =   330
      Left            =   2880
      TabIndex        =   41
      Top             =   3375
      Width           =   2250
   End
   Begin VB.Label Gosearch 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "GO"
      Enabled         =   0   'False
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
      Height          =   285
      Left            =   3105
      TabIndex        =   40
      Top             =   4995
      Width           =   1710
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   2880
      X2              =   5265
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Shape Shape6 
      FillStyle       =   0  'Solid
      Height          =   285
      Left            =   2790
      Shape           =   3  'Circle
      Top             =   5490
      Width           =   315
   End
   Begin VB.Label comVolume 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Volume"
      BeginProperty Font 
         Name            =   "Federation"
         Size            =   11.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   285
      Left            =   5355
      TabIndex        =   35
      Top             =   2610
      Width           =   2010
   End
   Begin VB.Label Vol 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "94"
      BeginProperty Font 
         Name            =   "Federation"
         Size            =   6
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   6705
      TabIndex        =   34
      Top             =   5580
      Width           =   600
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "?"
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
      Height          =   330
      Left            =   495
      TabIndex        =   32
      ToolTipText     =   "Help"
      Top             =   5580
      Width           =   390
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "On top?"
      BeginProperty Font 
         Name            =   "Federation"
         Size            =   8.25
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1020
      TabIndex        =   31
      Top             =   1380
      Width           =   975
   End
   Begin VB.Label comCD 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "CD Player"
      BeginProperty Font 
         Name            =   "Federation"
         Size            =   11.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   285
      Left            =   5355
      TabIndex        =   29
      Top             =   4995
      Width           =   2010
   End
   Begin VB.Label comCALC 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Calculator"
      BeginProperty Font 
         Name            =   "Federation"
         Size            =   11.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   285
      Left            =   5355
      TabIndex        =   28
      Top             =   4590
      Width           =   2010
   End
   Begin VB.Label mnuPLAY 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Playlist"
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
      Height          =   345
      Left            =   480
      TabIndex        =   27
      Top             =   3780
      Width           =   1710
   End
   Begin VB.Label Alarmbut 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Alarm"
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
      Left            =   480
      TabIndex        =   26
      Top             =   4260
      Width           =   1710
   End
   Begin VB.Shape Shape3 
      FillStyle       =   0  'Solid
      Height          =   435
      Left            =   2340
      Shape           =   3  'Circle
      Top             =   4740
      Width           =   255
   End
   Begin VB.Shape Shape2 
      FillStyle       =   0  'Solid
      Height          =   435
      Left            =   2340
      Shape           =   3  'Circle
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label aboutshow 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "?"
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
      Height          =   330
      Left            =   180
      TabIndex        =   25
      ToolTipText     =   "About"
      Top             =   5580
      Width           =   390
   End
   Begin VB.Label MemoMess 
      BackStyle       =   0  'Transparent
      Caption         =   "Messages"
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
      Height          =   1215
      Left            =   2880
      TabIndex        =   24
      Top             =   1800
      Width           =   2235
   End
   Begin VB.Label Timelbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Time"
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
      Left            =   5340
      TabIndex        =   23
      Top             =   780
      Width           =   1875
   End
   Begin VB.Label Datelbl 
      BackColor       =   &H00000000&
      Caption         =   "Date"
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
      Left            =   3000
      TabIndex        =   22
      Top             =   780
      Width           =   2235
   End
   Begin VB.Label ver1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "Federation"
         Size            =   8.25
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   480
      TabIndex        =   21
      Top             =   900
      Width           =   1695
   End
   Begin VB.Label Min 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "X"
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
      Height          =   330
      Left            =   7380
      TabIndex        =   20
      ToolTipText     =   "Exit program"
      Top             =   5580
      Width           =   390
   End
   Begin VB.Label maintxt 
      BackStyle       =   0  'Transparent
      Caption         =   "Main Text"
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
      Height          =   510
      Left            =   2835
      TabIndex        =   19
      Top             =   1320
      Width           =   2355
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   5280
      X2              =   5280
      Y1              =   1200
      Y2              =   5400
   End
   Begin VB.Label comDosEdit 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Dos Editor"
      BeginProperty Font 
         Name            =   "Federation"
         Size            =   11.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   285
      Left            =   5355
      TabIndex        =   15
      Top             =   4005
      Width           =   2010
   End
   Begin VB.Label comExplore 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Explorer"
      BeginProperty Font 
         Name            =   "Federation"
         Size            =   11.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   285
      Left            =   5355
      TabIndex        =   14
      Top             =   3600
      Width           =   2010
   End
   Begin VB.Label comControl 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Control Pnl."
      BeginProperty Font 
         Name            =   "Federation"
         Size            =   11.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   285
      Left            =   5355
      TabIndex        =   13
      Top             =   3195
      Width           =   2010
   End
   Begin VB.Label comIP 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "IP Address"
      BeginProperty Font 
         Name            =   "Federation"
         Size            =   11.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   285
      Left            =   5355
      TabIndex        =   12
      Top             =   2205
      Width           =   2010
   End
   Begin VB.Label comDOS 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Dos Prompt"
      BeginProperty Font 
         Name            =   "Federation"
         Size            =   11.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   285
      Left            =   5355
      TabIndex        =   11
      Top             =   1770
      Width           =   2010
   End
   Begin VB.Label comREG 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Registry"
      BeginProperty Font 
         Name            =   "Federation"
         Size            =   11.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   285
      Left            =   5355
      TabIndex        =   10
      Top             =   1350
      Width           =   2010
   End
   Begin VB.Label mnuMedia 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Media"
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
      Left            =   480
      TabIndex        =   9
      Top             =   3360
      Width           =   1710
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "By Fliblo - (c) 2000 J Compton"
      BeginProperty Font 
         Name            =   "Federation"
         Size            =   9
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   210
      Left            =   750
      TabIndex        =   8
      Top             =   5805
      Width           =   6615
   End
   Begin VB.Label mnuFavs 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Favs"
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
      Left            =   480
      TabIndex        =   7
      Top             =   2940
      Width           =   1710
   End
   Begin VB.Label lExit 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Hide"
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
      Left            =   480
      TabIndex        =   6
      ToolTipText     =   "Hide the program in the systray"
      Top             =   4680
      Width           =   1710
   End
   Begin VB.Label mnuMemo 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Memo"
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
      Left            =   480
      TabIndex        =   5
      Top             =   2100
      Width           =   1710
   End
   Begin VB.Label mnuProgs 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Progs"
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
      Left            =   480
      TabIndex        =   4
      Top             =   2520
      Width           =   1710
   End
   Begin VB.Label WhichButton 
      Caption         =   "0"
      Height          =   195
      Left            =   840
      TabIndex        =   3
      Top             =   5580
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label mnuNotes 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Notes"
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
      Left            =   480
      TabIndex        =   2
      Top             =   1680
      Width           =   1710
   End
   Begin VB.Shape Shape4 
      FillStyle       =   0  'Solid
      Height          =   195
      Left            =   2760
      Shape           =   3  'Circle
      Top             =   780
      Width           =   315
   End
   Begin VB.Shape Shape5 
      FillStyle       =   0  'Solid
      Height          =   195
      Left            =   7140
      Shape           =   3  'Circle
      Top             =   780
      Width           =   315
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   195
      Left            =   2940
      Top             =   780
      Width           =   4335
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   240
      Left            =   6660
      Top             =   5535
      Width           =   555
   End
   Begin VB.Shape Shape7 
      FillStyle       =   0  'Solid
      Height          =   240
      Left            =   7020
      Shape           =   3  'Circle
      Top             =   5535
      Width           =   315
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'+---------------------------------------------------------------------------+
'|FakeForm                                                                   |
'|---------------------------------------------------------------------------+
'|Create an bordered Form from an unbordered Form (Borderstyle = 0)          |
'|The Border has the Same color like the Form's Background                   |
'|It also creates a Titlebar with other Text than it is showen in the Taskbar|
'|The Titlebar colors are mixed                                              |
'+---------------------------------------------------------------------------+
'|By: Jan-Alexander Mock                                                     |
'|HP: http://www.jan-alexander.de or http://www.janalexander.de              |
'|EMail: arnold72@hotmail.com                                                |
'+---------------------------------------------------------------------------+
'|Please give me some Credits when u use this code!                          |
'+---------------------------------------------------------------------------+

Private Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
'Api Functions Declaration
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Selected As Integer
Dim Focused As Boolean
Private Function ProgPath()
' If dragged file is in the root, append filename.
If Mid(App.Path, Len(App.Path)) = "\" Then
ProgPath = App.Path
' If dragged file is not in root, append "\" and filename.
Else
ProgPath = App.Path & "\"
End If
End Function
Private Sub setoriginalcolour()
mnuNotes.Forecolor = &H80FF&
mnuProgs.Forecolor = &H80FF&
mnuMemo.Forecolor = &H80FF&
lExit.Forecolor = &H80FF&
mnuFavs.Forecolor = &H80FF&
mnuMedia.Forecolor = &H80FF&
Alarmbut.Forecolor = &H80FF&
mnuPLAY.Forecolor = &H80FF&
comCD.Forecolor = &H80FF&
comCALC.Forecolor = &H80FF&
Label3.Forecolor = &H80FF&
comCom.Forecolor = &H80FF&
aboutshow.Forecolor = &H808080
comREG.Forecolor = &H80FF&
comIP.Forecolor = &H80FF&
comDOS.Forecolor = &H80FF&
comControl.Forecolor = &H80FF&
comExplore.Forecolor = &H80FF&
comDosEdit.Forecolor = &H80FF&
comVolume.Forecolor = &H80FF&
Gosearch.Forecolor = &H80FF&
Min.Forecolor = &H808080
End Sub
Private Sub SetTrans()
'Free the memory set
If hRgn Then DeleteObject hRgn

'Scan the Bitmap and remove all transparent pixels from it, creating a new region
hRgn = GetBitmapRegion(Me.Picture, vbWhite)

'Set the Forms new Region
SetWindowRgn Me.hwnd, hRgn, True

End Sub



Private Sub btnClose_GotFocus()
Focused = True
End Sub

Private Sub btnClose_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Focused = True Then
  Picture1.SetFocus
  Focused = False
End If
End Sub

Private Sub aboutshow_Click()
AlwaysOnTop Me, False 'allow aboutbox to be ontop of console
frmAbout.Show 1
RestoreWindows
If Main.ontop.Value = 1 Then AlwaysOnTop Me, True
If Main.ontop.Value = 0 Then AlwaysOnTop Me, False
End Sub

Private Sub aboutshow_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
setoriginalcolour
aboutshow.Forecolor = &HFFFF&
End Sub

Private Sub Alarmbut_Click()
'Show or hide the form?
MainCaption.Caption = "Console - Alarm Config"
Blank.Visible = True

frmMemo.Hide
frmProgs.Hide
frmFavs.Hide
frmMedia.Hide
frmNotes.Hide
frmPlaylist.Hide

WhichButton.Caption = 6
frmAlarmConfig.Show

Playsound "Button.wav"
End Sub

Private Sub Alarmbut_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
setoriginalcolour
Alarmbut.Forecolor = &HFFFF&
End Sub

Private Sub AlarmCon_Click()

End Sub

Private Sub AlarmTimer_Timer()
'clock
Timelbl.Caption = Format(Time, "HH:MM:SS AMPM")
'date
Datelbl.Caption = Format(Date, "dd/mm/yyyy")
End Sub



Private Sub Blank_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
setoriginalcolour
End Sub


Private Sub comCALC_Click()
On Error GoTo Error
X = Shell("c:\windows\calc.exe", vbNormalFocus)
Exit Sub

Error:
message = MsgBox("Error : " & Err.Number & " : " & Err.Description, vbCritical, "Error")

End Sub

Private Sub comCALC_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
setoriginalcolour
comCALC.Forecolor = &HFFFF&
End Sub

Private Sub comCD_Click()
On Error GoTo Error
X = Shell("c:\windows\cdplayer.exe", vbNormalFocus)
Exit Sub

Error:
message = MsgBox("Error : " & Err.Number & " : " & Err.Description, vbCritical, "Error")

End Sub

Private Sub comCD_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
setoriginalcolour
comCD.Forecolor = &HFFFF&
End Sub

Private Sub comCom_Click()
MainCaption.Caption = "Console - Command"
Blank.Visible = False
frmNotes.Hide
frmProgs.Hide
frmMemo.Hide
frmMedia.Hide
frmFavs.Hide
frmAlarmConfig.Hide
frmPlaylist.Hide
WhichButton.Caption = 0
Playsound "Button.wav"
End Sub

Private Sub comCom_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
setoriginalcolour
comCom.Forecolor = &HFFFF&
End Sub

Private Sub comControl_Click()
On Error GoTo Error
X = Shell("c:\windows\control.exe", 1)
Exit Sub

Error:
message = MsgBox("Error : " & Err.Number & " : " & Err.Description, vbCritical, "Error")

End Sub

Private Sub comControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
setoriginalcolour
comControl.Forecolor = &HFFFF&
End Sub

Private Sub comDOS_Click()
On Error GoTo Error
X = Shell("c:\windows\command.com", 3)
Exit Sub

Error:
message = MsgBox("Error : " & Err.Number & " : " & Err.Description, vbCritical, "Error")

End Sub

Private Sub comDOS_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
setoriginalcolour
comDOS.Forecolor = &HFFFF&
End Sub

Private Sub comDosEdit_Click()
On Error GoTo Error
X = Shell("c:\windows\command\edit.com", 3)
Exit Sub

Error:
message = MsgBox("Error : " & Err.Number & " : " & Err.Description, vbCritical, "Error")

End Sub

Private Sub comDosEdit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
setoriginalcolour
comDosEdit.Forecolor = &HFFFF&

End Sub

Private Sub comExplore_Click()
On Error GoTo Error
X = Shell("c:\windows\explorer.exe", 3)
Exit Sub

Error:
message = MsgBox("Error : " & Err.Number & " : " & Err.Description, vbCritical, "Error")

End Sub

Private Sub comExplore_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
setoriginalcolour
comExplore.Forecolor = &HFFFF&
End Sub

Private Sub comIP_Click()
On Error GoTo Error
X = Shell("c:\windows\winipcfg.exe", 1)
Exit Sub

Error:
message = MsgBox("Error : " & Err.Number & " : " & Err.Description, vbCritical, "Error")

End Sub

Private Sub comIP_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
setoriginalcolour
comIP.Forecolor = &HFFFF&
End Sub







Private Sub comREG_Click()
On Error GoTo Error
X = Shell("C:\Windows\Regedit.exe", 3)
Exit Sub

Error:
message = MsgBox("Error : " & Err.Number & " : " & Err.Description, vbCritical, "Error")

End Sub

Private Sub comREG_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
setoriginalcolour
comREG.Forecolor = &HFFFF&
End Sub



Private Sub comVolume_Click()
On Error GoTo Error
X = Shell("c:\windows\Sndvol32.exe", vbNormalFocus)
Exit Sub

Error:
message = MsgBox("Error : " & Err.Number & " : " & Err.Description, vbCritical, "Error")

End Sub

Private Sub comVolume_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
setoriginalcolour
comVolume.Forecolor = &HFFFF&
End Sub

Private Sub Gosearch_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
setoriginalcolour
Gosearch.Forecolor = &HFFFF&
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
setoriginalcolour
End Sub

Private Sub Label2_Click()
RestoreWindows
End Sub



Private Sub Datelbl_Click()
RestoreWindows
End Sub

Private Sub Form_Click()
RestoreWindows
End Sub

Private Sub Form_Terminate()
DeleteIconFromTray
End Sub

Private Sub Form_Unload(Cancel As Integer)
DeleteIconFromTray
End Sub

Private Sub Label1_Click()
RestoreWindows
End Sub

Private Sub Label3_Click()
On Error GoTo Error

Call ProgPath
frmHelp.Show
Exit Sub

Error:
message = MsgBox("Error loading help file: " & Err.Number & " : " & Err.Description, vbCritical, "Error")

End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
setoriginalcolour
Label3.Forecolor = &HFFFF&
End Sub

Private Sub GoSearch_Click()
Dim S As String
'Generate the search command string for the selected
'Search Engine
    Select Case Selected
    Case 0
        S = "http://search.yahoo.com/bin/search?p=" & Escaped.Caption
    Case 1
        S = "http://astalavista1.box.sk/cgi-bin/robot?srch=" & Escaped.Caption & "&submit=+search+&project=robot&gfx=robot"
        End Select
    
    ShellExecute Me.hwnd, "open", S, "", "", 1
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
setoriginalcolour
End Sub

Private Sub lExit_Click()
frmNotes.Hide
frmProgs.Hide
frmMemo.Hide
frmMedia.Hide
frmFavs.Hide
Me.Hide

End Sub

Private Sub lExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
setoriginalcolour
lExit.Forecolor = &HFFFF&
End Sub

Private Sub Form_Load()
Dim SavMin As String
Dim SavHour As String
'Center form - Cant use the forms properties, because when restored
' from the tray, it will go back to the center!

    Top = Screen.Height / 2 - Height / 2
    Left = Screen.Width / 2 - Width / 2
    searchtype(0).Value = True
'Console Command text
maintxt.Caption = "Console loaded at " & Format(Date, "dd/mm/yyyy") & " - " & Format(Time, "HH:MM:SS AMPM")

'Memo message count
If frmMemo.File1.ListCount <> 1 Then
MemoMess.Caption = " -------------------------" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "You have " & frmMemo.File1.ListCount & " messages in your memo."
Else
MemoMess.Caption = " -------------------------" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "You have " & frmMemo.File1.ListCount & " message in your memo."
End If
ver1.Caption = "V" & App.Major & "." & App.Minor & App.Revision
'Creates the Form
FakeForm Me, Picture1, RGB(255, 255, 255)
'Make Transparant
SetTrans

AddIconToTray True

'Load alarm stuff

Call ProgPath
On Error GoTo Error
Open ProgPath & "Alarm.dat" For Input As #1
Input #1, SavHour
Input #1, SavMin
Input #1, savAMPM
Input #1, savChecked
Input #1, savORIG
Input #1, savSHUT
Close #1

frmAlarmConfig.AlarmHOUR.Text = SavHour
frmAlarmConfig.AlarmMIN.Text = SavMin
frmAlarmConfig.AMPM.Text = savAMPM
frmAlarmConfig.Check1.Value = savChecked
frmAlarmConfig.AlarmOrig.Value = savORIG
frmAlarmConfig.AlarmShut.Value = savSHUT

If frmAlarmConfig.AlarmOrig.Value = 1 And frmAlarmConfig.AlarmShut.Value = 1 Then frmAlarmConfig.AlarmOrig.Value = 1

'Volume
VolumeCtrl.Value = frmMedia.Media1.Volume + 5000
Vol.Forecolor = RGB(VolumeCtrl.Value / 10, 255, 0)
Vol.Caption = VolumeCtrl.Value \ 50 & " %"
'OK! ALL DONE, play loading sound
Playsound "Beep.wav"
Exit Sub

Error:
MsgBox "Error loading alarm data - " & Err.Number & " : " & Err.Description
Close #1

End Sub

Private Sub FormDrag(TheForm As Form)
Dim X As Integer
ReleaseCapture
Call SendMessage(TheForm.hwnd, &HA1, 2, 0&)
'SNAP ON CODE
If Me.Left < 300 Then Me.Left = 0
If Me.Left > Screen.Width - (300 + Me.Width) Then Me.Left = Screen.Width - Me.Width
If Me.Top < 300 Then Me.Top = 0
If Me.Top > Screen.Height - (300 + Me.Height) Then Me.Top = Screen.Height - Me.Height

RestoreWindows

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
setoriginalcolour
'This line in AddIconToTray causes callback messages to be
'sent to this event: .uCallbackMessage = WM_MOUSEMOVE
'
'The actual callback message is contained in the X parameter.
'Note: when using this technique, X is a message not a coordinate.
On Local Error Resume Next
Err.Clear

Static bBusy As Boolean
    If bBusy = False Then           'Do one thing at a time
        bBusy = True
        Select Case CLng(X \ 15)
            Case WM_LBUTTONDBLCLK   'Double-click left mouse button: same as selecting About
            'frmAbout.Show 1
            Case WM_LBUTTONDOWN     'Left mouse button pressed: change traffic light icon & tip
                
            Case WM_LBUTTONUP       'Left mouse button released
                Main.Visible = True
                DoEvents
                AppActivate "BORDER"
                'Restore hidden windows:
RestoreWindows
AlwaysOnTop Me, True ' bring it to the front
If ontop.Value = 1 Then AlwaysOnTop Me, True
If ontop.Value = 0 Then AlwaysOnTop Me, False

            Case WM_RBUTTONDBLCLK   'Double-click right mouse button
            
            Case WM_RBUTTONDOWN     'Right mouse button pressed
            
            Case WM_RBUTTONUP       'Right mouse button released: display popup menu
                'PopupMenu frmmen.frmpop
        End Select
        bBusy = False
    End If
End Sub

Private Sub maintxt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
setoriginalcolour
End Sub

Private Sub MemoMess_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
setoriginalcolour
End Sub

Private Sub Min_Click()
message = MsgBox("Exiting the program will remove it from the systray and any alarms you have set will not sound.  Are you sure?", 36, "Exit?")
If message = 6 Then
Unload frmMemo
Unload frmNotes
Unload frmProgs
Unload frmFavs
Unload frmMedia
Unload Me
End
Else
RestoreWindows
End If
End Sub

Private Sub Min_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
setoriginalcolour
Min.Forecolor = &HFFFF&
End Sub

Private Sub mnuFavs_Click()
MainCaption.Caption = "Console - Directories"

Blank.Visible = True
'Show or hide the form?
frmMemo.Hide
frmProgs.Hide
frmNotes.Hide
frmMedia.Hide
frmAlarmConfig.Hide
frmPlaylist.Hide

WhichButton.Caption = 4

frmFavs.Show
Playsound "Button.wav"
End Sub

Private Sub mnuFavs_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
setoriginalcolour
mnuFavs.Forecolor = &HFFFF&
End Sub

Private Sub mnuMedia_Click()
MainCaption.Caption = "Console - Media"

Blank.Visible = True
'Show or hide the form?
frmMemo.Hide
frmProgs.Hide
frmFavs.Hide
frmNotes.Hide
frmAlarmConfig.Hide
frmPlaylist.Hide

WhichButton.Caption = 5

frmMedia.Show
Playsound "Button.wav"
End Sub

Private Sub mnuMedia_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
setoriginalcolour
mnuMedia.Forecolor = &HFFFF&
End Sub

Private Sub mnuMemo_Click()
MainCaption.Caption = "Console - Memo"

Blank.Visible = True
'Show or hide the form?
frmNotes.Hide
frmProgs.Hide
frmFavs.Hide
frmMedia.Hide
frmAlarmConfig.Hide
frmPlaylist.Hide

WhichButton.Caption = 3

frmMemo.Show
Playsound "Button.wav"
End Sub

Private Sub mnuMemo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
setoriginalcolour
mnuMemo.Forecolor = &HFFFF&
End Sub

Private Sub mnuPLAY_Click()
MainCaption.Caption = "Console - Playlist"

Blank.Visible = True
'Show or hide the form?
frmMemo.Hide
frmNotes.Hide
frmFavs.Hide
frmMedia.Hide
frmAlarmConfig.Hide
frmProgs.Hide

WhichButton.Caption = 7

frmPlaylist.Show
Playsound "Button.wav"

End Sub

Private Sub mnuPLAY_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
setoriginalcolour
mnuPLAY.Forecolor = &HFFFF&
End Sub

Private Sub mnuProgs_Click()
MainCaption.Caption = "Console - Programs"

Blank.Visible = True
'Show or hide the form?
frmMemo.Hide
frmNotes.Hide
frmFavs.Hide
frmMedia.Hide
frmAlarmConfig.Hide
frmPlaylist.Hide

WhichButton.Caption = 2

frmProgs.Show
Playsound "Button.wav"
End Sub

Private Sub mnuProgs_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
setoriginalcolour
mnuProgs.Forecolor = &HFFFF&
End Sub

Private Sub MainCaption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag Me

End Sub

Private Sub mnuNotes_Click()
'Show or hide the form?
MainCaption.Caption = "Console - Notes"
Blank.Visible = True

frmMemo.Hide
frmProgs.Hide
frmFavs.Hide
frmMedia.Hide
frmAlarmConfig.Hide
frmPlaylist.Hide

WhichButton.Caption = 1
frmNotes.Show

Playsound "Button.wav"
End Sub

Private Sub mnuNotes_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
setoriginalcolour
mnuNotes.Forecolor = &HFFFF&
End Sub

Private Sub netSearch_Change()
'Declare the required local variables
    Dim I As Integer, Buffer As String, CBuffer As String
'Get the Text1 TextBox text
    Buffer = netSearch.Text
'Check if it is empty
    If Buffer = "" Then
'If so, disable the Search CommandButton
        Gosearch.Enabled = False
    Else
'If not, enable the Search CommandButton
        Gosearch.Enabled = True
    End If
'Do for each letter of the Search String
    For I = 1 To Len(Buffer)
'Check the letters ASCII value
        Select Case Asc(Mid(Buffer, I, 1))
'Letters with no special encoding required, stay the same
        Case 42, 43, 45 To 57, 64 To 90, 95, 97 To 122
            CBuffer = CBuffer + Mid(Buffer, I, 1)
'Letters with special encoding required, are now coded
        Case Else
            CBuffer = CBuffer + "%" & Hex(Asc(Mid(Buffer, I, 1)))
        End Select
    Next I
'Show the encoded string on the Escaped Label
    Escaped.Caption = CBuffer
End Sub



Private Sub netSearch_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
setoriginalcolour
End Sub

Private Sub ontop_Click()
If ontop.Value = 1 Then AlwaysOnTop Me, True
If ontop.Value = 0 Then AlwaysOnTop Me, False
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Call the Form to Drag
FormDrag Me
End Sub

Private Function FakeForm(Form As Form, Titlebar As PictureBox, Forecolor As String)
'This is the Main-Function
'It calls the other needed Functions and sets the neede Controls to needed Values
'Remember: Borderstyle must be 0 and u mustn't use menus
'If Form.BorderStyle <> 0 Then Exit Function
'MakeBorder Form
'Titlebar.Left = 60
'Titlebar.Top = 60
'Titlebar.Height = 270
'Titlebar.Width = Form1.Width - 125
Titlebar.AutoRedraw = True
'If Title = "" Then Title = Form.Caption
Titlebar.Forecolor = Forecolor
Titlebar.CurrentX = 3
Titlebar.CurrentY = (Titlebar.ScaleHeight - Titlebar.TextHeight(Title)) / 2
Titlebar.Print Title

End Function

Private Sub Picture2_Click()
MainCaption.Caption = "Console - Command"
Blank.Visible = False
frmNotes.Hide
frmProgs.Hide
frmMemo.Hide
frmMedia.Hide
frmFavs.Hide
frmAlarmConfig.Hide
frmPlaylist.Hide
WhichButton.Caption = 0
Playsound "Button.wav"
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
setoriginalcolour
comCom.Forecolor = &HFFFF&
End Sub

Private Sub searchtype_Click(Index As Integer)
'Check if user selected a diferent Search Engine
    If Selected <> Index Then
'Set the selected option button FontBold property to True
'and the old one to False
        searchtype(Selected).FontBold = False
        searchtype(Index).FontBold = True
'Update the selected engine variable
        Selected = Index
    End If
End Sub

Private Sub searchtype_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
setoriginalcolour
End Sub

Private Sub snd_Click()
If snd.Value = 0 Then
frmMedia.Media1.Mute = False
VolumeCtrl.Enabled = True
Vol.Caption = VolumeCtrl.Value \ 50 & " %"
Else
frmMedia.Media1.Mute = True
VolumeCtrl.Enabled = False
Vol.Caption = "MUTE"
End If
End Sub

Private Sub Timelbl_Click()
RestoreWindows
End Sub


Private Sub ver1_Click()
RestoreWindows
End Sub

Private Sub VolumeCtrl_Change()
frmMedia.Media1.Volume = VolumeCtrl.Value - 5000
Vol.Forecolor = RGB(VolumeCtrl.Value / 10, 255, 0)
Vol.Caption = VolumeCtrl.Value \ 50 & " %"
End Sub

Private Sub VolumeCtrl_Scroll()
frmMedia.Media1.Volume = VolumeCtrl.Value - 5000
Vol.Forecolor = RGB(VolumeCtrl.Value / 10, 255, 0)
Vol.Caption = VolumeCtrl.Value \ 50 & " %"
End Sub
