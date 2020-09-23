VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmProgs 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Console-PROGRAMS"
   ClientHeight    =   4365
   ClientLeft      =   2805
   ClientTop       =   1125
   ClientWidth     =   4620
   LinkTopic       =   "Form1"
   ScaleHeight     =   6019.976
   ScaleMode       =   0  'User
   ScaleWidth      =   4560.769
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog Dialog1 
      Left            =   2400
      Top             =   3300
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Right click to choose new program."
      BeginProperty Font 
         Name            =   "Federation"
         Size            =   8.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   210
      Left            =   60
      TabIndex        =   10
      Top             =   4140
      Width           =   4470
   End
   Begin VB.Label NAME5 
      BackStyle       =   0  'Transparent
      Caption         =   "NAME5"
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
      Height          =   390
      Left            =   180
      TabIndex        =   9
      Top             =   3480
      Width           =   4170
   End
   Begin VB.Label NAME4 
      BackStyle       =   0  'Transparent
      Caption         =   "NAME4"
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
      Height          =   390
      Left            =   180
      TabIndex        =   8
      Top             =   2700
      Width           =   4170
   End
   Begin VB.Label NAME3 
      BackStyle       =   0  'Transparent
      Caption         =   "NAME3"
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
      Height          =   390
      Left            =   180
      TabIndex        =   7
      Top             =   1920
      Width           =   4170
   End
   Begin VB.Label NAME2 
      BackStyle       =   0  'Transparent
      Caption         =   "NAME2"
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
      Height          =   390
      Left            =   180
      TabIndex        =   6
      Top             =   1140
      Width           =   4170
   End
   Begin VB.Label NAME1 
      BackStyle       =   0  'Transparent
      Caption         =   "NAME1"
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
      Height          =   390
      Left            =   180
      TabIndex        =   5
      Top             =   360
      Width           =   4170
   End
   Begin VB.Label Prog5 
      Caption         =   "Prog5"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   3840
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Prog4 
      Caption         =   "Prog4"
      Height          =   255
      Left            =   180
      TabIndex        =   3
      Top             =   3180
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Prog3 
      Caption         =   "Prog3"
      Height          =   255
      Left            =   180
      TabIndex        =   2
      Top             =   2340
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Prog2 
      Caption         =   "Prog2"
      Height          =   255
      Left            =   180
      TabIndex        =   1
      Top             =   1560
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Prog1 
      Caption         =   "Prog1"
      Height          =   255
      Left            =   180
      TabIndex        =   0
      Top             =   780
      Visible         =   0   'False
      Width           =   1335
   End
End
Attribute VB_Name = "frmProgs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub setoriginalcolour()
NAME1.Forecolor = &H80FF&
NAME2.Forecolor = &H80FF&
NAME3.Forecolor = &H80FF&
NAME4.Forecolor = &H80FF&
NAME5.Forecolor = &H80FF&
End Sub
Private Sub LoadData()
Dim tmpNAME1 As String
Dim tmpNAME2 As String
Dim tmpNAME3 As String
Dim tmpNAME4 As String
Dim tmpNAME5 As String
Dim tmpProg1 As String
Dim tmpProg2 As String
Dim tmpProg3 As String
Dim tmpProg4 As String
Dim tmpProg5 As String

'Load data to temporary strings
On Error GoTo Error
Call ProgPath
Open ProgPath & "Programs.dat" For Input As #1
Input #1, tmpNAME1
Input #1, tmpProg1
Input #1, tmpNAME2
Input #1, tmpProg2
Input #1, tmpNAME3
Input #1, tmpProg3
Input #1, tmpNAME4
Input #1, tmpProg4
Input #1, tmpNAME5
Input #1, tmpProg5

Close #1

'Assign to objects
NAME1.Caption = tmpNAME1
Prog1.Caption = tmpProg1
NAME2.Caption = tmpNAME2
Prog2.Caption = tmpProg2
NAME3.Caption = tmpNAME3
Prog3.Caption = tmpProg3
NAME4.Caption = tmpNAME4
Prog4.Caption = tmpProg4
NAME5.Caption = tmpNAME5
Prog5.Caption = tmpProg5

Exit Sub

Error:
MsgBox "Error loading programs.dat data - " & Err.Number & " : " & Err.Description

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

Private Sub Form_Load()
zz = SetParent(Me.hwnd, Main.hwnd)
LoadData
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
setoriginalcolour
End Sub

Private Sub NAME1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
setoriginalcolour
NAME1.Forecolor = &HFFFF&
End Sub

Private Sub NAME1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Friendly1 As String
On Error GoTo Error
If Button = 2 Then
AlwaysOnTop Main, False
'Open Dialog
Dialog1.CancelError = True
Dialog1.DialogTitle = "Choose a program"
Dialog1.Filter = "Executables *.exe|*.exe"
Dialog1.filename = ""
Dialog1.ShowOpen

'Friendly Name
Friendly1 = InputBox("Enter a friendly name for this program:", "Friendly name")
Call ProgPath
Open ProgPath & "Programs.dat" For Output As #1
Print #1, Friendly1
Print #1, Dialog1.filename
Print #1, NAME2.Caption
Print #1, Prog2.Caption
Print #1, NAME3.Caption
Print #1, Prog3.Caption
Print #1, NAME4.Caption
Print #1, Prog4.Caption
Print #1, NAME5.Caption
Print #1, Prog5.Caption
Close #1
LoadData
RestoreWindows
If Main.ontop.Value = 1 Then AlwaysOnTop Main, True
If Main.ontop.Value = 0 Then AlwaysOnTop Main, False
Exit Sub
Else
If Button = 1 Then
'ChDir Prog1.Caption
X = Shell(Prog1.Caption, vbNormalFocus)
Exit Sub
Else

End If
End If

Exit Sub
Error:
RestoreWindows
If Main.ontop.Value = 1 Then AlwaysOnTop Main, True
If Main.ontop.Value = 0 Then AlwaysOnTop Main, False
If Err.Number = 32755 Then Exit Sub
Close #1
If Err.Number = 5 Then
MsgBox "Please choose a valid program first!", vbCritical, "Error"
Else
MsgBox "Error opening program - " & Err.Number & " : " & Err.Description
End If
End Sub

Private Sub NAME2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
setoriginalcolour
NAME2.Forecolor = &HFFFF&
End Sub

Private Sub NAME2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Friendly2 As String

'If user right clicked....let them choose a program
On Error GoTo Error
If Button = 2 Then
AlwaysOnTop Main, False
'Open Dialog
Dialog1.CancelError = True
Dialog1.DialogTitle = "Choose a program"
Dialog1.Filter = "Executables *.exe|*.exe"
Dialog1.filename = ""
Dialog1.ShowOpen

'Friendly Name
Friendly2 = InputBox("Enter a friendly name for this program:", "Friendly name")
Call ProgPath
Open ProgPath & "Programs.dat" For Output As #1
Print #1, NAME1.Caption
Print #1, Prog1.Caption
Print #1, Friendly2
Print #1, Dialog1.filename
Print #1, NAME3.Caption
Print #1, Prog3.Caption
Print #1, NAME4.Caption
Print #1, Prog4.Caption
Print #1, NAME5.Caption
Print #1, Prog5.Caption
Close #1
LoadData
RestoreWindows
If Main.ontop.Value = 1 Then AlwaysOnTop Main, True
If Main.ontop.Value = 0 Then AlwaysOnTop Main, False
Exit Sub
Else

'If Mouse Button 1 was pressed, RUN the program
If Button = 1 Then
'ChDir Prog1.Caption
X = Shell(Prog2.Caption, vbNormalFocus)
Exit Sub
Else

End If
End If

Exit Sub
Error:
RestoreWindows
If Main.ontop.Value = 1 Then AlwaysOnTop Main, True
If Main.ontop.Value = 0 Then AlwaysOnTop Main, False
If Err.Number = 32755 Then Exit Sub
Close #1
If Err.Number = 5 Then
MsgBox "Please choose a valid program first!", vbCritical, "Error"
Else
MsgBox "Error opening program - " & Err.Number & " : " & Err.Description
End If
End Sub


Private Sub NAME3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
setoriginalcolour
NAME3.Forecolor = &HFFFF&
End Sub

Private Sub NAME3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Friendly3 As String

'If user right clicked....let them choose a program
On Error GoTo Error
If Button = 2 Then
AlwaysOnTop Main, False
'Open Dialog
Dialog1.CancelError = True
Dialog1.DialogTitle = "Choose a program"
Dialog1.Filter = "Executables *.exe|*.exe"
Dialog1.filename = ""
Dialog1.ShowOpen

'Friendly Name
Friendly3 = InputBox("Enter a friendly name for this program:", "Friendly name")
Call ProgPath
Open ProgPath & "Programs.dat" For Output As #1
Print #1, NAME1.Caption
Print #1, Prog1.Caption
Print #1, NAME2.Caption
Print #1, Prog2.Caption
Print #1, Friendly3
Print #1, Dialog1.filename
Print #1, NAME4.Caption
Print #1, Prog4.Caption
Print #1, NAME5.Caption
Print #1, Prog5.Caption
Close #1
LoadData
RestoreWindows
If Main.ontop.Value = 1 Then AlwaysOnTop Main, True
If Main.ontop.Value = 0 Then AlwaysOnTop Main, False
Exit Sub
Else

'If Mouse Button 1 was pressed, RUN the program
If Button = 1 Then
'ChDir Prog1.Caption
X = Shell(Prog3.Caption, vbNormalFocus)
Exit Sub
Else

End If
End If

Exit Sub
Error:
RestoreWindows
If Main.ontop.Value = 1 Then AlwaysOnTop Main, True
If Main.ontop.Value = 0 Then AlwaysOnTop Main, False
If Err.Number = 32755 Then Exit Sub
Close #1
If Err.Number = 5 Then
MsgBox "Please choose a valid program first!", vbCritical, "Error"
Else
MsgBox "Error opening program - " & Err.Number & " : " & Err.Description
End If
End Sub


Private Sub NAME4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
setoriginalcolour
NAME4.Forecolor = &HFFFF&
End Sub

Private Sub NAME4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Friendly4 As String

'If user right clicked....let them choose a program
On Error GoTo Error
If Button = 2 Then
AlwaysOnTop Main, False
'Open Dialog
Dialog1.CancelError = True
Dialog1.DialogTitle = "Choose a program"
Dialog1.Filter = "Executables *.exe|*.exe"
Dialog1.filename = ""
Dialog1.ShowOpen

'Friendly Name
Friendly4 = InputBox("Enter a friendly name for this program:", "Friendly name")
Call ProgPath
Open ProgPath & "Programs.dat" For Output As #1
Print #1, NAME1.Caption
Print #1, Prog1.Caption
Print #1, NAME2.Caption
Print #1, Prog2.Caption
Print #1, NAME3.Caption
Print #1, Prog3.Caption
Print #1, Friendly4
Print #1, Dialog1.filename
Print #1, NAME5.Caption
Print #1, Prog5.Caption
Close #1
LoadData
RestoreWindows
If Main.ontop.Value = 1 Then AlwaysOnTop Main, True
If Main.ontop.Value = 0 Then AlwaysOnTop Main, False
Exit Sub
Else

'If Mouse Button 1 was pressed, RUN the program
If Button = 1 Then
'ChDir Prog1.Caption
X = Shell(Prog4.Caption, vbNormalFocus)
Exit Sub
Else

End If
End If

Exit Sub
Error:
RestoreWindows
If Main.ontop.Value = 1 Then AlwaysOnTop Main, True
If Main.ontop.Value = 0 Then AlwaysOnTop Main, False
If Err.Number = 32755 Then Exit Sub
Close #1
If Err.Number = 5 Then
MsgBox "Please choose a valid program first!", vbCritical, "Error"
Else
MsgBox "Error opening program - " & Err.Number & " : " & Err.Description
End If
End Sub

Private Sub NAME5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
setoriginalcolour
NAME5.Forecolor = &HFFFF&
End Sub

Private Sub NAME5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Friendly5 As String

'If user right clicked....let them choose a program
On Error GoTo Error
If Button = 2 Then
AlwaysOnTop Main, False
'Open Dialog
Dialog1.CancelError = True
Dialog1.DialogTitle = "Choose a program"
Dialog1.Filter = "Executables *.exe|*.exe"
Dialog1.filename = ""
Dialog1.ShowOpen

'Friendly Name
Friendly5 = InputBox("Enter a friendly name for this program:", "Friendly name")
Call ProgPath
Open ProgPath & "Programs.dat" For Output As #1
Print #1, NAME1.Caption
Print #1, Prog1.Caption
Print #1, NAME2.Caption
Print #1, Prog2.Caption
Print #1, NAME3.Caption
Print #1, Prog3.Caption
Print #1, NAME4.Caption
Print #1, Prog4.Caption
Print #1, Friendly5
Print #1, Dialog1.filename
Close #1
LoadData
RestoreWindows
If Main.ontop.Value = 1 Then AlwaysOnTop Main, True
If Main.ontop.Value = 0 Then AlwaysOnTop Main, False
Exit Sub
Else

'If Mouse Button 1 was pressed, RUN the program
If Button = 1 Then
'ChDir Prog1.Caption
X = Shell(Prog5.Caption, vbNormalFocus)
Exit Sub
Else

End If
End If

Exit Sub
Error:
RestoreWindows
If Main.ontop.Value = 1 Then AlwaysOnTop Main, True
If Main.ontop.Value = 0 Then AlwaysOnTop Main, False
If Err.Number = 32755 Then Exit Sub
Close #1
If Err.Number = 5 Then
MsgBox "Please choose a valid program first!", vbCritical, "Error"
Else
MsgBox "Error opening program - " & Err.Number & " : " & Err.Description
End If
End Sub
