VERSION 5.00
Begin VB.Form frmFavs 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Console-DIR"
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
   Begin VB.Label tmpDir 
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   3780
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Right click to choose new directory."
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
      TabIndex        =   5
      Top             =   4140
      Width           =   4470
   End
   Begin VB.Label DIR3 
      BackStyle       =   0  'Transparent
      Caption         =   "DIR3"
      BeginProperty Font 
         Name            =   "Federation"
         Size            =   12
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   270
      Left            =   180
      TabIndex        =   4
      Top             =   1860
      Width           =   4170
   End
   Begin VB.Label DIR4 
      BackStyle       =   0  'Transparent
      Caption         =   "DIR4"
      BeginProperty Font 
         Name            =   "Federation"
         Size            =   12
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   270
      Left            =   180
      TabIndex        =   3
      Top             =   2580
      Width           =   4170
   End
   Begin VB.Label DIR5 
      BackStyle       =   0  'Transparent
      Caption         =   "DIR5"
      BeginProperty Font 
         Name            =   "Federation"
         Size            =   12
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   270
      Left            =   180
      TabIndex        =   2
      Top             =   3300
      Width           =   4170
   End
   Begin VB.Label DIR2 
      BackStyle       =   0  'Transparent
      Caption         =   "DIR2"
      BeginProperty Font 
         Name            =   "Federation"
         Size            =   12
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   270
      Left            =   180
      TabIndex        =   1
      Top             =   1140
      Width           =   4170
   End
   Begin VB.Label DIR1 
      BackStyle       =   0  'Transparent
      Caption         =   "DIR1"
      BeginProperty Font 
         Name            =   "Federation"
         Size            =   12
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   270
      Left            =   180
      TabIndex        =   0
      Top             =   420
      Width           =   4170
   End
End
Attribute VB_Name = "frmFavs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub setoriginalcolour()
Dir1.Forecolor = &H80FF&
DIR2.Forecolor = &H80FF&
DIR3.Forecolor = &H80FF&
DIR4.Forecolor = &H80FF&
DIR5.Forecolor = &H80FF&
End Sub
Private Sub LoadData()
Dim tDir1 As String
Dim tDir2 As String
Dim tDir3 As String
Dim tDir4 As String
Dim tDir5 As String

On Error GoTo Error

'Load directories
Call ProgPath
Open ProgPath & "Directory.dat" For Input As #1
Input #1, tDir1
Input #1, tDir2
Input #1, tDir3
Input #1, tDir4
Input #1, tDir5
Close #1

'Assign em!
Dir1.Caption = tDir1
DIR2.Caption = tDir2
DIR3.Caption = tDir3
DIR4.Caption = tDir4
DIR5.Caption = tDir5
Exit Sub

Error:
MsgBox "Error loading directory.dat data - " & Err.Number & " : " & Err.Description

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

Private Sub Dir1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
setoriginalcolour
Dir1.Forecolor = &HFFFF&
End Sub

Private Sub DIR1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
On Error GoTo Error
'Open Dialog
AlwaysOnTop Main, False
frmDirSelect.Show 1
RestoreWindows
If Main.ontop.Value = 1 Then AlwaysOnTop Main, True
If Main.ontop.Value = 0 Then AlwaysOnTop Main, False

Call ProgPath
Open ProgPath & "Directory.dat" For Output As #1
Print #1, tmpDir.Caption
Print #1, DIR2.Caption
Print #1, DIR3.Caption
Print #1, DIR4.Caption
Print #1, DIR5.Caption
Close #1
LoadData
Me.Show
Exit Sub
Else
If Button = 1 Then
'ChDir Prog1.Caption
X = Shell("C:\Windows\Explorer.exe " & Dir1.Caption, vbNormalFocus)
Exit Sub
Else

End If
End If

Exit Sub
Error:
Close #1
MsgBox "Error " & Err.Number & " : " & Err.Description
RestoreWindows
If Main.ontop.Value = 1 Then AlwaysOnTop Main, True
If Main.ontop.Value = 0 Then AlwaysOnTop Main, False
Exit Sub
End Sub

Private Sub DIR2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
setoriginalcolour
DIR2.Forecolor = &HFFFF&
End Sub

Private Sub DIR2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
On Error GoTo Error
AlwaysOnTop Main, False
'Open Dialog
frmDirSelect.Show 1

RestoreWindows
If Main.ontop.Value = 1 Then AlwaysOnTop Main, True
If Main.ontop.Value = 0 Then AlwaysOnTop Main, False

Call ProgPath
Open ProgPath & "Directory.dat" For Output As #1
Print #1, Dir1.Caption
Print #1, tmpDir.Caption
Print #1, DIR3.Caption
Print #1, DIR4.Caption
Print #1, DIR5.Caption
Close #1
LoadData
Me.Show
Exit Sub
Else
If Button = 1 Then
'ChDir Prog1.Caption
X = Shell("C:\Windows\Explorer.exe " & DIR2.Caption, vbNormalFocus)
Exit Sub
Else
End If
End If

Exit Sub
Error:
Close #1
MsgBox "Error " & Err.Number & " : " & Err.Description
RestoreWindows
If Main.ontop.Value = 1 Then AlwaysOnTop Main, True
If Main.ontop.Value = 0 Then AlwaysOnTop Main, False
Exit Sub
End Sub

Private Sub DIR3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
setoriginalcolour
DIR3.Forecolor = &HFFFF&
End Sub

Private Sub DIR3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
On Error GoTo Error
AlwaysOnTop Main, False
'Open Dialog
frmDirSelect.Show 1
RestoreWindows
If Main.ontop.Value = 1 Then AlwaysOnTop Main, True
If Main.ontop.Value = 0 Then AlwaysOnTop Main, False

Call ProgPath
Open ProgPath & "Directory.dat" For Output As #1
Print #1, Dir1.Caption
Print #1, DIR2.Caption
Print #1, tmpDir.Caption
Print #1, DIR4.Caption
Print #1, DIR5.Caption
Close #1
LoadData
Me.Show
Exit Sub
Else
If Button = 1 Then
'ChDir Prog1.Caption
X = Shell("C:\Windows\Explorer.exe " & DIR3.Caption, vbNormalFocus)
Exit Sub
Else
End If
End If

Exit Sub
Error:
Close #1
MsgBox "Error " & Err.Number & " : " & Err.Description
RestoreWindows
If Main.ontop.Value = 1 Then AlwaysOnTop Main, True
If Main.ontop.Value = 0 Then AlwaysOnTop Main, False
Exit Sub
End Sub

Private Sub DIR4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
setoriginalcolour
DIR4.Forecolor = &HFFFF&
End Sub

Private Sub DIR4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
On Error GoTo Error
AlwaysOnTop Main, False
'Open Dialog
frmDirSelect.Show 1
RestoreWindows
If Main.ontop.Value = 1 Then AlwaysOnTop Main, True
If Main.ontop.Value = 0 Then AlwaysOnTop Main, False
Call ProgPath
Open ProgPath & "Directory.dat" For Output As #1
Print #1, Dir1.Caption
Print #1, DIR2.Caption
Print #1, DIR3.Caption
Print #1, tmpDir.Caption
Print #1, DIR5.Caption
Close #1
LoadData
Me.Show
Exit Sub
Else
If Button = 1 Then
'ChDir Prog1.Caption
X = Shell("C:\Windows\Explorer.exe " & DIR4.Caption, vbNormalFocus)
Exit Sub
Else
End If
End If

Exit Sub
Error:
Close #1
MsgBox "Error " & Err.Number & " : " & Err.Description
RestoreWindows
If Main.ontop.Value = 1 Then AlwaysOnTop Main, True
If Main.ontop.Value = 0 Then AlwaysOnTop Main, False
Exit Sub
End Sub

Private Sub DIR5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
setoriginalcolour
DIR5.Forecolor = &HFFFF&
End Sub

Private Sub DIR5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
On Error GoTo Error
AlwaysOnTop Main, False
'Open Dialog
frmDirSelect.Show 1
RestoreWindows
If Main.ontop.Value = 1 Then AlwaysOnTop Main, True
If Main.ontop.Value = 0 Then AlwaysOnTop Main, False
Call ProgPath
Open ProgPath & "Directory.dat" For Output As #1
Print #1, Dir1.Caption
Print #1, DIR2.Caption
Print #1, DIR3.Caption
Print #1, DIR4.Caption
Print #1, tmpDir.Caption
Close #1
LoadData
Me.Show
Exit Sub
Else
If Button = 1 Then
'ChDir Prog1.Caption
X = Shell("C:\Windows\Explorer.exe " & DIR5.Caption, vbNormalFocus)
Exit Sub
Else
End If
End If

Exit Sub
Error:
Close #1
MsgBox "Error " & Err.Number & " : " & Err.Description
RestoreWindows
If Main.ontop.Value = 1 Then AlwaysOnTop Main, True
If Main.ontop.Value = 0 Then AlwaysOnTop Main, False
Exit Sub
End Sub

Private Sub Form_Load()
zz = SetParent(Me.hwnd, Main.hwnd)
LoadData
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
setoriginalcolour
End Sub
