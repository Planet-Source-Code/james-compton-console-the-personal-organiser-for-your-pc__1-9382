VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form frmHelp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Console Help"
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   7845
   Icon            =   "frmHelp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   7845
   StartUpPosition =   3  'Windows Default
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   5820
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7845
      ExtentX         =   13838
      ExtentY         =   10266
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
On Error GoTo Error
WebBrowser1.Navigate ProgPath & "help\index.htm"
Exit Sub

Error:
MsgBox "A problem occured while trying to access the help file. " & Err.Number & " - " & Err.Description, vbCritical, "Help file problem"
End Sub

