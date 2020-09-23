VERSION 5.00
Begin VB.Form frmDirSelect 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Choose Directory"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   3900
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   3900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   495
      Left            =   1140
      TabIndex        =   2
      Top             =   3540
      Width           =   1575
   End
   Begin VB.DirListBox Dir1 
      Height          =   3015
      Left            =   60
      TabIndex        =   1
      Top             =   420
      Width           =   3795
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   3795
   End
End
Attribute VB_Name = "frmDirSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmFavs.tmpDir.Caption = DIR1.Path
Unload Me
End Sub

Private Sub Drive1_Change()
On Error GoTo error
DIR1.Path = Drive1.Drive
Exit Sub

error:
MsgBox "Error " & Err.Number & " : " & Err.Description
End Sub
