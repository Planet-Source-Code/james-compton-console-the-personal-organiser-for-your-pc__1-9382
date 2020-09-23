VERSION 5.00
Begin VB.Form frmmen 
   Caption         =   "Form1"
   ClientHeight    =   3165
   ClientLeft      =   165
   ClientTop       =   765
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3165
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu frmpop 
      Caption         =   "POPUP MENU"
      Begin VB.Menu mediaplayerplaytrack 
         Caption         =   "Play"
      End
      Begin VB.Menu trkinfo 
         Caption         =   "Track Info."
      End
   End
End
Attribute VB_Name = "frmmen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub conexit_Click()
Unload frmMemo
Unload frmNotes
Unload frmProgs
Unload frmFavs
Unload frmMedia
Unload Main
End
End Sub

Private Sub frmaboutshow_Click()
frmAbout.Show 1
End Sub

Private Sub mediaplayerplaytrack_Click()
Dim TMP As Integer
If frmPlaylist.plyLIST.ListCount = 0 Then
MsgBox "There are no files to transfer!", vbCritical
Exit Sub
Else
frmMedia.playlistfile.Caption = 1
frmMedia.plyLIST.Clear
frmMedia.Media1.AutoStart = True
frmMedia.Media1.Visible = True
TMP = frmPlaylist.plyLIST.ListIndex ' get users track
frmPlaylist.plyLIST.Visible = False ' to stop scrolling
For X = 0 To frmPlaylist.plyLIST.ListCount - 1
DoEvents
    frmPlaylist.plyLIST.ListIndex = X
    frmMedia.plyLIST.AddItem (frmPlaylist.plyLIST.Text)
Next
frmPlaylist.plyLIST.Visible = True ' bring it back!
frmMedia.plyLIST.ListIndex = TMP
frmPlaylist.plyLIST.ListIndex = TMP

On Error Resume Next
If frmMedia.plyLIST.Text <> "" Then
frmMedia.Media1.filename = frmMedia.plyLIST.Text
frmMedia.Media1.Play
frmMedia.Dialog1.filename = ""
Playsound "Beep.wav"
Exit Sub

End If
Exit Sub
End If

End Sub

Private Sub trkinfo_Click()
MsgBox "File : " & frmPlaylist.plyLIST.Text & Chr(13) & Chr(10) & "Track No : " & frmPlaylist.plyLIST.ListIndex + 1 & Chr(13) & Chr(10) & "Total : " & frmPlaylist.plyLIST.ListCount, vbInformation, "File Information"
End Sub
