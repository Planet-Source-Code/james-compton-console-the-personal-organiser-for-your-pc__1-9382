Attribute VB_Name = "Sounds"
  Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
  
  Public Sub Playsound(SoundName As String)
On Error GoTo error
  'Play sound!
    Dim X As Integer
    Dim Var As Long
    
If Mid(App.Path, Len(App.Path)) = "\" Then
Path = App.Path
' If dragged file is not in root, append "\" and filename.
Else
Path = App.Path & "\"
End If
    DoEvents
    SoundName2$ = Path & "Sounds\" & SoundName
    Var& = SND_ASYNC Or SND_NODEFAULT
    X% = sndPlaySound(SoundName2$, Var&)
 Exit Sub
error:
 MsgBox "Error finding sound - " & Err.Number & " : " & Err.Description
  End Sub
