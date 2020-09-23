Attribute VB_Name = "FindInList"

'System Tray
 Const EWX_SHUTDOWN = 1
'Parent
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

'Messages
Public Const NIM_ADD = &H0             'Adds an icon to the taskbar notification area
Public Const NIM_MODIFY = &H1          'Changes the icon, tooltip text or notification message for an icon in the notification area
Public Const NIM_DELETE = &H2          'Deletes an icon from the taskbar notification area

'Flags
Public Const NIF_MESSAGE = &H1         'hIcon is valid
Public Const NIF_ICON = &H2            'uCallbackMessage is valid
Public Const NIF_TIP = &H4             'szTip is valid

Public Const WM_MOUSEMOVE = &H200      'MouseMove message identifier
                                    
Public Const WM_LBUTTONDBLCLK = &H203  'Messages sent to the form's MouseMove event
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205

Public Type NOTIFYICONDATA
    cbSize              As Long
    hwnd                As Long         'Handle of window that receives notification messages
    uID                 As Long         'Application-defined identifier of the taskbar icon
    uFlags              As Long         'Flags indicating which structure members contain valid data
    uCallbackMessage    As Long         'Application defined callback message
    hIcon               As Long         'Handle of taskbar icon
    szTip               As String * 64  'Tooltip text to display for icon
End Type
'MISTIFY STUFF
Public RST As Integer, A As Integer, DelLines As Boolean
Public PSWD As String, Pts As Integer, P As Boolean, SS As Boolean

Public mtIconData As NOTIFYICONDATA, mnLight As Integer
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Long
Private Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long




Function FindExactInCombo(Ctl As Control, S As String)
    Dim I As Long, j As Long, k As Long
    FindExactInCombo = -1
    I = 0: j = Ctl.ListCount - 1


    Do
        'non trovato, esce
        If I > j Then Exit Function
        k = (I + j) / 2


        Select Case StrComp(Ctl.List(k), S)
            Case 0: Exit Do
            Case -1: I = k + 1 ' if < look in the second half
            Case 1: j = k - 1 ' if > look in the first half
        End Select
Loop
'sequential search backwards to found th
'     e first matching element


Do While k > 0
    If StrComp(Ctl.List(k - 1), S) <> 0 Then Exit Do
    k = k - 1
Loop
FindExactInCombo = k
End Function

Public Sub AddIconToTray(XA As Boolean) 'Adds an icon to the taskbar notification area

With mtIconData
    .cbSize = Len(mtIconData)
    .hwnd = Main.hwnd                                     'Use the form to receive callback messages.
    .uCallbackMessage = WM_MOUSEMOVE                    'Tell icon to send MouseMove messages.
    .uID = 1&                                           'Application defined identifier
    .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    .hIcon = Main.Icon                        'Initial icon
    .szTip = Main.Tag & Chr(0)               'Initial tooltip for icon
    If XA Then
        If Shell_NotifyIcon(NIM_ADD, mtIconData) = 0 Then   'Create icon in tray
            'MsgBox "Unable to add icon to system tray!"
        End If
     Else
        'NIM_MODIFY
        If Shell_NotifyIcon(NIM_MODIFY, mtIconData) = 0 Then   'Create icon in tray
            'MsgBox "Unable to edit icon in system tray!"
        End If
    End If
End With
    
End Sub

Public Sub DeleteIconFromTray()
    
Dim TT As Integer
If Shell_NotifyIcon(NIM_DELETE, mtIconData) = 0 Then TT = 1

End Sub
Sub PercentBar(Shape As Control, Done As Integer, Total As Variant)

'Call PercentBar(Picture1, Label1.Caption, Label2.Caption)

On Error Resume Next
Shape.AutoRedraw = True
Shape.FillStyle = 0
Shape.DrawStyle = 0
Shape.FontName = "MS Sans Serif"
Shape.FontSize = 8.25
Shape.FontBold = False
X = Done / Total * Shape.Width
Shape.Line (0, 0)-(Shape.Width, Shape.Height), RGB(0, 0, 0), BF
Shape.Line (0, 0)-(X - 10, Shape.Height), RGB(255, 127, 0), BF
Shape.CurrentX = (Shape.Width / 2) - 100
Shape.CurrentY = (Shape.Height / 2) - 125
Shape.Forecolor = RGB(192, 192, 192)
'Shape.Print Percent(Done, Total, 100) & "%"
End Sub

Public Sub Pause(Duration As Long)
    Dim Current As Long
    Current = Timer
    Do Until Timer - Current >= Duration
        DoEvents
    Loop
End Sub
Public Sub ShutDownWindows()
Dim Ret As Long
On Error GoTo errHandler
    'Shut Down  the computer
    Ret& = ExitWindowsEx(EWX_SHUTDOWN, 0)
Exit Sub
errHandler:
MsgBox "Error shutting down computer - " & Err.Number & " : " & Err.Description
End Sub
Public Sub RestoreWindows()
'Restore hidden windows:
Select Case Main.WhichButton.Caption
'Notepad
Case 1
frmNotes.Show

'Programs
Case 2
frmProgs.Show

'Memo
Case 3
frmMemo.Show

'Favs
Case 4
frmFavs.Show

'Media
Case 5
frmMedia.Show

'Alarm Config
Case 6
frmAlarmConfig.Show

End Select
End Sub
