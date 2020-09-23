Attribute VB_Name = "Window_Ontop"
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Function AlwaysOnTop(Form1, Optional TopMost As Boolean = True)
    Const HWND_TOPMOST = -&H1
    Const HWND_NOTOPMOST = -&H2
    Const SWP_NOSIZE = &H1
    Const SWP_NOMOVE = &H2
    AlwaysOnTop = SetWindowPos(Form1.hwnd, IIf(TopMost, HWND_TOPMOST, HWND_NOTOPMOST), 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
End Function



