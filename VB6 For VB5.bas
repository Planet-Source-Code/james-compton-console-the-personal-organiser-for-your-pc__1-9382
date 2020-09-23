Attribute VB_Name = "basVB6ForVB5"
Option Explicit


Public Function InStrRevVB5(ByVal StringToCheck As String, ByVal StringToMatch As String, Optional ByVal StartAt As Long = -1, Optional ByVal Compare As VbCompareMethod = vbBinaryCompare) As Long
 
Dim lPos        As Long
Dim lSavePos    As Long
 
    ' -1 means search entire string. A positive number
    ' means search only up to that position from the left.
    If StartAt = -1 Then StartAt = Len(StringToCheck)
    
    ' Find the last instance of StringToMatch within StringToCheck.
    lPos = InStr(1, StringToCheck, StringToMatch, Compare)
    While lPos > 0 And lPos < StartAt
        lSavePos = lPos
        lPos = InStr(lPos + 1, StringToCheck, StringToMatch, Compare)
    Wend
    
    InStrRevVB5 = lSavePos
        
End Function

