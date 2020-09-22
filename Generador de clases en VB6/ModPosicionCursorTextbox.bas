Attribute VB_Name = "ModPosicionCursorTextbox"
Const EM_LINEFROMCHAR = &HC9
Const EM_LINEINDEX = &HBB
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Function GetCurrentCol(tbox As RichTextBox) As Long
Dim CurLine As Long
Dim TotalLines As Long
Dim t As Long
CurLine = SendMessage(tbox.hwnd, EM_LINEFROMCHAR, -1, 0)
t = 0
If CurLine > 0 Then
    CurCol = SendMessage(tbox.hwnd, EM_LINEINDEX, CurLine, 0)
    t = CurCol
End If
GetCurrentCol = tbox.SelStart - t + 1
End Function

Public Function GetCurrentLine(tbox As RichTextBox) As Long
GetCurrentLine = SendMessage(tbox.hwnd, EM_LINEFROMCHAR, -1, 0) + 1
End Function
