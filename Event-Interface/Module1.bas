Attribute VB_Name = "Module1"
Declare Function GetTickCount Lib "kernel32" () As Long
Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, lpRect As Any, ByVal bErase As Long) As Long
Public u As Long
