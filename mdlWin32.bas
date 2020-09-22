Attribute VB_Name = "mdlWin32"
Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

Type POINTAPI
    X As Long
    Y As Long
End Type

