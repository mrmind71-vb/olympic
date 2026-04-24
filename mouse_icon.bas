Attribute VB_Name = "Module2"
Private Const IDC_HAND = 32649&

Private Const IDC_APPSTARTING = 32650&
Private Const IDC_ARROW = 32512&
Private Const IDC_CROSS = 32515&
Private Const IDC_IBEAM = 32513&
Private Const IDC_ICON = 32641&
Private Const IDC_NO = 32648&
Private Const IDC_SIZE = 32640&
Private Const IDC_SIZEALL = 32646&
Private Const IDC_SIZENESW = 32643&
Private Const IDC_SIZENS = 32645&
Private Const IDC_SIZENWSE = 32642&
Private Const IDC_SIZEWE = 32644&
Private Const IDC_UPARROW = 32516&
Private Const IDC_WAIT = 32514&

Public Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" _
 (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
 
Public Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long

