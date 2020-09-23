Attribute VB_Name = "zOnTop"
Option Explicit
#If Win16 Then 'Conditional Compile statements
    Declare Sub SetWindowPos Lib "User" (ByVal hWnd As Integer, ByVal hWndInsertAfter As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal wFlags As Integer)
#Else
    Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
#End If

Sub KeepOnTop(F As Form)
'sets the given form On TopMost
Const SWP_NOMOVE = 2
Const SWP_NOSIZE = 1

Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2

    SetWindowPos F.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE

End Sub
Sub KeepOffTop(F As Form)
'sets the given form Off TopMost
Const SWP_NOMOVE = 2
Const SWP_NOSIZE = 1

Const HWND_TOPMOST = 1
Const HWND_NOTOPMOST = 2

    SetWindowPos F.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE

End Sub
