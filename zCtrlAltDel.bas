Attribute VB_Name = "zCtrlAltDel"
Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Public Declare Function RegisterServiceProcess Lib "kernel32" (ByVal dwProcessID As Long, ByVal dwType As Long) As Long
    Public Const RSP_SIMPLE_SERVICE = 1
    Public Const RSP_UNREGISTER_SERVICE = 0
Public Sub NOSHOW()
    Dim pid As Long
    Dim reserv As Long
    pid = GetCurrentProcessId()
    regserv = RegisterServiceProcess(pid, RSP_SIMPLE_SERVICE)
End Sub
Public Sub DOSHOW()
    Dim pid As Long
    Dim reserv As Long
    pid = GetCurrentProcessId()
    regserv = RegisterServiceProcess(pid, RSP_UNREGISTER_SERVICE)
End Sub
Sub LoadEXE(Dir As String)
    On Error GoTo err:
    X% = Shell(Dir, 1): NoFreeze% = DoEvents(): Exit Sub
    Exit Sub
err:
    'make your own error messages like mine
    '     below, or use the default:
    If err.Number = 6 Then Exit Sub
    'default: MsgBox "Error:" & vbCrLf & err
    '     .Description & vbCrLf & err.Number, vbEx
    '     clamation
    
End Sub
