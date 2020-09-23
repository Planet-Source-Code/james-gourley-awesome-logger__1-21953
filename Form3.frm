VERSION 5.00
Object = "{6FBA474E-43AC-11CE-9A0E-00AA0062BB4C}#1.0#0"; "SYSINFO.OCX"
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "mIRC Password"
   ClientHeight    =   1356
   ClientLeft      =   36
   ClientTop       =   300
   ClientWidth     =   3768
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1356
   ScaleWidth      =   3768
   StartUpPosition =   3  'Windows Default
   Begin SysInfoLib.SysInfo SysInfo1 
      Left            =   -12
      Top             =   0
      _ExtentX        =   804
      _ExtentY        =   804
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   372
      Left            =   2748
      TabIndex        =   4
      Top             =   888
      Width           =   852
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   384
      Left            =   2736
      TabIndex        =   3
      Top             =   456
      Width           =   852
   End
   Begin VB.TextBox Text1 
      Height          =   288
      IMEMode         =   3  'DISABLE
      Left            =   1020
      MaxLength       =   28
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   528
      Width           =   1524
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "&Password:"
      Height          =   240
      Left            =   96
      TabIndex        =   2
      Top             =   576
      Width           =   888
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Please enter your password"
      Height          =   240
      Left            =   -24
      TabIndex        =   0
      Top             =   96
      Width           =   3804
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1 = "runmirc" Then
LoadEXE App.Path & "\mIRC32.exe"
Unload Form3
Exit Sub
End If
If Text1 = "runlogger" Then
Form1.Show
Unload Form3
Else
Beep
Text1 = ""
End If
End Sub

Private Sub Command2_Click()
Unload Me
Form1.YEs = True
Unload Form1
Unload Form2
End Sub

Private Sub Form_Load()
Me.Left = SysInfo1.WorkAreaWidth / 2 - Me.Width / 2
Me.Top = SysInfo1.WorkAreaHeight / 2 - Me.Height / 2
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Command1_Click
KeyAscii = 0
End If
End Sub
