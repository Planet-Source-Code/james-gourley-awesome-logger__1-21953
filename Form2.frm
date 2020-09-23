VERSION 5.00
Begin VB.Form Form2 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Logger"
   ClientHeight    =   2520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3744
   LinkTopic       =   "Form2"
   ScaleHeight     =   2520
   ScaleWidth      =   3744
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   2508
      Left            =   0
      ScaleHeight     =   2508
      ScaleWidth      =   3756
      TabIndex        =   0
      Top             =   0
      Width           =   3756
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Picture1_Click()
KeepOffTop Form2
Form2.Hide
Form1.Show
End Sub

