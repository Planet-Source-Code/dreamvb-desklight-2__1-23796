VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   510
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   1590
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   510
   ScaleWidth      =   1590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Image res1 
      Height          =   240
      Index           =   2
      Left            =   720
      Picture         =   "Form2.frx":0000
      Top             =   2250
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image res1 
      Height          =   240
      Index           =   1
      Left            =   480
      Picture         =   "Form2.frx":014A
      Top             =   2235
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image res1 
      Height          =   240
      Index           =   0
      Left            =   75
      Picture         =   "Form2.frx":0294
      Top             =   2220
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image img 
      Height          =   240
      Index           =   1
      Left            =   45
      Picture         =   "Form2.frx":03DE
      Top             =   30
      Width           =   240
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000A&
      Caption         =   "        Drive A:\"
      Height          =   270
      Index           =   1
      Left            =   330
      TabIndex        =   0
      Top             =   15
      Width           =   1200
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_Click(Index As Integer)
    OpenProgram hwnd, Label1(Index).Tag, vbNormalFocus
    Form2.Visible = False
    
End Sub

Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer
    For i = 1 To DriveCount
        Label1(i).BackColor = &H8000000F
        Label1(i).ForeColor = vbBlack
    Next
        Label1(Index).BackColor = vbBlue
        Label1(Index).ForeColor = vbWhite
        
End Sub
