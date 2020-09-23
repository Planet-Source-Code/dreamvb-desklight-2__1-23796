VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   2325
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   2325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Project1.Line3D Line3D1 
      Height          =   105
      Left            =   -90
      TabIndex        =   12
      Top             =   1275
      Width           =   2385
      _ExtentX        =   4207
      _ExtentY        =   185
   End
   Begin Project1.Line3D Line3D4 
      Height          =   105
      Left            =   -30
      TabIndex        =   13
      Top             =   3690
      Width           =   2385
      _ExtentX        =   4207
      _ExtentY        =   185
   End
   Begin Project1.Line3D Line3D3 
      Height          =   105
      Left            =   15
      TabIndex        =   14
      Top             =   2340
      Width           =   2385
      _ExtentX        =   4207
      _ExtentY        =   185
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   12
      Left            =   45
      Picture         =   "Form1.frx":0000
      Top             =   4155
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   11
      Left            =   60
      Picture         =   "Form1.frx":014A
      Top             =   1740
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   10
      Left            =   45
      Picture         =   "Form1.frx":0294
      Top             =   3360
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   9
      Left            =   45
      Picture         =   "Form1.frx":03DE
      Top             =   3060
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   8
      Left            =   30
      Picture         =   "Form1.frx":0528
      Top             =   2760
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   7
      Left            =   60
      Picture         =   "Form1.frx":0672
      Top             =   2460
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   6
      Left            =   45
      Picture         =   "Form1.frx":07BC
      Top             =   3840
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   5
      Left            =   60
      Picture         =   "Form1.frx":0906
      Top             =   2025
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   4
      Left            =   60
      Picture         =   "Form1.frx":0A50
      Top             =   1455
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   3
      Left            =   60
      Picture         =   "Form1.frx":0B9A
      Top             =   990
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   2
      Left            =   60
      Picture         =   "Form1.frx":0CE4
      Top             =   690
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   1
      Left            =   60
      Picture         =   "Form1.frx":0E2E
      Top             =   390
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   0
      Left            =   60
      Picture         =   "Form1.frx":0F78
      Top             =   75
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "        Browse"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   11
      Left            =   15
      TabIndex        =   11
      Top             =   1740
      Width           =   2265
   End
   Begin VB.Label Label1 
      Caption         =   "        MS-DOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   10
      Left            =   15
      TabIndex        =   10
      Top             =   3345
      Width           =   2265
   End
   Begin VB.Label Label1 
      Caption         =   "        MS Paint"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   9
      Left            =   15
      TabIndex        =   9
      Top             =   3045
      Width           =   2265
   End
   Begin VB.Label Label1 
      Caption         =   "        Notepad"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   8
      Left            =   15
      TabIndex        =   8
      Top             =   2760
      Width           =   2265
   End
   Begin VB.Label Label1 
      Caption         =   "        Reg-Edit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   7
      Left            =   15
      TabIndex        =   7
      Top             =   2475
      Width           =   2265
   End
   Begin VB.Label Label1 
      Caption         =   "        About DeskLight ver 2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   6
      Left            =   15
      TabIndex        =   6
      Top             =   3840
      Width           =   2265
   End
   Begin VB.Label Label1 
      Caption         =   "        Shut Down Windows"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   5
      Left            =   15
      TabIndex        =   5
      Top             =   2025
      Width           =   2265
   End
   Begin VB.Label Label1 
      Caption         =   "        Run a Program"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   4
      Left            =   15
      TabIndex        =   4
      Top             =   1440
      Width           =   2265
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000A&
      Caption         =   "        Minsize All Windows"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   3
      Left            =   15
      TabIndex        =   3
      Top             =   975
      Width           =   2265
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000A&
      Caption         =   "        Arrange Icons"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   2
      Left            =   15
      TabIndex        =   2
      Top             =   690
      Width           =   2265
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000A&
      Caption         =   "        Hide Task Bar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   1
      Left            =   15
      TabIndex        =   1
      Top             =   390
      Width           =   2265
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000A&
      Caption         =   "        Hide Desktop Icons"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   0
      Left            =   15
      TabIndex        =   0
      Top             =   60
      Width           =   2265
   End
   Begin VB.Label Label1 
      Caption         =   "        Close DeskLight"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   12
      Left            =   15
      TabIndex        =   15
      Top             =   4155
      Width           =   2265
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim HideIcons As Boolean
Dim HideTask As Boolean

Sub HideDeskIcons()
    Select Case HideIcons
        Case True
            Label1(0).Caption = "        Hide Desktop Icons"
            HideDeskTopIcons nHide
            HideIcons = False
        Case False
            Label1(0).Caption = "        Show Desktop Icons"
            HideDeskTopIcons nshow
            HideIcons = True
        End Select

End Sub
Sub HideTaskBar()
    Select Case HideTask
        Case True
            Label1(1).Caption = "        Hide Task Bar"
            HideTask = False
            HideTaskBarA nHide
        Case False
            Label1(1).Caption = "        Show Task Bar"
            HideTask = True
            HideTaskBarA nshow
        End Select
        
End Sub
Private Sub Form_Load()
Dim DeskWidth, DeskHeight As Long
Dim TotalSize As String
LoadDrives
    DeskWidth = (Screen.Width / Screen.TwipsPerPixelX)
    DeskHeight = (Screen.Height / Screen.TwipsPerPixelY)
    
    Form1.Left = Screen.Width - Form1.Width - 30
    Form1.Top = Screen.Height - Form1.Height - 400
    HideIcons = False
    
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    For i = 0 To 12
        Label1(i).BackColor = &H8000000F
        Label1(i).ForeColor = vbBlack
    Next
    
End Sub

Private Sub Label1_Click(Index As Integer)
    Select Case Index
        Case 0
            HideDeskIcons
        Case 1
            HideTaskBar
        Case 2
            ArrangeIcons
        Case 3
            Desk32.MinsizeAllWindows
        Case 4
            frmRun.Show
        Case 5
            SHShutDownDialog 0
        Case 6
            frmabout.Show
            
        Case 7
            If FindFile(GetWinPath & "\regedit.exe") = False Then
                MsgBox "Can't find regedit on your system", vbCritical, "Error"
            Else
                OpenProgram hwnd, GetWinPath & "\regedit.exe", vbNormalFocus
            End If
        Case 8
            If FindFile(GetWinPath & "\Notepad.exe") = False Then
                MsgBox "Can't find notepad on your system", vbCritical, "Error"
            Else
                OpenProgram hwnd, GetWinPath & "\notepad.exe", vbNormalFocus
            End If
        Case 9
            If FindFile(GetWinPath & "\pbrush.exe") = False Then
                MsgBox "Can't find MS Paint on your system", vbCritical, "Error"
            Else
                OpenProgram hwnd, GetWinPath & "\pbrush.exe", vbNormalFocus
            End If
        Case 10
                OpenProgram hwnd, "command.com", vbNormalFocus
        Case 12
            Unload Form1: End
    End Select
    
    
End Sub

Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer
    For i = 0 To 12
        Label1(i).BackColor = &H8000000F
        Label1(i).ForeColor = vbBlack
    Next
        Label1(Index).BackColor = vbBlue
        Label1(Index).ForeColor = vbWhite
        If Index = 11 Then
            Form2.Left = Me.Left - Form2.Width
            Form2.Top = Label1(11).Top + Form1.Top
            Form2.Height = Form2.Label1(DriveCount).Top + Form2.Label1(DriveCount).Height + 100
            Form2.Show
        Else
            Form2.Visible = False
        End If
        
End Sub


Private Sub Label1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
        Label1(Index).BackColor = &H8000000F
        Label1(Index).ForeColor = vbBlack
        
End Sub
