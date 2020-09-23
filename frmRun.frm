VERSION 5.00
Begin VB.Form frmRun 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Run Program"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5145
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   5145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1230
      TabIndex        =   9
      Text            =   "Combo1"
      Top             =   1995
      Width           =   3735
   End
   Begin VB.TextBox txtpar 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   645
      TabIndex        =   7
      Top             =   1605
      Width           =   4290
   End
   Begin VB.CheckBox chkpar 
      Caption         =   "Uses Parameters."
      Height          =   195
      Left            =   675
      TabIndex        =   6
      Top             =   1335
      Width           =   1635
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Enabled         =   0   'False
      Height          =   390
      Left            =   960
      TabIndex        =   5
      Top             =   2550
      Width           =   1275
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "&Cancel"
      Height          =   390
      Left            =   2340
      TabIndex        =   4
      Top             =   2550
      Width           =   1275
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "&Browse"
      Height          =   390
      Left            =   3720
      TabIndex        =   3
      Top             =   2550
      Width           =   1275
   End
   Begin VB.TextBox txtfile 
      Height          =   315
      Left            =   675
      TabIndex        =   1
      Top             =   915
      Width           =   4245
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Window State"
      Height          =   195
      Left            =   165
      TabIndex        =   8
      Top             =   2070
      Width           =   1005
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Open"
      Height          =   195
      Left            =   165
      TabIndex        =   2
      Top             =   975
      Width           =   390
   End
   Begin VB.Label Label1 
      Caption         =   "Type in the name of the program in the textbox you want to run or use the browse button."
      Height          =   465
      Left            =   945
      TabIndex        =   0
      Top             =   210
      Width           =   3795
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   150
      Picture         =   "frmRun.frx":0000
      Top             =   165
      Width           =   480
   End
End
Attribute VB_Name = "frmRun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WState As Integer

Private Sub Check1_Click()
    
End Sub

Private Sub chkpar_Click()
    If chkpar Then
        txtpar.Enabled = True
    Else
        txtpar.Enabled = False
        txtpar = ""
    End If
    
End Sub

Private Sub cmdBrowse_Click()
Dim sFile As String
    sFile = Desk32.CDialogOpen("Open Program", "Program Files", "*.exe")
    If Len(sFile) > 0 Then
        txtfile.Text = sFile
    End If

End Sub

Private Sub cmdCan_Click()
    Unload frmRun
    
End Sub

Private Sub cmdOk_Click()
    If FindFile(txtfile) = False Then
        MsgBox "Can't find file " & txtfile.Text, vbCritical, "Error..."
        Exit Sub
    ElseIf chkpar Then
        OpenProgram hwnd, txtfile.Text & txtpar, WState
        Unload frmRun
        Exit Sub
        Else
        OpenProgram hwnd, txtfile.Text & txtpar, WState
        Unload frmRun
    End If
    
End Sub

Private Sub Combo1_Click()
    WState = Combo1.ListIndex
    
End Sub

Private Sub Form_Load()
    Icon = Nothing
    Combo1.AddItem "Hide Window"
    Combo1.AddItem "Normal Window"
    Combo1.AddItem "Minsized Window"
    Combo1.AddItem "Maxsized Window"
    Combo1.ListIndex = 1
    
End Sub

Private Sub txtfile_Change()
    If Len(txtfile) > 0 Then
        cmdOk.Enabled = True
    Else
        cmdOk.Enabled = False
    End If
    
End Sub
