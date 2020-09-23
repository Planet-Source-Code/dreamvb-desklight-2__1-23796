Attribute VB_Name = "Desk32"
Declare Function GetLogicalDrives Lib "kernel32" () As Long
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOPENFILENAME As OPENFILENAME) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Declare Function SHShutDownDialog Lib "shell32" Alias "#60" (ByVal YourGuess As Long) As Long

Public DriveCount As Integer
Private Const KEYEVENTF_KEYUP = &H2
Private Const VK_LWIN = &H5B
Private Const GW_CHILD = 5
Private Const LVA_ALIGNLEFT = &H1
Private Const LVM_ARRANGE = &H1016

Private Type OPENFILENAME
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Enum TWindow
  nshow = 0
  nHide = 1
End Enum

Public Function ArrangeIcons()
Dim TDesk As Long
Dim THwnd As Long
Dim ret As Long

    TDesk = FindWindow("Progman", vbNullString)
    If TDesk = 0 Then Exit Function
    THwnd = GetWindow(TDesk, GW_CHILD)
    TDesk = GetWindow(THwnd, GW_CHILD)
    ret = SendMessage(TDesk, LVM_ARRANGE, LVA_ALIGNLEFT, 0)

End Function
Public Function LoadDrives()
Dim TDrive As Long
Dim TCount As Integer
Dim Drive_Letter As String

    TDrive = GetLogicalDrives
    
    For TCount = 0 To 25
        If (TDrive And 2 ^ TCount) <> 0 Then
            DriveCount = DriveCount + 1
            Drive_Letter = Chr(65 + TCount) & ":\"
            On Error Resume Next
            Load Form2.Label1(DriveCount)
            Load Form2.img(DriveCount)
            Form2.Label1(DriveCount).Top = Form2.Label1(DriveCount - 1).Top + Form2.Label1(DriveCount).Height
            Form2.img(DriveCount).Top = Form2.Label1(DriveCount - 1).Top + Form2.img(DriveCount).Height + 10
            
            Select Case GetDriveType(Drive_Letter)
                Case 2
                   Form2.img(DriveCount).Picture = Form2.res1(0).Picture
                Case 3
                   Form2.img(DriveCount).Picture = Form2.res1(1).Picture
                Case 5
                    Form2.img(DriveCount).Picture = Form2.res1(2).Picture
            End Select
            
            Form2.Label1(DriveCount).Width = Form2.ScaleWidth - 1
            Form2.Label1(DriveCount).Tag = Chr(65 + TCount) & ":\"
            Form2.Label1(DriveCount).Caption = "Browse (" & Chr(65 + TCount) & ":\" & ")"
            Form2.Label1(DriveCount).Visible = True
            Form2.img(DriveCount).Visible = True
        End If
    Next
    
End Function
Public Function GetWinPath() As String
Dim StrWinPath As String
    StrWinPath = String(255, Chr(0))
    GetWindowsDirectory StrWinPath, 255
    StrWinPath = Left(StrWinPath, InStr(StrWinPath, Chr(0)) - 1)
    GetWinPath = StrWinPath
    StrWinPath = ""
    
End Function
Public Function MinsizeAllWindows()
    Call keybd_event(VK_LWIN, 0, 0, 0)
    Call keybd_event(77, 0, 0, 0)
    Call keybd_event(VK_LWIN, 0, KEYEVENTF_KEYUP, 0)
End Function
Public Function FindFile(lzFilename As String) As Boolean
    If Dir(lzFilename) = "" Then FindFile = False Else FindFile = True
    
End Function
Public Function OpenProgram(mHwnd As Long, ProgramNamePath As String, ByVal ShowWindow As VbAppWinStyle)
    ShellExecute mHwnd, vbNullString, ProgramNamePath, vbNullString, vbNullString, ShowWindow
    
End Function
Public Function CDialogOpen(mTitle, mFileType, mFileExt As String) As String
 Dim ofn As OPENFILENAME
    ofn.lStructSize = Len(ofn)
    ofn.hWndOwner = Form1.hwnd
    ofn.hInstance = App.hInstance
    ofn.lpstrFilter = mFileType + Chr$(0) + mFileExt
        ofn.lpstrFile = Space$(254)
        ofn.nMaxFile = 255
        ofn.lpstrFileTitle = Space$(254)
        ofn.nMaxFileTitle = 255
        ofn.lpstrInitialDir = App.Path
        ofn.lpstrTitle = mTitle
        ofn.flags = 0
       
        A = GetOpenFileName(ofn)
        If (A) Then
                CDialogOpen = Trim$(ofn.lpstrFile)
        End If
        
 End Function

Public Function HideDeskTopIcons(mShow As TWindow)
Dim TDeskicon As Long
    TDeskicon = FindWindow("Progman", vbNullString)
    If TDeskicon <> 0 Then
        TDeskicon = ShowWindow(TDeskicon, mShow)
    End If
    
 End Function
Public Function HideTaskBarA(mShow As TWindow)
Dim hTBar As Long
    hTBar = FindWindow("Shell_traywnd", vbNullString)
    If hTBar <> 0 Then
        ShowWindow hTBar, mShow
    End If
    
End Function

