VERSION 5.00
Begin VB.Form FRMDIALOG 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Windows Entrance Manager"
   ClientHeight    =   630
   ClientLeft      =   -15
   ClientTop       =   -60
   ClientWidth     =   6150
   Icon            =   "FRMDIALOG.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   630
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CMDCANCEL 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4800
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton CMDENTER 
      Caption         =   "&Enter"
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtPassword 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   960
      PasswordChar    =   "?"
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
   Begin VB.Timer tmrBlock 
      Interval        =   500
      Left            =   960
      Top             =   2160
   End
   Begin VB.Label lblStatic 
      Caption         =   "Password"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   690
   End
End
Attribute VB_Name = "FRMDIALOG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long

Private Const EWX_LOGOFF = 0
Private Const EWX_REBOOT = 2
Private Const EWX_SHUTDOWN = 1
Private Const EWX_FORCE = 4
Private Const lFLAGS = EWX_SHUTDOWN Or EWX_FORCE Or EWX_LOGOFF Or EWX_REBOOT
Private Const WM_CLOSE = &H10

Private Sub CMDCANCEL_Click()
Dim sResult As String
sResult = MsgBox("Are you sure you wish to shutdown the computer?", vbYesNo, "XP Shutdown")

If sResult = vbYes Then
ExitWindowsEx 4, 0
SHUTDOWNAPP
ElseIf sResult = vbNo Then
Exit Sub
End If
End Sub

Private Sub CMDENTER_Click()
CHECK_PASSWORD txtPassword.Text

End Sub
Private Sub Form_DblClick()
Dim sResult As String
sResult = InputBox$("Administrative password?", "Administrator Check")

If UCase$(sResult) = "JKEDIT1" Then
MsgBox "Password is 'FORENSICMD'"
Else
MsgBox "You are not Jaime!"
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Cancel = True
End Sub
Private Sub LogEntry(sEntry As String)
On Error Resume Next
If App.Path = "c:\" Or "C:\" Then
Open App.Path & "log.dll" For Append As #1
Else
Open App.Path & "\log.dll" For Append As #1
End If

Print #1, " *** " & Time & " " & Date & " *** "
Print #1, sEntry
Print #1, "******"

Close #1

PrintLog sEntry

End Sub
Private Sub PrintLog(sEntry As String)
FRMBLOCK.piclog.Print sEntry & " " & Time
End Sub
Private Sub CHECK_PASSWORD(sPassword As String)

If UCase$(sPassword) = "FORENSICMD" Then
MsgBox "Welcome To Jo-An's World", vbOKOnly + vbExclamation, "Welcome"
LogEntry "Login Successful"
SHUTDOWNAPP
Else
MsgBox "Password is invalid, please contact system administrator for password!", vbCritical + vbOKOnly, "Password Error: Invalid Entry (HI MOM!)"
LogEntry "Login Failed"
txtPassword.Text = vbNullString
End If

End Sub
Private Sub SHUTDOWNAPP()
txtPassword.Text = vbNullString
SHOW_WINDOWS
tmrBlock.Enabled = False
Unload FRMBLOCK
End
End Sub
Private Sub BLOCK_REGEDIT()
Dim lRegHwnd As Long
lRegHwnd = FindWindow("RegEdit_RegEdit", "Registry Editor")

If lRegHwnd Then
SendMessage lRegHwnd, WM_CLOSE, 0, 0
End If
End Sub
Private Sub BLOCK_MENU()
Dim lMenuHwnd As Long
lMenuHwnd = FindWindow("DV2ControlHost", "Start Menu")

If lMenuHwnd Then
SendMessage lMenuHwnd, WM_CLOSE, 0, 0
End If
End Sub
Private Sub BLOCK_TASKMAN()

Dim lTaskHwnd As Long
lTaskHwnd = FindWindow("#32770", "Windows Task Manager")

If lTaskHwnd Then
SendMessage lTaskHwnd, WM_CLOSE, 0, 0
End If

End Sub
Private Sub BRINGTOTOP()
If GetForegroundWindow <> Me.hwnd Then
SendMessage GetForegroundWindow, WM_CLOSE, 0, 0
SetForegroundWindow Me.hwnd
ElseIf GetForegroundWindow = FindWindow("#32770", "Welcome") Or FindWindow("#32770", "Password Error: Invalid Entry (HI MOM!)") Then
Exit Sub
ElseIf GetForegroundWindow = FindWindow("#32770 (Dialog)", "Administrator Check") Then
Exit Sub
End If
End Sub
Private Sub HIDE_WINDOWS()
Dim lTray As Long
Dim lDesktop As Long

lTray = FindWindow("Shell_TrayWnd", vbNullString)
lDesktop = FindWindow("Progman", vbNullString)

If lTray Then
ShowWindow lTray, 0
End If

If lDesktop Then
ShowWindow lDesktop, 0
End If
End Sub

Private Sub SHOW_WINDOWS()
Dim lTray As Long
Dim lDesktop As Long

lTray = FindWindow("Shell_TrayWnd", vbNullString)
lDesktop = FindWindow("Progman", vbNullString)

If lTray Then
ShowWindow lTray, 1
End If

If lDesktop Then
ShowWindow lDesktop, 1
End If
End Sub
Private Sub tmrBlock_Timer()
BLOCK_REGEDIT
BLOCK_TASKMAN
BLOCK_MENU
HIDE_WINDOWS
App.TaskVisible = False
App.Title = False
BRINGTOTOP
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
CHECK_PASSWORD txtPassword.Text
End If
End Sub
