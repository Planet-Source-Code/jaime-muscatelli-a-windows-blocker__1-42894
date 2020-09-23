VERSION 5.00
Begin VB.Form FRMBLOCK 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   6045
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7095
   LinkTopic       =   "Form1"
   ScaleHeight     =   6045
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox piclog 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H000000FF&
      Height          =   3975
      Left            =   120
      ScaleHeight     =   3975
      ScaleWidth      =   4935
      TabIndex        =   4
      Top             =   1920
      Width           =   4935
   End
   Begin VB.Label lblcaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Log:"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   360
      TabIndex        =   3
      Top             =   1320
      Width           =   315
   End
   Begin VB.Label lbl2 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "By Jaime Muscatelli"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label lbltime 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2640
      TabIndex        =   1
      Top             =   840
      Width           =   45
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Windows has been blocked!"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   2280
      TabIndex        =   0
      Top             =   240
      Width           =   2025
   End
End
Attribute VB_Name = "FRMBLOCK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Me.Show
lbltime.Caption = "Logged in @ " & Time
piclog.Height = Me.Height
FRMDIALOG.Show vbModal
End Sub

