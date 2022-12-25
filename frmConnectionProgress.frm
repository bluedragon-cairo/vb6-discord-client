VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmConnectionProgress 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "연결 중..."
   ClientHeight    =   1335
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6045
   ControlBox      =   0   'False
   Icon            =   "frmConnectionProgress.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   6045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin ComctlLib.ProgressBar pbStatus 
      Height          =   240
      Left            =   960
      TabIndex        =   0
      Top             =   840
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   423
      _Version        =   327682
      Appearance      =   0
   End
   Begin VB.Label lblStatus 
      Caption         =   "연결 중..."
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   360
      Width           =   4935
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "frmConnectionProgress.frx":000C
      Top             =   240
      Width           =   480
   End
End
Attribute VB_Name = "frmConnectionProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'https://www.vbforums.com/showthread.php?462116-Marquee-Progressbar
Private Const WM_USER = &H400
Private Const PBM_SETMARQUEE = WM_USER + 10
Private Const GWL_STYLE = (-16)
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long

Private Sub Form_Load()
    SetWindowLong pbStatus.hwnd, GWL_STYLE, GetWindowLong(pbStatus.hwnd, GWL_STYLE) Or PBM_SETMARQUEE
    SendMessage pbStatus.hwnd, PBM_SETMARQUEE, 1, CLng(20)
End Sub
