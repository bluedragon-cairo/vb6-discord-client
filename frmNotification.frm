VERSION 5.00
Begin VB.Form frmNotification 
   BorderStyle     =   3  '크기 고정 대화 상자
   ClientHeight    =   1605
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   Icon            =   "frmNotification.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Flasher 
      Interval        =   500
      Left            =   5640
      Top             =   1200
   End
   Begin VB.TextBox txtBody 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '없음
      Height          =   855
      Left            =   960
      MultiLine       =   -1  'True
      ScrollBars      =   2  '수직
      TabIndex        =   3
      Top             =   600
      Width           =   3615
   End
   Begin VB.Timer Timer 
      Interval        =   1000
      Left            =   5040
      Top             =   1080
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "닫기(&C)"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "보기(&V)"
      Default         =   -1  'True
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblHeader 
      Height          =   255
      Left            =   960
      TabIndex        =   2
      Top             =   240
      Width           =   3495
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "frmNotification.frx":000C
      Top             =   240
      Width           =   480
   End
End
Attribute VB_Name = "frmNotification"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
  ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, Optional ByVal dwDuration As Long = 250) As Long
Private Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal bInvert As Long) As Long
Private Const SND_APPLICATION = &H80         '  look for application specific association
Private Const SND_ALIAS = &H10000     '  name is a WIN.INI [sounds] entry
Private Const SND_ALIAS_ID = &H110000    '  name is a WIN.INI [sounds] entry identifier
Private Const SND_ASYNC = &H1         '  play asynchronously
Private Const SND_FILENAME = &H20000     '  name is a file name
Private Const SND_LOOP = &H8         '  loop the sound until next sndPlaySound
Private Const SND_MEMORY = &H4         '  lpszSoundName points to a memory file
Private Const SND_NODEFAULT = &H2         '  silence not default, if sound not found
Private Const SND_NOSTOP = &H10        '  don't stop any currently playing sound
Private Const SND_NOWAIT = &H2000      '  don't wait if the driver is busy
Private Const SND_PURGE = &H40               '  purge non-static events for task
Private Const SND_RESOURCE = &H40004     '  name is a resource name or atom
Private Const SND_SYNC = &H0         '  play synchronously (default)
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Dim t As Integer

Private Sub Flasher_Timer()
    FlashWindow Me.hwnd, 1
End Sub

Private Sub Form_Load()
    Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    t = 5
    Me.Hide
    Me.Top = Screen.Height - Me.Height - 777
    Me.Left = Screen.Width - Me.Width - 150
    Me.Visible = -1
    PlaySound Environ$("windir") & "\MEDIA\DING.WAV", ByVal 0&, SND_FILENAME Or SND_ASYNC
End Sub

Private Sub Timer_Timer()
    t = t - 1
    cmdClose.Caption = "닫기 (" & t & ")"
    If t <= 0 Then Unload Me
End Sub
