VERSION 5.00
Begin VB.Form frmLogs 
   Caption         =   "로그"
   ClientHeight    =   5565
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4200
   Icon            =   "frmLogs.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5565
   ScaleWidth      =   4200
   StartUpPosition =   3  'Windows 기본값
   Begin VB.ListBox lvLogs 
      Height          =   5340
      IntegralHeight  =   0   'False
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "frmLogs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
    If Me.WindowState <> 1 Then
        lvLogs.Width = Me.Width - 345
        lvLogs.Height = Me.Height - 795
    End If
End Sub
