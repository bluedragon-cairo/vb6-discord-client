VERSION 5.00
Begin VB.Form frmAuthentication 
   BorderStyle     =   3  'ũ�� ���� ��ȭ ����
   Caption         =   "���� ��ȯ"
   ClientHeight    =   3420
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6285
   Icon            =   "frmAuthentication.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   6285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '������ ���
   Begin VB.CommandButton cmdRegister 
      Caption         =   "ȸ������(&R)..."
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   14
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)..."
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   340
      Left            =   3120
      TabIndex        =   13
      Top             =   1060
      Width           =   1335
   End
   Begin VB.CheckBox chkBot 
      Caption         =   "�� ����(&B)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1080
      Width           =   1575
   End
   Begin VB.OptionButton rEmail 
      Caption         =   "���� �α���(&P)"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   6015
      Begin VB.TextBox txtPassword 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         IMEMode         =   3  '��� ����
         Left            =   360
         PasswordChar    =   "*"
         TabIndex        =   12
         Top             =   1320
         Width           =   3975
      End
      Begin VB.TextBox txtEmail 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   360
         TabIndex        =   10
         Top             =   600
         Width           =   3975
      End
      Begin VB.Label Label3 
         Caption         =   "��й�ȣ(&A):"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "���ڿ��� �ּ�(&E):"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.OptionButton rToken 
      Caption         =   "API ��ū(&T)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Value           =   -1  'True
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4455
      Begin VB.TextBox txtToken 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   360
         TabIndex        =   6
         Top             =   600
         Width           =   3975
      End
      Begin VB.Label Label2 
         Caption         =   "��ū(Bearer �� Bot ���λ� ����)(&O):"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   3135
      End
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "�α���(&L)"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmAuthentication"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Http As New WinHttp.WinHttpRequest

Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub cmdRegister_Click()
    Shell "explorer.exe https://discord.com/register"
End Sub

Private Sub Form_Load()
    EnableTLS Http
    txtToken.Text = Token
    If Left$(txtToken.Text, 4) = "Bot " Then
        txtToken.Text = Right$(txtToken.Text, Len(txtToken.Text) - 4)
        chkBot.Value = 1
    End If
End Sub

Private Sub OKButton_Click()
    If chkBot.Value Then txtToken.Text = "Bot " & txtToken.Text
    
    Http.Open "GET", "https://discord.com/api/v8/users/@me", False
    Http.SetRequestHeader "Content-Type", "application/json"
    Http.SetRequestHeader "Authorization", txtToken.Text
    Http.SetRequestHeader "User-Agent", "My XML App V1.0"
    Http.Send
    
    Set p = JSON.parse(CStr(Http.ResponseText))
    If Http.Status >= 400 Then
        If Left$(txtToken.Text, 4) = "Bot " Then
            txtToken.Text = Right$(txtToken.Text, Len(txtToken.Text) - 4)
            chkBot.Value = 1
        Else
            chkBot.Value = 0
        End If
        MsgBox "�α��� ������ ��ū�� �߸��Ǿ����ϴ�. (���� �ڵ� " & p("code") & ")", 16, "�α��� ����"
        Exit Sub
    End If
    
    SaveSetting "VB6 Discord", "Authorization", "Token", txtToken.Text
    Token = txtToken.Text
    Unload Me
End Sub
