VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmMain 
   Caption         =   "VBWebsocket Example "
   ClientHeight    =   4395
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9630
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "frmMain"
   MaxButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   9630
   StartUpPosition =   3  'Windows 기본값
   Begin ComctlLib.TreeView tvGuilds 
      Height          =   3735
      Left            =   120
      TabIndex        =   11
      Top             =   480
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   6588
      _Version        =   327682
      HideSelection   =   0   'False
      Indentation     =   471
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   6
      Appearance      =   1
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   975
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1720
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   2
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "친구"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "채팅"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   ">>"
      Height          =   375
      Left            =   7320
      TabIndex        =   15
      Top             =   3840
      Width           =   615
   End
   Begin VB.TextBox txtSendMessage 
      Height          =   375
      Left            =   2640
      MaxLength       =   1999
      TabIndex        =   14
      Top             =   3840
      Width           =   4575
   End
   Begin VB.CommandButton cmdAttach 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   13
      Top             =   3840
      Width           =   375
   End
   Begin SHDocVwCtl.WebBrowser wbChat 
      Height          =   3615
      Left            =   2160
      TabIndex        =   12
      Top             =   120
      Width           =   5775
      ExtentX         =   10186
      ExtentY         =   6376
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.CommandButton cmdToken 
      Caption         =   "토큰"
      Height          =   375
      Left            =   8040
      TabIndex        =   10
      Top             =   720
      Width           =   1455
   End
   Begin VB.Timer timHeartbeat 
      Enabled         =   0   'False
      Interval        =   45000
      Left            =   8520
      Top             =   2280
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Ping Server"
      Height          =   540
      Left            =   9840
      TabIndex        =   6
      Top             =   1680
      Width           =   1440
   End
   Begin VB.ListBox List2 
      Height          =   780
      Left            =   8280
      TabIndex        =   5
      Top             =   7080
      Width           =   2895
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Send"
      Height          =   540
      Left            =   10080
      TabIndex        =   4
      Top             =   6480
      Width           =   1065
   End
   Begin VB.TextBox Text3 
      Height          =   465
      Left            =   8280
      TabIndex        =   3
      Text            =   "Hello World"
      Top             =   6480
      Width           =   1665
   End
   Begin VB.ListBox List1 
      Height          =   2580
      Left            =   3240
      TabIndex        =   2
      Top             =   6600
      Width           =   2865
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Disconnect"
      Height          =   540
      Left            =   9720
      TabIndex        =   1
      Top             =   960
      Width           =   1440
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Connect"
      Height          =   540
      Left            =   8040
      TabIndex        =   0
      Top             =   120
      Width           =   1440
   End
   Begin VB6_Discord.Websocket ws 
      Height          =   465
      Left            =   8040
      Top             =   1200
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   820
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '투명
      Caption         =   "Click Here To Support Humanitarian Aid to Ukraine"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   12225
      MouseIcon       =   "frmMain.frx":030A
      MousePointer    =   99  '사용자 정의
      TabIndex        =   9
      Top             =   3975
      Width           =   3765
   End
   Begin VB.Label lblUKRAINE 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00B85E00&
      Caption         =   "SLAVA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0001D1FF&
      Height          =   1815
      Index           =   0
      Left            =   11475
      TabIndex        =   7
      Top             =   225
      Width           =   5415
   End
   Begin VB.Label lblUKRAINE 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H0001D1FF&
      Caption         =   "UKRAYINI"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00B85E00&
      Height          =   1815
      Index           =   1
      Left            =   11475
      TabIndex        =   8
      Top             =   2025
      Width           =   5415
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Websocket sample by lewis (dethbomb@hotmail.com)

'This project is intended to show the basics of how to use the websocket.
'error handling is omitted in order to keep the code more clear.
'added a btc ticker example using the coinbase pro websocket server

'see the included html help page for more info on how to use the websocket

'DISCLAIMER:
'THE SOFTWARE (AND SOURCE CODE) IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
'IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
'FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
'AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
'LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
'OUT OF OR IN CONNECTION WITH THE SOFTWARE OR SOURCE CODE, OR THE USE OR OTHER DEALINGS IN THE
'SOFTWARE OR SOURCE CODE.

'flag to let onMessage know we are in demo mode
Dim BtcDemo As Boolean
Dim HeartbeatInterval As Long
Dim sequence As Long
Dim sessionID As String
Dim ResumeURL As String
Dim Resuming As Boolean
Dim users As New Dictionary
Dim channels As New Dictionary
Dim guilds As New Dictionary

Sub ResumeDiscord()
    Dim Headers As New Collection
    If ws.readyState <> STATE_CLOSED Then
        ws.Disconnect
    End If
    
    Resuming = True
    Log ResumeURL & "으로 다시 연결합니다..."
    ws.Connect ResumeURL, Text2, Text5, "", Headers
End Sub

Private Sub cmdSend_Click()
    If Not channels.Exists(tvGuilds.SelectedItem.key) Then Exit Sub
    If channels(tvGuilds.SelectedItem.key)("type") <> 0 Then Exit Sub
    
    Dim Http As New WinHttp.WinHttpRequest
    EnableTLS Http
    
    Http.Open "POST", "https://discord.com/api/v10/channels/" & tvGuilds.SelectedItem.key & "/messages", True
    Http.SetRequestHeader "Content-Type", "application/json"
    Http.SetRequestHeader "Authorization", Token
    Http.SetRequestHeader "User-Agent", "My XML App V1.0"
    Http.Send "{""content"":""" & Escape(txtSendMessage.Text) & """}"
    Http.WaitForResponse 60
    
    If Http.Status >= 400 Then
        MsgBox "메시지를 보낼 수 없습니다. (HTTP " & Http.Status & ")", 16, "오류"
        Exit Sub
    End If
End Sub

'connect to the websocket server
Private Sub Command1_Click()
    Dim Headers As New Collection
    
    If ws.readyState <> STATE_CLOSED Then
        ws.Disconnect
    End If
    
    ws.UseCompression = False
    ws.ChunkSize = 4096

    Log "wss://gateway.discord.gg/?v=10&encoding=json" & "에 연결하는 중..."
    frmConnectionProgress.Show '1, Me
    frmConnectionProgress.lblStatus.Caption = "디스코드 서버에 연결하는 중..."
    ws.Connect "wss://gateway.discord.gg/?v=10&encoding=json", "443", "", "", Headers
End Sub



Private Sub cmdToken_Click()
    frmAuthentication.Show 1, Me
End Sub

'disconnect from the websocket server
Private Sub Command2_Click()

    If ws.readyState <> STATE_CLOSED Then
        ''close' is a VB reserved keyword so we use 'disconnect'
        ws.Disconnect
    End If

End Sub


'send chat/data to the websocket server
Private Sub Command3_Click()
    Dim Bytes() As Byte
    Dim aBinaryString As String

    If ws.readyState = STATE_OPEN Then

        If Len(Text3.Text) Then
            ws.SendAdvanced Text3.Text, 1, True, False, True, False, False, False

            AddListItem List1, "Myself: " & Text3
            Text3 = ""
        End If

    Else
        MsgBox "VBWebsocket is NOT connected", vbCritical
    End If

End Sub

'start piesocket demo
Private Sub Command4_Click()

    Dim NewPiesocketClient As Form
    
    'spawn at least 4 clients the first time
    Do
        Set NewPiesocketClient = New frmPieSocket
        NewPiesocketClient.Show , Me
    Loop While Forms.Count < 5

End Sub

'echo server demo
Private Sub Command5_Click()
    BtcDemo = False
    Text1 = "wss://echo.websocket.events/.ws"
    Text4(0) = ""
    Text4(1) = ""
    Text5 = ""
    Check2.Value = 0
    Check1.Value = 0
    MsgBox "All values for echo websocket demo server have been set. To try the demo please press connect.", vbInformation
End Sub


'start btc ticker
Private Sub Command6_Click()

    'already running
    If BtcDemo = True Then Exit Sub

    BtcDemo = True

    Text1 = "wss://ws-feed.pro.coinbase.com/"

    'make sure all values are clear

    'no headers
    Text4(0) = ""
    Text4(1) = ""

    'no sub protocol
    Text5 = ""

    'no compression
    Check2.Value = 0

    'send as text
    Check1.Value = 0

    'hit the connect button
    Command1_Click

End Sub

'stop btc demo
Private Sub Command7_Click()

    BtcDemo = False

    List3.List(0) = "BTC SELL:"
    List3.List(1) = "BTC BUY:"

    'hit disconnect button
    Command2_Click

End Sub


'view the help file
Private Sub Command8_Click()

    Shell "explorer " & """" & App.Path & "\HELP.html" & """", vbNormalFocus

End Sub

'visit forum post regarding vbwebsocket
Private Sub Command9_Click()

    Shell "explorer " & """" & "https://www.vbforums.com/showthread.php?892835-VB6-Visual-Basic-6-Client-Websocket-Control" & """", vbNormalFocus

End Sub

'ping the server
Private Sub Command10_Click()
    If ws.readyState = STATE_OPEN Then
        AddListItem List1, "Me: Ping"
        ws.Ping
    End If
End Sub


'open the rfc for webwocket protocol
Private Sub Command11_Click()

    Dim FSO As Object

    Set FSO = CreateObject("scripting.filesystemobject")
    If FSO.FileExists(App.Path & "\rfc6455.txt") Then
        Shell "explorer " & """" & App.Path & "\rfc6455.txt" & """", vbNormalFocus
    Else
        Shell "explorer " & """" & "https://www.rfc-editor.org/rfc/rfc6455.txt" & """", vbNormalFocus
    End If
End Sub




Private Sub Form_Load()
    Load frmLogs
    frmLogs.Show
    Token = GetSetting("VB6 Discord", "Authorization", "Token", "")
    SetFont Me
    Me.Caption = App.Title
    wbChat.Navigate "about:blank"
    wbChat.Document.parentWindow.execScript "window.createMessage = function createMessage(sender, content) { document.getElementById('message-container').innerHTML += '<table><tr><td rowspan=2><img /></td><td>' + sender.username + '</td></tr><tr><td>' + content + '</td></tr></table>'; }"
    wbChat.Document.parentWindow.execScript "window.onload = function() { document.body.style.backgroundColor = 'rgb(41, 57, 69)'; document.body.style.color = 'white'; document.body.style.fontFamily = '돋움,Dotum'; document.body.innerHTML = '<div id=message-container></div>'; }"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'make sure all the pie socket forms are unloaded

    Dim F As Form

    For Each F In Forms
        If Not (F Is Me) Then
            Unload F
            Set F = Nothing
        End If
    Next

End Sub

'show freedom!
Private Sub Label4_Click()
    Me.Width = 17325
End Sub

'help ukraine
Private Sub Label7_Click()
    Shell "explorer " & """" & "https://www.timeout.com/news/22-ways-you-can-help-the-people-of-ukraine-right-now-032322" & """"
End Sub

'==================================================================
' misc GUI functions
'==================================================================

'set list1 tooltip to selected text
Private Sub List1_Click()

    If List1.ListIndex <> -1 Then
        List1.ToolTipText = List1.List(List1.ListIndex)
    End If

End Sub

'set list2 tooltip to selected text
Private Sub List2_Click()
    With List2
        If .ListIndex <> -1 Then
            .ToolTipText = .List(.ListIndex)
        End If
    End With
End Sub

'change port automatically depending on url
Private Sub Text1_Change()
    If Len(Text1) > 3 Then
        If Left$(Text1, 3) = "wss" Then
            Text2 = "443"
        ElseIf Left$(Text1, 2) = "ws" Then
            Text2 = "80"
        End If
    End If
End Sub


'if user hits enter key, send chat text
Private Sub Text3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0  'eliminates the windows DING sound
        Command3_Click
    End If
End Sub

'add an item to the list and scroll
Sub AddListItem(lstBox As ListBox, ByVal newVal As String)

    lstBox.AddItem newVal
    lstBox.ListIndex = lstBox.NewIndex    'scroll to new entry

End Sub

Private Sub timHeartbeat_Timer()
    Dim Bytes() As Byte
    Dim aBinaryString As String

    If ws.readyState = STATE_OPEN Then
        'send as text?
        'default is text (aka string)
        'ws.Send Text3.Text
        'here is how to use the new SendAdvanced() function, you could also use opText instead of 1
        ws.Send "{""op"":1,""d"":" & filter(sequence, "null") & "}"
        Log "하트비트를 보냅니다..."
        AddListItem List1, "Myself: " & "{""op"":1,""d"":" & filter(sequence, "null") & "}"
    Else
        timHeartbeat.Enabled = 0
    End If
End Sub

Private Sub tvGuilds_Click()
    On Error Resume Next
    Dim Channel As Dictionary
    Dim msg
    If Not channels.Exists(tvGuilds.SelectedItem.key) Then Exit Sub
    Set Channel = channels(tvGuilds.SelectedItem.key)
    If Channel("type") <> 0 Then Exit Sub
    wbChat.Document.parentWindow.execScript "document.getElementById('message-container').innerHTML = '';"
    If Channel("messages").Count = 0 Then
        Dim Http As New WinHttp.WinHttpRequest
        EnableTLS Http
        
        Http.Open "GET", "https://discord.com/api/v10/channels/" & Channel("id") & "/messages", True
        Http.SetRequestHeader "Content-Type", "application/json"
        Http.SetRequestHeader "Authorization", Token
        Http.SetRequestHeader "User-Agent", "My XML App V1.0"
        Http.Send
        Http.WaitForResponse 60
        
        If Http.Status >= 400 Then
            MsgBox "메시지를 불러올 수 없습니다.", 16, "오류"
            Exit Sub
        End If
        
        Dim p As Object
        Set p = JSON.parse("{""messages"":" & CStr(Http.ResponseText) & "}")
        For Each msg In p("messages")
            Channel("messages").Add CStr(msg("id")), msg
        Next msg
    End If
    
    Log "메시지 " & Channel("messages").Count & "개가 있습니다"
    
    For Each msg In Channel("messages").Items
        wbChat.Document.parentWindow.execScript "createMessage(" & JSON.toString(msg("author")) & ", """ & EscapeHTML(Escape(msg("content"))) & """);"
        If Not CStr(users.Exists(msg("author")("id"))) And Not IsNull(msg("webhook_id")) Then
            users.Add CStr(msg("author")("id")), msg("author")
        End If
    Next msg
End Sub

'==================================================================
'websocket events
'==================================================================

'server has closed
Private Sub ws_onClose(ByVal eCode As WebsocketStatus, ByVal reason As String)
    Select Case CInt(CStr(eCode))
        Case 4000 To 4003, 4005 To 4009
            ResumeDiscord
    End Select
    Log "서버와의 접속이 끊겼습니다 - " & eCode & "/" & reason
End Sub

'websocket is connected
Private Sub ws_OnConnect(ByVal RemoteHost As String, ByVal RemoteIP As String, ByVal RemotePort As String)

    AddListItem List2, "Connected to server " & RemoteHost & " (" & RemoteIP & ") on port " & RemotePort

    'if in btc demo mode then send subscribe request
    If BtcDemo Then
        'subscribe to bitcoin price ticker
        ws.Send "{ ""type"": ""subscribe"",""channels"": [{ ""name"": ""ticker"", ""product_ids"": [""BTC-USD""] }] }"
    End If

End Sub

'websocket server error
Private Sub ws_onError(ByVal eCode As WebsocketStatus, ByVal reason As String)
    'note: you can use ws.getstatuscodetext(ecode) for a status text which is usually similar or same as reason
    AddListItem List2, "Error! " & CStr(eCode) & ": " & reason
End Sub



Private Sub ws_OnMessage(ByVal msg As Variant, ByVal OpCode As WebsocketOpCode)
    'if coinbase demo not active proceed as normal
    Dim evt As Dictionary
    Select Case OpCode
        Case opText     'normal text string
            AddListItem List1, "Server: " & msg
            
            Set evt = JSON.parse(CStr(msg))
            Select Case CInt(evt("op"))
                Case 10 'hello
                    Log "게이트웨이 연결 완료!"
                    HeartbeatInterval = evt("d")("heartbeat_interval")
                    timHeartbeat.Interval = HeartbeatInterval
                    timHeartbeat.Enabled = -1
                    If Not IsNull(evt("s")) Then
                        sequence = evt("s")
                    End If
                    
                    If Resuming Then
                        Log "세션을 다시 시작합니다"
                        ws.Send "{""op"":6,""d"":{""token"":""" & Token & """,""session_id"":""" & sessionID & """,""seq"":" & sequence & "}}"
                        AddListItem List1, "{""op"":6,""d"":{""token"":""" & Token & """,""session_id"":""" & sessionID & """,""seq"":" & sequence & "}}"
                        Resuming = False
                    Else
                        frmConnectionProgress.lblStatus.Caption = "로그인하는 중..."
                        Log "계정 인증 중..."
                        ws.Send "{""op"":2,""d"":{""token"":""" & Token & """,""properties"":{""os"":""Android"",""browser"":""Discord Android"",""device"":""""},""compress"":false,""large_threshold"":250,""presence"":{""activities"":[],""status"":""online"",""afk"":false},""intents"":65535}}"
                        AddListItem List1, "{""op"":2,""d"":{""token"":""" & Token & """,""properties"":{""os"":""Android"",""browser"":""Discord Android"",""device"":""""},""compress"":false,""large_threshold"":250,""presence"":{""activities"":[],""status"":""online"",""afk"":false},""intents"":65535}}"
                    End If
                Case 7 'reconnect
                    Log "게이트웨이에 다시 연결합니다..."
                    ResumeDiscord
                Case 0
                    sequence = evt("s")
                    Select Case evt("t")
                        Case "READY"
                            frmConnectionProgress.lblStatus.Caption = "연결되었습니다."
                            Pause 1
                            Unload frmConnectionProgress
                            Log "계정 인증 완료!"
                            sessionID = evt("d")("session_id")
                            ResumeURL = evt("d")("resume_gateway_url")
                            Me.Caption = evt("d")("user")("username") & " - " & App.Title
                            
                            '서버 목록 불러오기
                            For Guild = 1 To evt("d")("guilds").Count
                                tvGuilds.Nodes.Add , , CStr(evt("d")("guilds")(Guild)("id")), "(불러오는 중...)"
                            Next Guild
                        Case "GUILD_CREATE"
                            Log CStr(evt("d")("name")) & " 서버에 대한 정보를 받았습니다."
                            tvGuilds.Nodes(CStr(evt("d")("id"))).Text = CStr(evt("d")("name"))
                            guilds.Add CStr(evt("d")("id")), evt("d")
                            
                            Dim guildChannels As Object
                            Set guildChannels = evt("d")("channels")
                            For i% = 1 To guildChannels.Count
                                If Len(CStr(guildChannels(i)("parent_id"))) = 0 Then
                                    If guildChannels(i)("type") = 4 Then
                                        tvGuilds.Nodes.Add CStr(evt("d")("id")), tvwChild, CStr(guildChannels(i)("id")), "[ " & guildChannels(i)("name") & " ]"
                                    Else
                                        tvGuilds.Nodes.Add CStr(evt("d")("id")), tvwChild, CStr(guildChannels(i)("id")), guildChannels(i)("name")
                                    End If
                                End If
                                channels.Add CStr(guildChannels(i)("id")), guildChannels(i)
                                channels(CStr(guildChannels(i)("id"))).Add "messages", New Dictionary
                            Next i
                            For i% = 1 To guildChannels.Count
                                If Len(CStr(guildChannels(i)("parent_id"))) > 0 Then
                                    tvGuilds.Nodes.Add CStr(guildChannels(i)("parent_id")), tvwChild, CStr(guildChannels(i)("id")), guildChannels(i)("name")
                                End If
                            Next i
                        Case "MESSAGE_CREATE"
                            Log "새로운 메시지가 왔습니다."
                            Dim message As Dictionary
                            Set message = evt("d")
                            channels(CStr(message("channel_id")))("messages").Add CStr(message("id")), message
                            wbChat.Document.parentWindow.execScript "createMessage(" & JSON.toString(message("author")) & ", """ & EscapeHTML(Escape(message("content"))) & """);"
                            If Not users.Exists(CStr(message("author")("id"))) And Not IsNull(message("webhook_id")) Then
                                users.Add CStr(message("author")("id")), message("author")
                            End If
                    End Select
                Case 11
                    Log "하트비트가 정상적으로 전송되었습니다"
            End Select
        Case opBinary  'handle binary data

            If VarType(msg) = vbArray + vbByte Then
                AddListItem List1, "Server: Binary data recieved"
                'do something with binary data
            Else
                'error, data type cannot be handled
                AddListItem List1, "Error, invalid data type received."
            End If

        Case 3 To 7, 11 To 15
            AddListItem List1, "Extension opCode and data received. ( " & CStr(OpCode) & " )"
            'Msg is the raw untouched data (byte array) from the web server, not uncompressed or unencoded

    End Select
End Sub

'server has answered ping
Private Sub ws_OnPong(ByVal IncludedMsg As String)
    AddListItem List1, "Server: Pong " & IIf(Len(IncludedMsg) > 0, " - " & IncludedMsg, "")
End Sub


'NEW EVENT: the remote server is redirecting our connection
Private Sub ws_OnReConnect(ByVal newURI As String)
    Text1 = newURI
    AddListItem List2, "The server is redirecting the connection to " & newURI
End Sub


'====================================================================
'helper functions
'====================================================================

'display the BTC price for coinbase demo
Sub ShowBtcPrice(ByVal strData As String)
    '    "type":"ticker",
    '    "trade_id":20153558,
    '    "sequence":3262786978,
    '    "time":"2017-09-02T17:05:49.250000Z",
    '    "product_id":"BTC-USD",
    '    "price":"4388.01000000",
    '    "side":"buy", // Taker side
    '    "last_size":"0.03000000",
    '    "best_bid":"4388",
    '    "best_ask":"4388.01"

    Dim Pos As Long
    Dim Last As Long
    Dim Price As String

    'extract the price
    Pos = InStr(strData, """price"":")
    If Pos Then
        Pos = Pos + 9
        Last = InStr(Pos + 1, strData, """")    'look for end quote
        If Last Then
            Price = Mid$(strData, Pos, Last - Pos)

            If InStr(strData, """side"":""buy""") Then
                List3.List(1) = "BTC BUY: " & Price
            Else
                List3.List(0) = "BTC SELL: " & Price
            End If
        End If
    End If


End Sub


'parse server address from server URL
Function ParseServerAddress(ByVal sUri As String) As String

    Dim lPos As Long
    Dim lEnd As Long

    lPos = InStr(sUri, "//")
    If lPos Then
        lPos = lPos + 2
        lEnd = InStr(lPos, sUri, "/")
        If lEnd = 0 Then
            lEnd = Len(sUri)
        End If
        ParseServerAddress = Mid$(sUri, lPos, lEnd - lPos)
        'check for port and remove
        lPos = InStr(ParseServerAddress, ":")
        If lPos Then
            ParseServerAddress = Left$(ParseServerAddress, lPos - 1)
        End If
    End If

End Function




