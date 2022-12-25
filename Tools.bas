Attribute VB_Name = "Tools"
Public Token As String
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'https://www.vbforums.com/showthread.php?546633-VB6-Sleep-Function
' Credits: (Milk (Sleep+Pause Sub)). (Wayne Spangler (Pause Sub))
Sub Pause(ByVal Delay As Single)
   Delay = Timer + Delay
   If Delay > 86400 Then 'more than number of seconds in a day
      Delay = Delay - 86400
      Do
          DoEvents ' to process events.
          Sleep 1 ' to not eat cpu
      Loop Until Timer < 1
   End If
   Do
       DoEvents ' to process events.
       Sleep 1 ' to not eat cpu
   Loop While Delay > Timer
End Sub

Function filter(ByVal X, ByVal Y)
    If X = 0 Or IsNull(X) Then
        filter = Y
    Else
        filter = X
    End If
End Function

Sub EnableTLS(ByRef Http)
    On Error Resume Next
    'Http.SetTimeouts ResolveTimeoutMs, ConnectTimeoutMs, SendTimeoutMs, ReceiveTimeoutMs
    Http.Option(9) = 2048
    Http.Option(6) = True
End Sub

Sub SetFont(frm As Form, Optional font As String = "±¼¸²", Optional fbFont As String = "Gulim")
    On Error Resume Next
    For Each ctrl In frm.Controls
        ctrl.FontName = "Tahoma"
        ctrl.FontName = "Segoe UI"
        ctrl.FontName = fbFont
        ctrl.FontName = font
        ctrl.FontSize = 9
    Next ctrl
End Sub

Sub Log(ByVal data)
    frmLogs.lvLogs.AddItem data
End Sub

Function Escape(ByVal str As String)
    Escape = Replace(Replace(str, """", "\"""), "\", "\\")
End Function

Function EscapeHTML(ByVal str As String)
    EscapeHTML = Replace(Replace(Replace(Replace(str, "&", "&amp;"), """", "&quot;"), "<", "&lt;"), ">", "&gt;")
End Function

Function Request(method As String, url As String, Optional body As String) As Object
    Dim Http As New WinHttp.WinHttpRequest
    EnableTLS Http
    
    Http.Open method, url, True
    Http.SetRequestHeader "Content-Type", "application/json"
    Http.SetRequestHeader "Authorization", Token
    Http.SetRequestHeader "User-Agent", "My XML App V1.0"
    Http.Send body
    Http.WaitForResponse 60
    
    Dim p As Object
    Set p = JSON.parse(CStr(Http.ResponseText))
    
    If Http.Status >= 400 Then
        Err.Raise vbObjectError + Http.Status, , p("message")
        Exit Function
    End If
    
    Set Request = p
End Function

Function SetupChannel(raw As Dictionary) As Channel
    Dim ch As New Channel
    ch.Initialize raw
    Set SetupChannel = ch
End Function

