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
    If X = 0 Then
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
