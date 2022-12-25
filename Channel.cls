VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Channel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Guild As Guild

Public Enum ChannelTypes
    Text = 0
    Voice = 2
    Category = 4
    News = 5
    Stage = 13
End Enum

Public Name As String
Public ID As String
Public Threads As New Dictionary
Public Messages As New Dictionary
Public ChannelType As ChannelTypes
Public ParentID As String
Public Parent As Channel
Public Position As Integer

Sub Initialize(raw As Dictionary)
    Name = raw("name")
    ID = raw("id")
    ChannelType = raw("type")
    ParentID = filter(raw("parent_id"), "")
    Position = raw("position")
End Sub

Sub Send(Content As String)
    Request "POST", "https://discord.com/api/v10/channels/" & ID & "/messages", "{""content"":""" & Escape(Content) & """}"
End Sub








