Attribute VB_Name = "modTimer"
Option Explicit

Public Sub tmrIcqUdp_Tick()
  'Keep The builtin timer ticking
  tmrICQUdp = tmrICQUdp + 1
  If tmrICQUdp = 32640 Then tmrICQUdp = 0
    
  'Check for message on queue if connected
  If udpSockInfo.ConnectionState <> icqDisconnected Then Queue_CheckTime

  'Send CMD_KEEP_ALIVE every 2 minute, to make sure the server knows we are still connected
  If udpSockInfo.ConnectionState = icqConnected Then
    tmrKeepAlive = tmrKeepAlive + 1
    If tmrKeepAlive > AliveInterval Then
      Send_KeepAlive
      tmrKeepAlive = 0
    End If
  Else
    tmrKeepAlive = 0
  End If
End Sub

