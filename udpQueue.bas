Attribute VB_Name = "udpQueue"
Option Explicit

Private Type QueueItem
  Packet As String
  Attempt As Integer
End Type

Private Type ExpireData
  TotalQueue As Integer
  QueueList As String
End Type

Private qPacket(0 To 65535) As QueueItem
Private qExpire(0 To 59) As ExpireData

Public Sub Queue_Set(ByVal strPacket As String, Seq As Integer, Optional SendAttempt As Integer = 1)
  Dim ExpireTime As Integer
  qPacket(Seq).Packet = strPacket
  qPacket(Seq).Attempt = SendAttempt
    
  ExpireTime = (tmrICQUdp + QInterval) Mod 60
  With qExpire(ExpireTime)
    .TotalQueue = .TotalQueue + 1
    .QueueList = .QueueList + String$(.TotalQueue - Len(.QueueList), 0)
    Poke StrPtr(.QueueList) + ((.TotalQueue - 1) * 2), Seq, vbInteger
  End With
End Sub

Public Sub Queue_Del(Seq As Integer)
  Events_UDPAck Seq
  If qPacket(Seq).Attempt <> 0 Then
    qPacket(Seq).Packet = vbNullString
    qPacket(Seq).Attempt = 0
  End If
End Sub

Public Sub Queue_Reset()
  Dim i As Integer
  For i = 0 To 59
    qExpire(i).QueueList = vbNullString
    qExpire(i).TotalQueue = 0
  Next i
End Sub

Public Sub Queue_CheckTime()
  Dim ExpireTime As Integer
  Dim Seq As Integer
  Dim i As Integer
  
  With qExpire(tmrICQUdp Mod 60)
  If .TotalQueue > 0 Then
    For i = .TotalQueue To 1 Step -1
      .TotalQueue = .TotalQueue - 1
      Seq = Peek(StrPtr(.QueueList) + ((i - 1) * 2), vbInteger)
      rCut .QueueList, 1
        
      '---------- Packet Resend Module -----------
      With qPacket(Seq)
        If .Attempt > 0 Then
          .Attempt = .Attempt + 1
          Select Case .Attempt - 1
            Case 6:     Send_TextCode "B_MESSAGE_ACK"
            Case 12:    Send_Logout: Exit Sub
            Case Else
              DebugTxt "Queue", "Resending packet" + Str$(.Attempt) + " times.", Seq
              Socket_Send .Packet
          End Select
          Queue_Set .Packet, Seq, .Attempt
        End If
      End With
      '---------- End Packet Resend Module -------
            
    Next i
  End If
  End With
End Sub
