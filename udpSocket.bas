Attribute VB_Name = "udpSocket"
Public Winsock As Winsock
'Public Winsock As aeSocket
'Public WinsockEvent As aeSocketEvent

Public Sub Socket_Connect()
  On Error GoTo ErrorProc

  'Set Winsock = New aeSocket
  'Set WinsockEvent = New clsAeSocketEvent
  'Winsock.SetEventInterface WinsockEvent
  'With Winsock
  '  .Protocol = sck_UDP
  '  .LocalPort = udpSockInfo.UDPLocalPort
  '  .wsOpen udpSockInfo.UDPRemoteHost, udpSockInfo.UDPRemotePort
  '  udpSockInfo.UDPLocalPort = .LocalPort
  'End With
  
  With Winsock
    .RemoteHost = udpSockInfo.UDPRemoteHost
    .RemotePort = udpSockInfo.UDPRemotePort
    .Bind udpSockInfo.UDPLocalPort
    udpSockInfo.UDPLocalPort = .LocalPort
  End With
  Exit Sub
  
ErrorProc:
  Events_ErrorTxt Err.Number, Err.Description
End Sub

Public Sub Socket_Close()
  On Error GoTo ErrorProc
  'Winsock.wsClose
  'Set WinsockEvent = Nothing
  'Set Winsock = Nothing
  Winsock.Close
  Exit Sub
  
ErrorProc:
  Events_ErrorTxt Err.Number, Err.Description

End Sub

Public Sub Socket_Send(ByVal strPacket As String)
  On Error GoTo ErrorProc
  'Winsock.wsSend strPacket
  Winsock.SendData strPacket
  Exit Sub
  
ErrorProc:
  Events_ErrorTxt Err.Number, Err.Description
End Sub
