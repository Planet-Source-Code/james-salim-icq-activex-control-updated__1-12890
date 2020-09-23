Attribute VB_Name = "udpRecv"
Option Explicit

Sub Recv_DataArrival(ByVal strPacket As String)
  Dim SrvReply As ServerHeader
  Dim SubSrvReply As ServerHeader
  Dim TotalPacket As Integer
  Dim i As Integer
      
  SrvReply = Recv_SplitServerHeader(strPacket)
  
  With SrvReply
    If .Version <> 5 Then
      DebugTxt "Recv", "ERROR - Unknown Packet version" + Str$(.Version)
      If udpHeader.SessionID <> SrvReply.SessionID Then Exit Sub
    End If
    
    Select Case .Command
      Case UDP_SRV_MULTI_PACKET
        'Send_ACK .SeqNum1
        TotalPacket = PacketRead(.Parameter, vbByte)
        
        DebugTxt "Recv", "Multi Packet" + Str$(TotalPacket), .SeqNum1
        
        For i = 1 To TotalPacket
          strPacket = PacketRead(.Parameter, vbString, , False)
          SubSrvReply = Recv_SplitServerHeader(strPacket)
          Recv_MainHandle SubSrvReply
        Next i
        
      Case Else
        Recv_MainHandle SrvReply
    End Select
  End With
End Sub

Sub Recv_MainHandle(SrvReply As ServerHeader)
  Dim lngUIN As Long
  Dim MSGRecv As typMessageHeader
  
  With SrvReply
    If .Command = UDP_SRV_ACK Then
      DebugTxt "Recv", convSrvReply(.Command), .SeqNum1
      Queue_Del .SeqNum1
      Exit Sub
    End If
  
    If .SeqNum1 = udpHeader.SrvSeq Then
      DebugTxt "Recv", convSrvReply(.Command), .SeqNum1
      Send_ACK .SeqNum1
    
      Select Case .Command
        Case UDP_SRV_NEW_UIN
          udpHeader.uin = SrvReply.uin
          Events_NewUIN SrvReply.uin
          Socket_Close
        Case UDP_SRV_LOGIN_REPLY:
          udpSockInfo.TCPExternalIP = CvIP(lCut(.Parameter, 4))
          udpSockInfo.ConnectionState = icqConnected
          Send_Login1
          Events_LoggedIn
        Case UDP_SRV_WRONG_PASSWORD
          Send_Logout
          Events_Error icqErrWrongPassword
        Case UDP_SRV_INVALID_UIN
          lngUIN = PacketRead(.Parameter, vbLong)
          Events_InvalidUIN lngUIN
        Case UDP_SRV_TRY_AGAIN
          Send_Logout
          Events_Error icqErrTryAgain
        Case UDP_SRV_GO_AWAY
          Send_Logout
          Events_Error icqErrGoAway
        Case UDP_SRV_X2:              Send_ACKMessages
        Case UDP_SRV_RAND_USER:       Recv_HandleRandomSearch .Parameter
        Case UDP_SRV_INFO_REPLY:      Recv_HandleInfoReply .Parameter, .Command
        Case UDP_SRV_EXT_INFO_REPLY:  Recv_HandleInfoReply .Parameter, .Command
        Case UDP_SRV_META_USER:       RecvMeta_Handle .Parameter
        Case UDP_SRV_USER_ONLINE:     Recv_HandleStatusChange .Parameter, .Command
        Case UDP_SRV_USER_OFFLINE:    Recv_HandleStatusChange .Parameter, .Command
        Case UDP_SRV_STATUS_UPDATE:   Recv_HandleStatusChange .Parameter, .Command
        Case UDP_SRV_USER_FOUND:      Recv_HandleSearch .Parameter, .Command
        Case UDP_SRV_END_OF_SEARCH:   Recv_HandleSearch .Parameter, .Command
        Case UDP_SRV_OFFLINE_MESSAGE: Recv_HandleMessageReply .Parameter, .Command
        Case UDP_SRV_ONLINE_MESSAGE:  Recv_HandleMessageReply .Parameter, .Command
      End Select
      udpHeader.SrvSeq = .SeqNum1 + 1
    ElseIf .SeqNum1 < udpHeader.SrvSeq Then
      Send_ACK .SeqNum1
    End If
  End With
End Sub

Function Recv_SplitServerHeader(ByVal strPacket As String) As ServerHeader
  Dim SrvReply As ServerHeader
  
  With SrvReply
    .Version = PacketRead(strPacket, vbInteger)
    Select Case .Version
      Case 5
        lCut strPacket, 1
        .SessionID = PacketRead(strPacket, vbLong)
        .Command = PacketRead(strPacket, vbInteger)
        .SeqNum1 = PacketRead(strPacket, vbInteger)
        .SeqNum2 = PacketRead(strPacket, vbInteger)
        .uin = PacketRead(strPacket, vbLong)
        lCut strPacket, 4
        .Parameter = strPacket
      Case 4
        .SessionID = PacketRead(strPacket, vbLong)
        .Command = PacketRead(strPacket, vbInteger)
        .SeqNum1 = PacketRead(strPacket, vbInteger)
        .SeqNum2 = PacketRead(strPacket, vbInteger)
        .uin = PacketRead(strPacket, vbLong)
        lCut strPacket, 4
        .Parameter = strPacket
      Case 3
        .Command = PacketRead(strPacket, vbInteger)
        .SeqNum1 = PacketRead(strPacket, vbInteger)
        .SeqNum2 = PacketRead(strPacket, vbInteger)
        .uin = PacketRead(strPacket, vbLong)
        lCut strPacket, 4
        .Parameter = strPacket
      Case 2
        .Command = PacketRead(strPacket, vbInteger)
        .SeqNum1 = PacketRead(strPacket, vbInteger)
        .Parameter = strPacket
    End Select
  End With
  Recv_SplitServerHeader = SrvReply
End Function

'--------------------------------------------'
' Sub Function for handling specific message '
'--------------------------------------------'
Sub Recv_HandleInfoReply(ByVal Parameter As String, ByVal Command As UDP_SERVER_REPLY)
  Dim UserDetail As typContactInfo
  
  With UserDetail
  Select Case Command
    Case UDP_SRV_INFO_REPLY
      .lngUIN = PacketRead(Parameter, vbLong)
      .strNickname = PacketRead(Parameter, vbString)
      .strFirstName = PacketRead(Parameter, vbString)
      .strLastName = PacketRead(Parameter, vbString)
      .strEmail = PacketRead(Parameter, vbString)
      .bAuthorize = IIf(PacketRead(Parameter, vbByte), False, True)
      .bWebPresence = IIf(PacketRead(Parameter, vbByte), True, False)
      .bPublishIP = IIf(PacketRead(Parameter, vbByte), False, True)
      Events_InfoReply icqBasic, UserDetail
    Case UDP_SRV_EXT_INFO_REPLY
      .lngUIN = PacketRead(Parameter, vbLong)
      .strCity = PacketRead(Parameter, vbString)
      .intCountryCode = PacketRead(Parameter, vbInteger)
      .byteTimeZone = PacketRead(Parameter, vbByte)
      .strState = PacketRead(Parameter, vbString)
      .intAge = PacketRead(Parameter, vbInteger)
      .byteGender = PacketRead(Parameter, vbByte)
      .strPhone = PacketRead(Parameter, vbString)
      .strHomepageURL = PacketRead(Parameter, vbString)
      .strAboutInfo = PacketRead(Parameter, vbString)
      Events_InfoReply icqMore, UserDetail
  End Select
  End With
End Sub

Sub Recv_HandleSearch(ByVal Parameter As String, ByVal Command As UDP_SERVER_REPLY)
  Select Case Command
    Case UDP_SRV_USER_FOUND
      Dim Output As typContactInfo
      Output.lngUIN = PacketRead(Parameter, vbLong)
      Output.strNickname = PacketRead(Parameter, vbString)
      Output.strFirstName = PacketRead(Parameter, vbString)
      Output.strLastName = PacketRead(Parameter, vbString)
      Output.strEmail = PacketRead(Parameter, vbString)
      Output.bAuthorize = IIf(PacketRead(Parameter, vbByte), False, True)
      Events_SearchFound Output
    Case UDP_SRV_END_OF_SEARCH
      Events_SearchDone IIf(PacketRead(Parameter, vbByte), True, False)
  End Select
End Sub

Sub Recv_HandleMessageReply(ByVal Parameter As String, ByVal Command As UDP_SERVER_REPLY)
  Dim MSGRecv As typMessageHeader
  Dim TempSplit As Variant
  
  With MSGRecv
    .lngUIN = PacketRead(Parameter, vbLong)
    Select Case Command
      Case UDP_SRV_OFFLINE_MESSAGE
        .msg_Date = Str$(PacketRead(Parameter, vbInteger)) + _
                    Str$(PacketRead(Parameter, vbByte)) + _
                    Str$(PacketRead(Parameter, vbByte))
        .msg_Time = Format(PacketRead(Parameter, vbByte), "00") + ":" + _
                    Format(PacketRead(Parameter, vbByte), "00")
        .msg_Type = PacketRead(Parameter, vbInteger)
        .msg_Text = PacketRead(Parameter, vbString)
      Case UDP_SRV_ONLINE_MESSAGE
        .msg_Date = Format(Date$, "dd-mm-yyyy")
        .msg_Time = Format(Time$, "hh:mm")
        .msg_Type = PacketRead(Parameter, vbInteger)
        .msg_Text = PacketRead(Parameter, vbString)
    End Select
    
    If .msg_Type = icqMsgText Then
      Events_RecvMessage MSGRecv
      Exit Sub
    End If
    
    TempSplit = Split(.msg_Text, Chr$(&HFE), , vbBinaryCompare)
    .msg_Text = ""
    
    If UBound(TempSplit) < 5 Then
      ReDim Preserve TempSplit(5)
    End If
    
    Select Case .msg_Type
      Case icqMsgURL
        .url_Description = TempSplit(0)
        .url_Address = TempSplit(1)
      Case icqMsgAuthReq
        .auth_NickName = TempSplit(0)
        .auth_FirstName = TempSplit(1)
        .auth_LastName = TempSplit(2)
        .auth_Email = TempSplit(3)
        .auth_Reason = TempSplit(5)
      Case icqMsgAuthDecline
        .auth_Reason = TempSplit(0)
      Case icqMsgAdded
        .auth_NickName = TempSplit(0)
        .auth_FirstName = TempSplit(1)
        .auth_LastName = TempSplit(2)
        .auth_Email = TempSplit(3)
      Case icqMsgWebpager
        .auth_NickName = TempSplit(0)
        .auth_FirstName = TempSplit(1)
        .auth_LastName = TempSplit(2)
        .auth_Email = TempSplit(3)
        .msg_Text = TempSplit(5)
      Case icqMsgExpress
        .auth_NickName = TempSplit(0)
        .auth_FirstName = TempSplit(1)
        .auth_LastName = TempSplit(2)
        .auth_Email = TempSplit(3)
        .msg_Text = TempSplit(5)
      Case icqMsgContact
        Dim TotalUIN, i As Integer
        TotalUIN = Val(TempSplit(0))
        If UBound(TempSplit) < TotalUIN * 2 Then ReDim Preserve TempSplit(TotalUIN * 2)
        
        ReDim lstNick(TotalUIN - 1) As String
        ReDim lstUin(TotalUIN - 1) As Long
        For i = 1 To TotalUIN
          lstUin(i - 1) = TempSplit(i * 2)
          lstNick(i - 1) = TempSplit(i * 2 - 1)
        Next i
        
        .cont_nick = lstNick
        .cont_uin = lstNick
        Events_RecvMessage MSGRecv
        Exit Sub
    End Select
  End With
  Events_RecvMessage MSGRecv
End Sub

Sub Recv_HandleStatusChange(ByVal Parameter As String, ByVal Command As UDP_SERVER_REPLY)
  Dim ContactDetail As typContactSock
  Dim OnlineState As enumOnlineState
  
  With ContactDetail
    .lngUIN = PacketRead(Parameter, vbLong)
    Select Case Command
      Case UDP_SRV_USER_ONLINE
        .tcp_TCPExternalIP = CvIP(lCut(Parameter, 4))
        .tcp_ExternalPort = PacketRead(Parameter, vbInteger)
        lCut Parameter, 2
        .tcp_TCPInternalIP = CvIP(lCut(Parameter, 4))
        .bTcpCapable = IIf(PacketRead(Parameter, vbByte) = &H4, True, False)
        OnlineState = PacketRead(Parameter, vbInteger)
        lCut Parameter, vbInteger
        .tcp_Version = PacketRead(Parameter, vbLong)
        Events_ContactOnline ContactDetail, OnlineState
      Case UDP_SRV_USER_OFFLINE
        Events_ContactOffline .lngUIN
      Case UDP_SRV_STATUS_UPDATE
        OnlineState = PacketRead(Parameter, vbInteger)
        Events_ContactStatusChange .lngUIN, OnlineState
    End Select
  End With
End Sub

Sub Recv_HandleRandomSearch(ByVal Parameter As String)
  Dim Output As typContactSock
  Output.lngUIN = PacketRead(Parameter, vbLong)
  Output.tcp_TCPExternalIP = CvIP(lCut(Parameter, 4))
  Output.tcp_ExternalPort = PacketRead(Parameter, vbInteger)
  lCut Parameter, 2
  Output.tcp_TCPInternalIP = CvIP(lCut(Parameter, 4))
  Output.bTcpCapable = IIf(PacketRead(Parameter, vbByte) = 4, True, False)
  lCut Parameter, 4
  Output.tcp_Version = PacketRead(Parameter, vbInteger)
  Events_RandomFound Output
End Sub
