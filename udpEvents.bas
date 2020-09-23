Attribute VB_Name = "udpEvents"
'Raise an event on the main user control, when SVR_ACK is received from server
Sub Events_UDPAck(SequenceNumber As Integer)
  Events.TriggerEvent 10, SequenceNumber
End Sub

'Raise usercontrol event, when connected
Sub Events_LoggedIn()
  Events.TriggerEvent 0
End Sub

'Raise usercontrol event, when disconnected
Sub Events_Disconnect()
  Events.TriggerEvent 1
End Sub

'Raise usercontrol event, when registration successful. New UIN stored on UserUIN properties
Sub Events_NewUIN(uin As Long)
  Events.UserUin = uin
  Events.TriggerEvent 2
End Sub

'Raise usercontrol event, when user is online
Sub Events_ContactOnline(Detail As typContactSock, OnlineState As enumOnlineState)
  With Detail
    Events.ContactOnline .lngUIN, OnlineState, .tcp_TCPInternalIP, .tcp_TCPExternalIP, .tcp_ExternalPort, .bTcpCapable, .tcp_Version
  End With
End Sub

'Raise usercontrol event, when user/contact changes status
Sub Events_ContactStatusChange(uin As Long, State As enumOnlineState)
  Events.ContactStatusChange uin, State
End Sub

'Raise usercontrol event, when user/contact went offline
Sub Events_ContactOffline(ByVal uin As Long)
  Events.TriggerEvent 4, uin
End Sub

'Raise usercontrol event, when message is received
Sub Events_RecvMessage(MSGRecv As typMessageHeader)
  With MSGRecv
    Events.MessageReceived .lngUIN, .msg_Date, .msg_Time, .msg_Type, _
      .msg_Text, .url_Address, .url_Description, _
      .auth_NickName, .auth_FirstName, .auth_LastName, .auth_Email, .auth_Reason, _
      .cont_nick, .cont_uin
  End With
End Sub

'Raise usercontrol event, when user detail is received
Sub Events_InfoReply(InfoType As enumInfoType, UserDetail As typContactInfo)
  Events.InfoReply InfoType, UserDetail
End Sub

'Raise usercontrol event, when a UIN is turn out to be invalid
Sub Events_InvalidUIN(ByVal uin As Long)
  Events.ErrorFound icqErrInvalidUIN, Format(uin, "0000000000") & " is an invalid UIN"
End Sub

'Raise usercontrol event, when search return result
Sub Events_SearchFound(Result As typContactInfo)
  With Result
    Events.SearchReply .lngUIN, .strNickname, .strFirstName, .strLastName, .strEmail, .bAuthorize, icqSearchUserFound
  End With
End Sub

'Raise usercontrol event, when search is done
Sub Events_SearchDone(ByVal bTooMany As Boolean)
  If bTooMany = True Then
    Events.SearchReply 0, "", "", "", "", False, icqSearchTooMany
  Else
    Events.SearchReply 0, "", "", "", "", False, icqSearchDone
  End If
End Sub

Sub Events_RandomFound(Result As typContactSock)
End Sub

Sub Events_Error(Number As enumErrorConstant)
  Dim Desc As String
  Select Case Number
    Case icqErrNotConnected
      Desc = "You are not currently connected/logged in to the ICQ network"
    Case icqErrTryAgain
      Desc = "Unable to connect to the ICQ network. Possible reason is another user is already logged in under the same UIN."
    Case icqErrWrongPassword
      Desc = "Unable to connect to the ICQ network. Wrong password or uin used."
    Case icqErrGoAway
      Desc = "You have been disconnected from the ICQ network. Please try reconnecting later."
  End Select
  Events.ErrorFound Number, Desc
End Sub

Sub Events_ErrorTxt(Number As Long, Description As String)
  Events.ErrorFound Number, Description
End Sub

Sub Events_Debug(DebugTxt As String)
    Events.DebugOut DebugTxt
End Sub
