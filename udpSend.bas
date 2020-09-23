Attribute VB_Name = "udpSend"
Option Explicit

'------------------------------'
' Simple Client Server Message '
'------------------------------'
Sub Send_KeepAlive()
  Asm_Simple UDP_CMD_KEEP_ALIVE
End Sub
Sub Send_ACK(ByVal Seq As Integer)
  Asm_Simple UDP_CMD_ACK, Seq
End Sub
Sub Send_ACKMessages()
  Asm_Simple UDP_CMD_ACK_MESSAGES
End Sub
Sub Send_TextCode(ByVal strTextCode As String)
  Asm_String UDP_CMD_SEND_TEXT_CODE, strTextCode + Mki(5)
End Sub

'------------------'
' UDP Send Message '
'------------------'
Function Send_OnlineMessage(MsgHead As typMessageHeader) As Integer
  Dim Output As String
  With MsgHead
    Select Case .msg_Type
      Case icqMsgText:          Output = .msg_Text
      Case icqMsgAuthDecline:   Output = .auth_Reason
      Case icqMsgAuthAccept:    Output = vbNullChar
      Case icqMsgURL:           Output = strJoinFE(.url_Description, .url_Address)
      Case icqMsgAdded:         Output = strJoinFE(.auth_NickName, .auth_FirstName, .auth_LastName, .auth_Email)
      Case icqMsgAuthReq:       Output = strJoinFE(.auth_NickName, .auth_FirstName, .auth_LastName, .auth_Email, "1", .auth_Reason)
      Case icqMsgContact
        Dim MaxContact As Integer, i As Integer
        MaxContact = UBound(.cont_uin)
        If MaxContact = -1 Then Exit Function
        
        ReDim Preserve .cont_nick(MaxContact)
        Output = Trim(Str$(MaxContact + 1)) + Chr$(&HFE)
        For i = 0 To MaxContact
          Output = Output + strJoinFE(Trim(Str$(.cont_uin(i))), .cont_nick(i), "")
        Next i
    End Select
    
    Output = Mkl(.lngUIN) + Mki(.msg_Type) + StrAppend(Output)
    Send_OnlineMessage = Asm_String(UDP_CMD_ONLINE_MSG, Output)
  End With
End Function

'-----------------------------------------------'
' Login / Logout / Status / Registration Packet '
'-----------------------------------------------'
Function Send_Login1() As Integer
  Send_Login1 = Asm_Simple(UDP_CMD_LOGIN_1)
End Function

Function Send_Login() As Integer
  Dim ParamTime As Long
  Randomize Timer
  
  ParamTime = DateDiff("d", "1-1-1971", Now()) * 24 * 60 * 60
  ParamTime = ParamTime + Timer
  
  udpHeader.SessionID = CLng(Rnd(Timer) * &H3FFFFFFF)
  udpHeader.SeqNum1 = CInt(Rnd(Timer) * &H7FFF)
  udpHeader.SeqNum2 = 0
  udpHeader.Parameter = _
      Mkl(ParamTime) + _
      Mkl(udpSockInfo.TCPListenPort) + _
      StrAppend(udpHeader.Password) + _
      Mkl(&H98) + _
      MkIP(udpSockInfo.TCPInternalIP) + _
      Chr$(udpSockInfo.ConnectionMethod) + _
      Mkl(udpSockInfo.OnlineStatus) + _
      Mkl(&H3) + _
      Mkl(0) + _
      Mkl(&H980010)

  Socket_Connect    ' Connect to ICQ Server
  udpSockInfo.ConnectionState = icqLogin
  Send_Login = Asm_String(UDP_CMD_LOGIN, udpHeader.Parameter)
End Function
Sub Send_Logout()
  ' Disconnect from ICQ Server
  Send_TextCode "B_USER_DISCONNECTED"
  Queue_Reset
  Socket_Close
  Events_Disconnect
  udpHeader.SrvSeq = 0
  udpSockInfo.ConnectionState = icqDisconnected
End Sub
Function Send_RegNewUser() As Integer
  Dim Parameter As String
  udpHeader.uin = 0
  udpHeader.SessionID = CLng(Rnd(Timer) * &H3FFFFFFF)
  udpHeader.SeqNum1 = udpHeader.SeqNum1 + 1
  udpHeader.SeqNum2 = 0
  
  Socket_Connect
  udpSockInfo.ConnectionState = icqRegistering
  
  Parameter = StrAppend(udpHeader.Password) + Mkl(&HA0) + Mkl(&H2461) + Mkl(&HA00000) + Mkl(0)
  Send_RegNewUser = Asm_String(UDP_CMD_REG_NEW_USER, Parameter)
End Function
Function Send_ChangeStatus(OnlineStatus As enumOnlineState) As Integer
  Send_ChangeStatus = Asm_String(UDP_CMD_STATUS_CHANGE, Mkl(OnlineStatus))
End Function

'-------------------------------------------'
' Contact / Visible / Invisible List Packet '
'-------------------------------------------'
Function Send_AddToList(ByVal uin As Long) As Integer
  Send_AddToList = Asm_UIN(UDP_CMD_ADD_TO_LIST, uin)
End Function
Function Send_ContactList(UINList As Variant) As Integer
  Send_ContactList = Asm_UINList(UDP_CMD_CONT_LIST, UINList)
End Function
Function Send_InvisList(UINList As Variant) As Integer
  Send_InvisList = Asm_UINList(UDP_CMD_INVIS_LIST, UINList)
End Function
Function Send_VisList(UINList As Variant) As Integer
  Send_VisList = Asm_UINList(UDP_CMD_VIS_LIST, UINList)
End Function
Function Send_UpdateList(ByVal uin As Long, ByVal bVisibleList As Boolean, ByVal bAdd As Boolean) As Integer
  Send_UpdateList = Asm_String(UDP_CMD_UPDATE_LIST, Mkl(uin) + Mki(IIf(bVisibleList, 2, 1)) + Mki(IIf(bAdd, 1, 0)))
End Function

'--------------------'
' Search User Packet '
'--------------------'
Function Send_SearchReqUIN(ByVal uin As Long) As Integer
  Send_SearchReqUIN = Asm_UIN(UDP_CMD_SEARCH_UIN, uin)
End Function
Function Send_SearchReq(ByVal strNick As String, ByVal strFirst As String, ByVal strLast As String, ByVal strEmail As String) As Integer
  Send_SearchReq = Asm_String(UDP_CMD_SEARCH_USER, StrAppend(strNick, strFirst, strLast, strEmail))
End Function
Function Send_RandomSearch(ByVal RandomGroup As enumRandomGroup)
  Send_RandomSearch = Asm_String(UDP_CMD_RAND_SEARCH, Mki(RandomGroup))
End Function

'--------------------'
' User Detail Packet '
'--------------------'
'** Request **
Function Send_InfoReq(ByVal uin As Long) As Integer
  Send_InfoReq = Asm_UIN(UDP_CMD_INFO_REQ, uin)
End Function
Function Send_ExtInfoReq(ByVal uin As Long) As Integer
  Send_ExtInfoReq = Asm_UIN(UDP_CMD_EXT_INFO_REQ, uin)
End Function

'** Update **
Function Send_RandomSet(ByVal RandomGroup As enumRandomGroup)
  Send_RandomSet = Asm_String(UDP_CMD_RAND_SET, Mkl(RandomGroup))
End Function
Function Send_UpdateInfo(UserDetail As typContactInfo) As Integer
  With UserDetail
    Send_UpdateInfo = Asm_String(UDP_CMD_UPDATE_INFO, StrAppend(.strNickname, .strFirstName, .strLastName, .strEmail))
  End With
End Function
Function Send_UpdateAuthInfo(bAuthorize As Boolean) As Integer
  Send_UpdateAuthInfo = Asm_String(UDP_CMD_UPDATE_AUTH, Mkl(bAuthorize))
End Function
Function Send_UpdateNewUserInfo(UserDetail As typContactInfo) As Integer
  With UserDetail
    Send_UpdateNewUserInfo = Asm_String(UDP_CMD_NEW_USER_INFO, StrAppend(.strNickname, .strFirstName, .strLastName, .strEmail) + String$(3, 1))
  End With
End Function

