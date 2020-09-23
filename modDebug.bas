Attribute VB_Name = "modDebug"
Option Explicit

Public Sub DebugTxt(strCategory As String, strText As String, Optional Seq As Integer)
    Events_Debug strCategory + " (" + Format(Seq, "00000") + ") - " + strText
End Sub

Public Function convSrvReply(ByVal Command As Integer) As String
  Dim Output As String
  Select Case Command
    Case UDP_SRV_ACK: Output = "SRV_ACK"
    Case UDP_SRV_LOGIN_REPLY: Output = "SRV_LOGIN_REPLY"
    Case UDP_SRV_USER_ONLINE: Output = "SRV_USER_ONLINE"
    Case UDP_SRV_USER_OFFLINE: Output = "SRV_USER_OFFLINE"
    Case UDP_SRV_USER_FOUND: Output = "SRV_USER_FOUND"
    Case UDP_SRV_OFFLINE_MESSAGE: Output = "SRV_OFFLINE_MESSAGE"
    Case UDP_SRV_END_OF_SEARCH: Output = "SRV_END_OF_SEARCH"
    Case UDP_SRV_INFO_REPLY: Output = "SRV_INFO_REPLY"
    Case UDP_SRV_EXT_INFO_REPLY: Output = "SRV_EXT_INFO_REPLY"
    Case UDP_SRV_STATUS_UPDATE: Output = "SRV_STATUS_UPDATE"
    Case UDP_SRV_X1: Output = "SRV_X1"
    Case UDP_SRV_X2: Output = "SRV_X2"
    Case UDP_SRV_UPDATE: Output = "SRV_UPDATE"
    Case UDP_SRV_UPDATE_EXT: Output = "SRV_UPDATE_EXT"
    Case UDP_SRV_NEW_UIN: Output = "SRV_NEW_UIN"
    Case UDP_SRV_NEW_USER: Output = "SRV_NEW_USER"
    Case UDP_SRV_QUERY: Output = "SRV_QUERY"
    Case UDP_SRV_SYSTEM_MESSAGE: Output = "SRV_SYSTEM_MESSAGE"
    Case UDP_SRV_ONLINE_MESSAGE: Output = "SRV_ONLINE_MESSAGE"
    Case UDP_SRV_GO_AWAY: Output = "SRV_GO_AWAY"
    Case UDP_SRV_TRY_AGAIN: Output = "SRV_TRY_AGAIN"
    Case UDP_SRV_FORCE_DISCONNECT: Output = "SRV_FORCE_DISCONNECT"
    Case UDP_SRV_MULTI_PACKET: Output = "SRV_MULTI_PACKET"
    Case UDP_SRV_WRONG_PASSWORD: Output = "SRV_WRONG_PASSWORD"
    Case UDP_SRV_INVALID_UIN: Output = "SRV_INVALID_UIN"
    Case UDP_SRV_META_USER: Output = "SRV_META_USER"
    Case UDP_SRV_RAND_USER: Output = "SRV_RAND_USER"
    Case UDP_SRV_AUTH_UPDATE: Output = "SRV_AUTH_UPDATE"
    Case Else: Output = "Unknown" + Str$(Command) + " (0x" + Hex$(Command) + ")"
  End Select
  convSrvReply = Output
End Function

Public Function convClientCmd(ByVal Command As Integer) As String
  Dim Output As String
  Select Case Command
    Case UDP_CMD_ACK: Output = "CMD_ACK"
    Case UDP_CMD_ONLINE_MSG: Output = "CMD_ONLINE_MSG"
    Case UDP_CMD_LOGIN: Output = "CMD_LOGIN"
    Case UDP_CMD_CONT_LIST: Output = "CMD_CONT_LIST"
    Case UDP_CMD_SEARCH_UIN: Output = "CMD_SEARCH_UIN"
    Case UDP_CMD_SEARCH_USER: Output = "CMD_SEARCH_USER"
    Case UDP_CMD_KEEP_ALIVE: Output = "CMD_KEEP_ALIVE"
    Case UDP_CMD_KEEP_ALIVE2: Output = "CMD_KEEP_ALIVE2"
    Case UDP_CMD_SEND_TEXT_CODE: Output = "CMD_SEND_TEXT_CODE"
    Case UDP_CMD_LOGIN_1: Output = "CMD_LOGIN_1"
    Case UDP_CMD_INFO_REQ: Output = "CMD_INFO_REQ"
    Case UDP_CMD_EXT_INFO_REQ: Output = "CMD_EXT_INFO_REQ"
    Case UDP_CMD_CHANGE_PW: Output = "CMD_CHANGE_PW"
    Case UDP_CMD_STATUS_CHANGE: Output = "CMD_STATUS_CHANGE"
    Case UDP_CMD_LOGIN_2: Output = "CMD_LOGIN_2"
    Case UDP_CMD_UPDATE_INFO: Output = "CMD_UPDATE_INFO"
    Case UDP_CMD_UPDATE_AUTH: Output = "CMD_UPDATE_AUTH"
    Case UDP_CMD_UPDATE_EXT_INFO: Output = "CMD_UPDATE_EXT_INFO"
    Case UDP_CMD_ADD_TO_LIST: Output = "CMD_ADD_TO_LIST"
    Case UDP_CMD_REQ_ADD_LIST: Output = "CMD_REQ_ADD_LIST"
    Case UDP_CMD_QUERY_SERVERS: Output = "CMD_QUERY_SERVERS"
    Case UDP_CMD_QUERY_ADDONS: Output = "CMD_QUERY_ADDONS"
    Case UDP_CMD_NEW_USER_1: Output = "CMD_NEW_USER_1"
    Case UDP_CMD_NEW_USER_INFO: Output = "CMD_NEW_USER_INFO"
    Case UDP_CMD_ACK_MESSAGES: Output = "CMD_ACK_MESSAGES"
    Case UDP_CMD_MSG_TO_NEW_USER: Output = "CMD_MSG_TO_NEW_USER"
    Case UDP_CMD_REG_NEW_USER: Output = "CMD_REG_NEW_USER"
    Case UDP_CMD_VIS_LIST: Output = "CMD_VIS_LIST"
    Case UDP_CMD_INVIS_LIST: Output = "CMD_INVIS_LIST"
    Case UDP_CMD_UPDATE_LIST: Output = "CMD_UPDATE_LIST"
    Case UDP_CMD_META_USER: Output = "CMD_META_USER"
    Case UDP_CMD_RAND_SEARCH: Output = "CMD_RAND_SEARCH"
    Case UDP_CMD_RAND_SET: Output = "CMD_RAND_SET"
    Case UDP_CMD_REVERSE_TCP_CONN: Output = "CMD_REVERSE_TCP_CONN"
    Case Else: Output = "Unknown" + Str$(Command) + " (0x" + Hex$(Command) + ")"
  End Select
  convClientCmd = Output
End Function

Public Function convOnlineState(ByVal State) As String
  Dim Output As String
  Select Case State
    Case icqOnline: Output = "Online"
    Case icqInvisible: Output = "Invisible"
    Case icqNa: Output = "Extended Away (N/A)"
    Case icqOccupied: Output = "Occupied"
    Case icqAway: Output = "Away"
    Case icqDND: Output = "Do Not Disturb"
    Case icqChat: Output = "Free for Chat"
    Case Else: Output = "Unknown State" + Str$(State) + " (" + Hex$(State) + ")"
  End Select
  convOnlineState = Output
End Function

Public Function convProxyState(ByVal State) As String
  Dim Output As String
  Select Case State
    Case icqNoTCP: Output = "No TCP"
    Case icqTCPSendOnly: Output = "TCP Send only"
    Case icqTCPSendRecv: Output = "TCP Send & Recv"
    Case Else: Output = "Unknown State" + Str$(State) + " (" + Hex$(State) + ")"
  End Select
  convProxyState = Output
End Function

Public Function convRandomGroup(ByVal RandomGroup As Integer) As String
  Dim Output As String
  Select Case RandomGroup
    Case icqGrpGeneral: Output = "General"
    Case icqGrpRomance: Output = "Romance"
    Case icqGrpGames: Output = "Games"
    Case icqGrpStudents: Output = "Students"
    Case icqGrp20Something: Output = "20 Something"
    Case icqGrp30Something: Output = "30 Something"
    Case icqGrp40Something: Output = "40 Something"
    Case icqGrp50Over: Output = "50Over"
    Case icqGrpManRequestWoman: Output = "Man request woman"
    Case icqGrpWomanRequestMan: Output = "Woman request man"
    Case Else: Output = "Unknown Group" + Str$(RandomGroup) + " (" + Hex$(RandomGroup) + ")"
  End Select
  convRandomGroup = Output
End Function

Public Function convMsgType(ByVal MsgType As Integer) As String
  Dim Output As String
  Select Case MsgType
    Case icqMsgText: Output = "Text"
    Case icqMsgChatReq: Output = "Chat"
    Case icqMsgFile: Output = "File"
    Case icqMsgURL: Output = "URL"
    Case icqMsgAuthReq: Output = "Auth. Request"
    Case icqMsgAuthDecline: Output = "Auth. Decline"
    Case icqMsgAuthAccept: Output = "Auth. Accept"
    Case icqMsgAdded: Output = "Added to List = &HC"
    Case icqMsgWebpager: Output = "WebPager"
    Case icqMsgExpress: Output = "Express"
    Case icqMsgContact: Output = "Contact List"
    Case Else: Output = "Unknown Type" + Str$(MsgType) + " (" + Hex$(MsgType) + ")"
  End Select
  convMsgType = Output
End Function

Public Function convConnectionState(ByVal State As Integer) As String
  Dim Output As String
  Select Case State
    Case icqDisconnected: Output = "Disconnected"
    Case icqRegistering: Output = "Registering New User"
    Case icqLogin: Output = "Logging in..."
    Case icqConnected: Output = "Connected"
    Case Else: Output = "Unknown State" + Str$(State) + " (" + Hex$(State) + ")"
  End Select
  convConnectionState = Output
End Function

