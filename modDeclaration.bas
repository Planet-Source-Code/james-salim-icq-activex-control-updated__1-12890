Attribute VB_Name = "modDeclaration"
Public Declare Sub Encrypt Lib "UDPCore.dll" (ByVal strPacket As String, ByVal lenPacket As Long)
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
Public Declare Sub CopyVariable Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Public Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Public Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long

Public Type ClientHeader
  uin As Long
  Password As String
  SessionID As Long
  Command As Integer
  SeqNum1 As Integer
  SeqNum2 As Integer
  Parameter As String
  SrvSeq As Integer
End Type

Public Type ServerHeader
  Version As Integer
  SessionID As Long
  Command As Integer
  SeqNum1 As Integer
  SeqNum2 As Integer
  uin As Long
  Parameter As String
End Type

Public Type ClientSocket
  TCPInternalIP As String
  TCPExternalIP As String
  TCPListenPort As Integer
  UDPRemoteHost As String
  UDPRemotePort As Integer
  UDPLocalPort As Integer
  ConnectionMethod As enumUseTCP
  ConnectionState As enumConnectionSate
  OnlineStatus As enumOnlineState
End Type

Public Type typContactSock
  lngUIN As Long
  tcp_TCPExternalIP As String
  tcp_ExternalPort As Integer
  tcp_TCPInternalIP As String
  tcp_Version As Long
  bTcpCapable As Boolean
End Type

Public Type typMessageHeader
  lngUIN As Long
  msg_Date As Date
  msg_Time As String
  msg_Type As enumMessageType
  msg_Text As String
  url_Address As String
  url_Description As String
  auth_NickName As String
  auth_FirstName As String
  auth_LastName As String
  auth_Email As String
  auth_Reason As String
  cont_nick As Variant
  cont_uin As Variant
End Type

'Enumeration
Public Enum UDP_CLIENT_COMMAND
  UDP_CMD_ACK = &HA
  UDP_CMD_ONLINE_MSG = &H10E
  UDP_CMD_LOGIN = &H3E8
  UDP_CMD_CONT_LIST = &H406
  UDP_CMD_SEARCH_UIN = &H41A
  UDP_CMD_SEARCH_USER = &H424
  UDP_CMD_KEEP_ALIVE = &H42E
  UDP_CMD_KEEP_ALIVE2 = &H51E
  UDP_CMD_SEND_TEXT_CODE = &H438
  UDP_CMD_LOGIN_1 = &H44C
  UDP_CMD_INFO_REQ = &H460
  UDP_CMD_EXT_INFO_REQ = &H46A
  UDP_CMD_CHANGE_PW = &H49C
  UDP_CMD_STATUS_CHANGE = &H4D8
  UDP_CMD_LOGIN_2 = &H528
  UDP_CMD_UPDATE_INFO = &H50A
  UDP_CMD_UPDATE_AUTH = &H514
  UDP_CMD_UPDATE_EXT_INFO = &H4B0
  UDP_CMD_ADD_TO_LIST = &H53C
  UDP_CMD_REQ_ADD_LIST = &H456
  UDP_CMD_QUERY_SERVERS = &H4BA
  UDP_CMD_QUERY_ADDONS = &H4C4
  UDP_CMD_NEW_USER_1 = &H4EC
  UDP_CMD_NEW_USER_INFO = &H4A6
  UDP_CMD_ACK_MESSAGES = &H442
  UDP_CMD_MSG_TO_NEW_USER = &H456
  UDP_CMD_REG_NEW_USER = &H3FC
  UDP_CMD_VIS_LIST = &H6AE
  UDP_CMD_INVIS_LIST = &H6A4
  UDP_CMD_UPDATE_LIST = &H6B8
  UDP_CMD_META_USER = &H64A
  UDP_CMD_RAND_SEARCH = &H56E
  UDP_CMD_RAND_SET = &H564
  UDP_CMD_REVERSE_TCP_CONN = &H15E
End Enum

Public Enum UDP_SERVER_REPLY
  UDP_SRV_ACK = &HA
  UDP_SRV_LOGIN_REPLY = &H5A
  UDP_SRV_USER_ONLINE = &H6E
  UDP_SRV_USER_OFFLINE = &H78
  UDP_SRV_USER_FOUND = &H8C
  UDP_SRV_OFFLINE_MESSAGE = &HDC
  UDP_SRV_END_OF_SEARCH = &HA0
  UDP_SRV_INFO_REPLY = &H118
  UDP_SRV_EXT_INFO_REPLY = &H122
  UDP_SRV_STATUS_UPDATE = &H1A4
  UDP_SRV_X1 = &H21C
  UDP_SRV_X2 = &HE6
  UDP_SRV_UPDATE = &H1E0
  UDP_SRV_UPDATE_EXT = &HC8
  UDP_SRV_NEW_UIN = &H46
  UDP_SRV_NEW_USER = &HB4
  UDP_SRV_QUERY = &H82
  UDP_SRV_SYSTEM_MESSAGE = &H1C2
  UDP_SRV_ONLINE_MESSAGE = &H104
  UDP_SRV_GO_AWAY = &HF0
  UDP_SRV_TRY_AGAIN = &HFA
  UDP_SRV_FORCE_DISCONNECT = &H28
  UDP_SRV_MULTI_PACKET = &H212
  UDP_SRV_WRONG_PASSWORD = &H64
  UDP_SRV_INVALID_UIN = &H12C
  UDP_SRV_META_USER = &H3DE
  UDP_SRV_RAND_USER = &H24E
  UDP_SRV_AUTH_UPDATE = &H1F4
End Enum

Public Enum META_COMMAND
  META_CMD_REQ_INFO = 1200
  META_CMD_SET_MAIN = 1000
  META_CMD_SET_MORE = 1020
  META_CMD_SET_ABOUT = 1030
  META_CMD_SET_SECURE = 1060
  META_CMD_SET_WORK = 1010
  META_CMD_SET_INTEREST = 4279
  META_CMD_SET_AFFILIATIONS = 19402
  META_CMD_SET_PASSWORD = 1070
  META_CMD_ACK = 1230
  META_CMD_SEARCH_UIN = &H51E
  META_CMD_SEARCH_EMAIL = &H528
  META_CMD_SEARCH_NAME = &H514
  META_SRV_ACK_INFO = 100
  META_SRV_ACK_HOMEPAGE = 120
  META_SRV_ACK_ABOUT = 130
  META_SRV_ACK_SECURE = 160
  META_SRV_ACK_PASS = 170
  META_SRV_USER_INFO = 200
  META_SRV_USER_WORK = 210
  META_SRV_USER_MORE = 220
  META_SRV_USER_ABOUT = 230
  META_SRV_USER_INTERESTS = 240
  META_SRV_USER_AFFILIATIONS = 250
  META_SRV_USER_HPCATEGORY = 270
  
  META_SRV_SEARCH_FOUND = 400
  META_SRV_SEARCH_LAST = 410
  META_SRV_SUCCESS = 10
  META_SRV_FAILURE = 50
End Enum

Public Const icqUdpVersion = 5
Public Const icqTcpVersion = 6

'Amount of second to wait before resending an unsuccessful packet
Public Const QInterval = 10

'Interval between sending of KeepAlive packet, normal is 2min(120 sec)
Public Const AliveInterval = 120

Public udpSockInfo As ClientSocket
Public udpHeader As ClientHeader
Public tmrICQUdp As Integer
Public tmrKeepAlive As Integer
Public Events As IcqUdp
Public TempUserInfo As typContactInfo
