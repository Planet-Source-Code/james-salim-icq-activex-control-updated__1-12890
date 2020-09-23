VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.UserControl IcqUdp 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   CanGetFocus     =   0   'False
   ClientHeight    =   945
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1020
   HasDC           =   0   'False
   InvisibleAtRuntime=   -1  'True
   Picture         =   "icqUDP.ctx":0000
   PropertyPages   =   "icqUDP.ctx":0C44
   ScaleHeight     =   945
   ScaleWidth      =   1020
   ToolboxBitmap   =   "icqUDP.ctx":0C58
   Begin MSWinsockLib.Winsock WinsockF 
      Left            =   525
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
End
Attribute VB_Name = "IcqUdp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit


Public Type typContactInfo
  lngUIN As Long
  '-- Main --
  strNickname As String
  strFirstName As String
  strLastName As String
  strEmail As String
  strEmail2 As String
  strEmail3 As String
  strCity As String
  strState As String
  strPhone As String
  strFax As String
  strStreet As String
  strCellular As String
  lngZip As Long
  intCountryCode As Integer
  byteTimeZone As Byte
  bEmailHide As Boolean
  '-- More --
  intAge As Integer
  byteGender As Byte
  strHomepageURL As String
  byteBirthYear As Byte
  byteBirthMonth As Byte
  byteBirthDay As Byte
  byteLanguage1 As Byte
  byteLanguage2 As Byte
  byteLanguage3 As Byte
  '-- About --
  strAboutInfo As String
  '-- Work --
  strWorkCity As String
  strWorkState As String
  strWorkPhone As String
  strWorkFax As String
  strWorkAddress As String
  lngWorkZip As Long
  intWorkCountry As Integer
  strWorkName As String
  strWorkDepartment As String
  strWorkPosition As String
  intWorkOccupation As Integer
  strWorkWebURL As String
  '-- Interest --
  byteInterestTotal As Byte
  intInterestCategory(3) As Integer
  strInterestName(3) As String
  '-- Past Background --
  byteBackgroundTotal As Byte
  intBackgroundCategory(3) As Integer
  strBackgroundName(3) As String
  byteOrganizationTotal As Byte
  intOrganizationCategory(3) As Integer
  strOrganizationName(3) As String
  '-- HomePage Category --
  byteHPCategoryTotal As Byte
  intHPCategoryCategory(3) As Integer
  strHPCategoryName(3) As String
  '-- Security --
  bAuthorize As Boolean
  bWebPresence As Boolean
  bPublishIP As Boolean
End Type


Private TimerID As Long
Private MetaUserUIN As Long
Private MetaUserAck(0 To 65535) As Long

Public Event Connected()
Public Event Disconnected()
Public Event Registered()
Public Event ContactOnline(uin As Long, OnlineState As enumOnlineState, IntIP As String, ExtIP As String, ExtPort As Integer, bTcpCapable As Boolean, TcpVersion As Long)
Public Event ContactStatusChange(uin As Long, State As enumOnlineState)
Public Event ContactOffline(uin As Long)
Public Event InfoReply(InfoType As enumInfoType, Info As typContactInfo)
Public Event SearchReply(uin As Long, Nick As String, First As String, Last As String, Email As String, bAuth As Boolean, SearchResult As enumSearchResult)
Public Event MessageReceived(uin As Long, MsgDate As Date, MsgTime As String, MsgType As enumMessageType, MsgText As String, URLAddress As String, URLDescription As String, authNick As String, authFirst As String, authLast As String, authEmail As String, authReason As String, contNick As Variant, contUin As Variant)
Public Event ErrorFound(Number As enumErrorConstant, Description As String)
Public Event PacketAcknowledge(PacketSeq As Integer)
Public Event DebugOut(DebugTxt As String)

'Default Property Values:
Const m_def_UserUin = 0
Const m_def_UserPassword = ""
Const m_def_LocalIP = ""
Const m_def_LocalRealIP = ""
Const m_def_LocalPort = 0
Const m_def_RemoteHost = "icq.mirabilis.com"
Const m_def_RemotePort = 4000
Const m_def_UseTCP = icqNoTCP
Const m_def_OnlineState = icqOnline
Const m_def_SocketState = icqDisconnected

Private Sub tmrIcq_Timer()
  tmrIcqUdp_Tick
End Sub

Private Sub UserControl_Initialize()
  Dim clsTest As clsIcqUtilities
  Set clsTest = New clsIcqUtilities

  Set Events = Me
  Set Winsock = WinsockF
  TimerID = SetTimer(UserControl.hwnd, UserControl.hwnd, 1000, AddressOf tmrIcqUdp_Tick)
End Sub

Private Sub UserControl_Terminate()
  KillTimer UserControl.hwnd, TimerID
End Sub

Private Sub UserControl_Resize()
  UserControl.Height = 510
  UserControl.Width = 510
End Sub

Friend Sub TriggerEvent(ByVal EventNo As Integer, ParamArray ArgList() As Variant)
  Select Case EventNo
    Case 0: RaiseEvent Connected
    Case 1: RaiseEvent Disconnected
    Case 2: RaiseEvent Registered
    Case 4: RaiseEvent ContactOffline(CLng(ArgList(0)))
    Case 10
      RaiseEvent PacketAcknowledge(CInt(ArgList(0)))
      If MetaUserAck(ArgList(0)) <> 0 Then
        MetaUserUIN = MetaUserAck(ArgList(0))
        MetaUserAck(ArgList(0)) = 0
      End If
  End Select
End Sub

Friend Sub ContactOnline(ByVal uin As Long, OnlineState As enumOnlineState, IntIP As String, ExtIP As String, ExtPort As Integer, _
  bTcpCapable As Boolean, TcpVersion As Long)
  RaiseEvent ContactOnline(uin, OnlineState, IntIP, ExtIP, ExtPort, bTcpCapable, TcpVersion)
End Sub
  
Friend Sub ContactStatusChange(uin As Long, State As enumOnlineState)
  RaiseEvent ContactStatusChange(uin, State)
End Sub

Friend Sub InfoReply(InfoType As enumInfoType, Info As typContactInfo)
  If Info.lngUIN = 0 Then Info.lngUIN = MetaUserUIN
  RaiseEvent InfoReply(InfoType, Info)
End Sub

Friend Sub SearchReply(uin As Long, Nick As String, First As String, _
  Last As String, Email As String, bAuth As Boolean, SearchResult As enumSearchResult)
  RaiseEvent SearchReply(uin, Nick, First, Last, Email, bAuth, SearchResult)
End Sub
  
Friend Sub MessageReceived(uin As Long, MsgDate As Date, MsgTime As String, MsgType As enumMessageType, _
  MsgText As String, URLAddress As String, URLDescription As String, _
  authNick As String, authFirst As String, authLast As String, authEmail As String, authReason As String, _
  contNick As Variant, contUin As Variant)
  
  RaiseEvent MessageReceived(uin, MsgDate, MsgTime, MsgType, MsgText, URLAddress, URLDescription, _
  authNick, authFirst, authLast, authEmail, authReason, contNick, contUin)
End Sub
  
Friend Sub ErrorFound(Number As enumErrorConstant, Description As String)
  RaiseEvent ErrorFound(Number, Description)
End Sub

Friend Sub DebugOut(DebugTxt As String)
  RaiseEvent DebugOut(DebugTxt)
End Sub

'-------------------- END EVENTS ----------------------


Public Sub Connect()
  Send_Login
End Sub

Public Sub Disconnect()
  Send_Logout
End Sub

Public Sub Register(Optional Password As String = vbNullString)
  If Password <> vbNullString Then udpHeader.Password = Password
  Send_RegNewUser
End Sub

Public Sub ChangePassword(Optional Password As String = vbNullString)
  If Password = vbNullString Then Password = udpHeader.Password
  SendMeta_UpdatePassword Password
End Sub

Public Sub ContactAdd(ParamArray UINList() As Variant)
  Dim TempList As Variant
  TempList = JoinArray(UINList)
  If UBound(TempList) = 0 Then
    Send_AddToList TempList(0)
  Else
    Send_ContactList TempList
  End If
End Sub

Public Sub VisibleAdd(ParamArray UINList() As Variant)
  Dim TempList As Variant
  TempList = JoinArray(UINList)
  If UBound(TempList) = 0 Then
    Send_UpdateList TempList(0), True, True
  Else
    Send_VisList TempList
  End If
End Sub

Public Sub InvisibleAdd(ParamArray UINList() As Variant)
  Dim TempList As Variant
  TempList = JoinArray(UINList)
  If UBound(TempList) = 0 Then
    Send_UpdateList TempList(0), False, True
  Else
    Send_InvisList TempList
  End If
End Sub

Public Sub VisibleRemove(ParamArray UINList() As Variant)
  Dim TempList As Variant, i As Integer
  TempList = JoinArray(UINList)
  For i = 0 To UBound(TempList)
    Send_UpdateList TempList(i), True, False
  Next i
End Sub
Public Sub InvisibleRemove(ParamArray UINList() As Variant)
  Dim TempList As Variant, i As Integer
  TempList = JoinArray(UINList)
  For i = 0 To UBound(TempList)
    Send_UpdateList TempList(i), False, False
  Next i
End Sub

Public Sub InfoRequestBasic(ByVal uin As Long)
  Send_InfoReq uin
End Sub
Public Sub InfoRequestMore(ByVal uin As Long)
  Send_ExtInfoReq uin
End Sub
Public Sub InfoRequestAll(ByVal uin As Long)
  MetaUserAck(SendMeta_RequestInfo(uin)) = uin
End Sub

Public Sub InfoUpdate(ByVal InfoUpdateType As enumInfoType, InfoDetail As typContactInfo)
  Select Case InfoUpdateType
    Case icqNewUser:        Send_UpdateNewUserInfo InfoDetail
    Case icqBasic:          Send_UpdateInfo InfoDetail
    Case icqMain:           SendMeta_UpdateMain InfoDetail
    Case icqMore:           SendMeta_UpdateMore InfoDetail
    Case icqWork:           SendMeta_UpdateWork InfoDetail
    Case icqInterest:       SendMeta_UpdateInterest InfoDetail
    Case icqAffiliations:   SendMeta_UpdateAffiliations InfoDetail
    Case icqAbout:          SendMeta_UpdateAbout InfoDetail
    Case icqSecurity:       SendMeta_UpdateSecurity InfoDetail
    Case Else
      SendMeta_UpdateMain InfoDetail
      SendMeta_UpdateMore InfoDetail
      SendMeta_UpdateWork InfoDetail
      SendMeta_UpdateInterest InfoDetail
      SendMeta_UpdateAffiliations InfoDetail
      SendMeta_UpdateAbout InfoDetail
      SendMeta_UpdateSecurity InfoDetail
  End Select
End Sub

Public Sub SearchUin(ByVal uin As Long)
  Send_SearchReqUIN uin
End Sub
Public Sub SearchName(ByVal Nickname As String, ByVal Firstname As String, ByVal Lastname As String)
  Send_SearchReq Nickname, Firstname, Lastname, ""
End Sub
Public Sub SearchEmail(ByVal EmailAddress As String)
  Send_SearchReq "", "", "", EmailAddress
End Sub

Public Function SendText(ByVal uin As Long, ByVal Message As String) As Integer
  Dim TempMsg As typMessageHeader
  TempMsg.lngUIN = uin
  TempMsg.msg_Type = icqMsgText
  TempMsg.msg_Text = Message
  SendText = Send_OnlineMessage(TempMsg)
End Function
Public Function SendURL(ByVal uin As Long, ByVal URLAddress As String, ByVal URLDescription As String) As Integer
  Dim TempMsg As typMessageHeader
  TempMsg.lngUIN = uin
  TempMsg.msg_Type = icqMsgURL
  TempMsg.url_Address = URLAddress
  TempMsg.url_Description = URLDescription
  SendURL = Send_OnlineMessage(TempMsg)
End Function
Public Function SendAuthReq(ByVal uin As Long, ByVal Nickname As String, ByVal Firstname As String, ByVal Lastname As String, ByVal EmailAddress As String, ByVal Reason As String) As Integer
  Dim TempMsg As typMessageHeader
  With TempMsg
    .lngUIN = uin
    .msg_Type = icqMsgAuthReq
    .auth_NickName = Nickname
    .auth_FirstName = Firstname
    .auth_LastName = Lastname
    .auth_Email = EmailAddress
    .auth_Reason = Reason
  SendAuthReq = Send_OnlineMessage(TempMsg)
  End With
End Function
Public Function SendAuthAccept(ByVal uin As Long) As Integer
  Dim TempMsg As typMessageHeader
  TempMsg.lngUIN = uin
  TempMsg.msg_Type = icqMsgAuthAccept
  SendAuthAccept = Send_OnlineMessage(TempMsg)
End Function
Public Function SendAuthDecline(ByVal uin As Long, ByVal Reason As String) As Integer
  Dim TempMsg As typMessageHeader
  TempMsg.lngUIN = uin
  TempMsg.msg_Type = icqMsgAuthDecline
  TempMsg.auth_Reason = Reason
  SendAuthDecline = Send_OnlineMessage(TempMsg)
End Function
Public Function SendContact(ByVal uin As Long, ByVal UINList As Variant, ByVal NickList As Variant) As Integer
  Dim TempMsg As typMessageHeader
  TempMsg.lngUIN = uin
  TempMsg.msg_Type = icqMsgContact
  TempMsg.cont_uin = JoinArray(UINList)
  TempMsg.cont_nick = JoinArray(NickList)
  SendContact = Send_OnlineMessage(TempMsg)
End Function
Public Function SendUserAdd(ByVal uin As Long) As Integer
  SendUserAdd = Send_AddToList(uin)
End Function

'########################################################################################
' P R O P E R T I E S
'########################################################################################

Public Property Get UserUin() As Long
Attribute UserUin.VB_ProcData.VB_Invoke_Property = "General"
  UserUin = udpHeader.uin
End Property
Public Property Let UserUin(ByVal New_UserUin As Long)
  udpHeader.uin = New_UserUin
  PropertyChanged "UserUin"
End Property

Public Property Get UserPassword() As String
  UserPassword = udpHeader.Password
End Property
Public Property Let UserPassword(ByVal New_UserPassword As String)
  udpHeader.Password = Left$(New_UserPassword, 8)
  PropertyChanged "UserPassword"
End Property

Public Property Get LocalIP() As String
Attribute LocalIP.VB_ProcData.VB_Invoke_Property = "General"
  LocalIP = udpSockInfo.TCPInternalIP
End Property
Public Property Let LocalIP(ByVal New_LocalIP As String)
  udpSockInfo.TCPInternalIP = CvIP(MkIP(New_LocalIP))
  PropertyChanged "LocalIP"
End Property

Public Property Get LocalRealIP() As String
Attribute LocalRealIP.VB_ProcData.VB_Invoke_Property = "General"
  LocalRealIP = udpSockInfo.TCPExternalIP
End Property
Public Property Let LocalRealIP(ByVal New_LocalRealIP As String)
  If Ambient.UserMode = False Then Err.Raise 387
  If Ambient.UserMode Then Err.Raise 382
  udpSockInfo.TCPExternalIP = New_LocalRealIP
  PropertyChanged "LocalRealIP"
End Property

Public Property Get LocalPort() As Integer
Attribute LocalPort.VB_ProcData.VB_Invoke_Property = "General"
  LocalPort = udpSockInfo.UDPLocalPort
End Property
Public Property Let LocalPort(ByVal New_LocalPort As Integer)
  udpSockInfo.UDPLocalPort = New_LocalPort
  PropertyChanged "LocalPort"
End Property

Public Property Get RemoteHost() As String
Attribute RemoteHost.VB_ProcData.VB_Invoke_Property = "General"
  RemoteHost = udpSockInfo.UDPRemoteHost
End Property
Public Property Let RemoteHost(ByVal New_RemoteHost As String)
  udpSockInfo.UDPRemoteHost = New_RemoteHost
  PropertyChanged "RemoteHost"
End Property

Public Property Get RemotePort() As Integer
Attribute RemotePort.VB_ProcData.VB_Invoke_Property = "General"
  RemotePort = udpSockInfo.UDPRemotePort
End Property
Public Property Let RemotePort(ByVal New_RemotePort As Integer)
  udpSockInfo.UDPRemotePort = New_RemotePort
  PropertyChanged "RemotePort"
End Property

Public Property Get UseTCP() As enumUseTCP
  UseTCP = udpSockInfo.ConnectionMethod
End Property
Public Property Let UseTCP(ByVal New_UseTCP As enumUseTCP)
  udpSockInfo.ConnectionMethod = New_UseTCP
  PropertyChanged "UseTCP"
End Property

Public Property Get SocketState() As enumConnectionSate
  SocketState = udpSockInfo.ConnectionState
End Property
Public Property Let SocketState(ByVal New_SocketState As enumConnectionSate)
  If Ambient.UserMode = False Then Err.Raise 387
  If Ambient.UserMode Then Err.Raise 382
  udpSockInfo.ConnectionState = New_SocketState
  PropertyChanged "SocketState"
End Property

Public Property Get OnlineState() As enumOnlineState
  OnlineState = udpSockInfo.OnlineStatus
End Property
Public Property Let OnlineState(ByVal New_OnlineState As enumOnlineState)
  Select Case udpSockInfo.ConnectionState
    Case icqConnected
        Send_ChangeStatus New_OnlineState
        udpSockInfo.OnlineStatus = New_OnlineState
        PropertyChanged "OnlineState"
    Case icqDisconnected
        udpSockInfo.OnlineStatus = New_OnlineState
        PropertyChanged "OnlineState"
  End Select
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
  udpHeader.uin = m_def_UserUin
  udpHeader.Password = m_def_UserPassword
  udpSockInfo.TCPInternalIP = m_def_LocalIP
  udpSockInfo.TCPExternalIP = m_def_LocalRealIP
  udpSockInfo.UDPLocalPort = m_def_LocalPort
  udpSockInfo.UDPRemoteHost = m_def_RemoteHost
  udpSockInfo.UDPRemotePort = m_def_RemotePort
  udpSockInfo.ConnectionMethod = m_def_UseTCP
  udpSockInfo.OnlineStatus = m_def_OnlineState
  udpSockInfo.ConnectionState = m_def_SocketState
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

  udpHeader.uin = PropBag.ReadProperty("UserUin", m_def_UserUin)
  udpHeader.Password = PropBag.ReadProperty("UserPassword", m_def_UserPassword)
  udpSockInfo.TCPInternalIP = PropBag.ReadProperty("LocalIP", m_def_LocalIP)
  udpSockInfo.TCPExternalIP = PropBag.ReadProperty("LocalRealIP", m_def_LocalRealIP)
  udpSockInfo.UDPLocalPort = PropBag.ReadProperty("LocalPort", m_def_LocalPort)
  udpSockInfo.UDPRemoteHost = PropBag.ReadProperty("RemoteHost", m_def_RemoteHost)
  udpSockInfo.UDPRemotePort = PropBag.ReadProperty("RemotePort", m_def_RemotePort)
  udpSockInfo.ConnectionMethod = PropBag.ReadProperty("UseTCP", m_def_UseTCP)
  udpSockInfo.OnlineStatus = PropBag.ReadProperty("OnlineState", m_def_OnlineState)
  udpSockInfo.ConnectionState = PropBag.ReadProperty("SocketState", m_def_SocketState)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

  Call PropBag.WriteProperty("UserUin", udpHeader.uin, m_def_UserUin)
  Call PropBag.WriteProperty("UserPassword", udpHeader.Password, m_def_UserPassword)
  Call PropBag.WriteProperty("LocalIP", udpSockInfo.TCPInternalIP, m_def_LocalIP)
  Call PropBag.WriteProperty("LocalRealIP", udpSockInfo.TCPExternalIP, m_def_LocalRealIP)
  Call PropBag.WriteProperty("LocalPort", udpSockInfo.UDPLocalPort, m_def_LocalPort)
  Call PropBag.WriteProperty("RemoteHost", udpSockInfo.UDPRemoteHost, m_def_RemoteHost)
  Call PropBag.WriteProperty("RemotePort", udpSockInfo.UDPRemotePort, m_def_RemotePort)
  Call PropBag.WriteProperty("UseTCP", udpSockInfo.ConnectionMethod, m_def_UseTCP)
  Call PropBag.WriteProperty("OnlineState", udpSockInfo.OnlineStatus, m_def_OnlineState)
  Call PropBag.WriteProperty("SocketState", udpSockInfo.ConnectionState, m_def_SocketState)
  Call PropBag.WriteProperty("SocketState", udpSockInfo.ConnectionState, m_def_SocketState)
End Sub

Sub ShowAboutBox()
Attribute ShowAboutBox.VB_UserMemId = -552
  frmAbout.Show vbModal
  Unload frmAbout
  Set frmAbout = Nothing
End Sub

Private Sub WinsockF_DataArrival(ByVal bytesTotal As Long)
  Dim Data As String
  WinsockF.GetData Data, vbString
  Recv_DataArrival Data
End Sub
