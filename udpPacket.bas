Attribute VB_Name = "udpPacket"
Option Explicit

Function Packet_Create(Header As ClientHeader, Optional Sequence As Integer = -1) As String
  'Convert the ClientHeader type into a string (packet) format and encrypt it before it is sent to the server
  
  Dim strPacket As String
  Dim Sequence1 As Integer
  Dim Sequence2 As Integer
  Dim OutLen As Long
  
  If Sequence = -1 Then
    With Header
      If .SeqNum1 = &H7FFF Then .SeqNum1 = 0 Else .SeqNum1 = .SeqNum1 + 1
      If .SeqNum2 = &H7FFF Then .SeqNum2 = 0 Else .SeqNum2 = .SeqNum2 + 1
      Sequence1 = .SeqNum1: Sequence2 = .SeqNum2
    End With
  Else
    Sequence1 = Sequence
    Sequence2 = 0
  End If
    
  strPacket = _
    Mki(icqUdpVersion) + _
    Mkl(0) + _
    Mkl(Header.uin) + _
    Mkl(Header.SessionID) + _
    Mki(Header.Command) + _
    Mki(Sequence1) + _
    Mki(Sequence2) + _
    Mkl(0) + _
    Header.Parameter
    
  Encrypt strPacket, Len(strPacket)
  Packet_Create = strPacket
  
  DebugTxt "Send", convClientCmd(Header.Command), Sequence1
End Function

'######################################################################################################
'######################################################################################################

Function Asm_Simple(ByVal ClientCommand As UDP_CLIENT_COMMAND, Optional seq As Integer = -1) As Integer
  Dim Packet As String
  udpHeader.Command = ClientCommand
  udpHeader.Parameter = Mkl(Rnd(Timer) * &H7FFFFFFF)

  If seq = -1 Then
    Packet = Packet_Create(udpHeader)
    Queue_Set Packet, udpHeader.SeqNum1
    Asm_Simple = udpHeader.SeqNum1
  Else
    Packet = Packet_Create(udpHeader, seq)
    If (ClientCommand <> UDP_CMD_ACK) And _
       (ClientCommand <> UDP_CMD_SEND_TEXT_CODE) Then _
    Queue_Set Packet, seq
    Asm_Simple = seq
  End If
  
  Socket_Send Packet
End Function

Function Asm_UIN(ByVal ClientCommand As UDP_CLIENT_COMMAND, ByVal uin As Long, Optional BeforeParameter As String = "", Optional AfterParameter As String = "") As Integer
  Asm_UIN = Asm_String(ClientCommand, BeforeParameter + Mkl(uin) + AfterParameter)
End Function

Function Asm_String(ByVal ClientCommand As UDP_CLIENT_COMMAND, ByVal strText As String) As Integer
  Dim Packet As String
  udpHeader.Command = ClientCommand
  udpHeader.Parameter = strText

  Packet = Packet_Create(udpHeader)
  Socket_Send Packet
  Queue_Set Packet, udpHeader.SeqNum1
  Asm_String = udpHeader.SeqNum1
End Function

Function Asm_UINList(ByVal ClientCommand As UDP_CLIENT_COMMAND, Optional UINList)
  Dim Total As Integer
  Dim Pos As Integer
  Dim i As Integer
  
  udpHeader.Command = ClientCommand
  
  If UBound(UINList) < 0 Then Exit Function
  
  Pos = -1
  Total = 0
  udpHeader.Parameter = ""
  
  Do
    Pos = Pos + 1
    Total = Total + 1
    udpHeader.Parameter = udpHeader.Parameter + Mkl(UINList(Pos))
      
    If Total = 120 Then
      udpHeader.Parameter = Chr$(Total) + udpHeader.Parameter
      Total = 0
      Asm_UINList = Asm_String(ClientCommand, udpHeader.Parameter)
      udpHeader.Parameter = ""
    End If
  Loop Until Pos = UBound(UINList)
  
  If Len(udpHeader.Parameter) > 0 Then
    udpHeader.Parameter = Chr$(Total) + udpHeader.Parameter
    Asm_UINList = Asm_String(ClientCommand, udpHeader.Parameter)
  End If
End Function
