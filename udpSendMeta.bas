Attribute VB_Name = "udpMetaSend"
Option Explicit

'-- Request Info --
Function SendMeta_RequestInfo(ByVal uin As Long) As Integer
  SendMeta_RequestInfo = Asm_UIN(UDP_CMD_META_USER, uin, Mki(META_CMD_REQ_INFO))
End Function

'-- Update Info --
Function SendMeta_UpdatePassword(ByVal strPass As String) As Integer
  SendMeta_UpdatePassword = Asm_String(UDP_CMD_META_USER, Mki(META_CMD_SET_PASSWORD) + StrAppend(strPass))
End Function

Function SendMeta_UpdateMain(UserDetail As typContactInfo) As Integer
  With UserDetail
    SendMeta_UpdateMain = Asm_String(UDP_CMD_META_USER, _
      Mki(META_CMD_SET_MAIN) + _
      StrAppend(.strNickname, .strFirstName, .strLastName, .strEmail, .strEmail2, .strEmail3, _
      .strCity, .strState, .strPhone, .strFax, .strStreet, .strCellular) + _
      Mkl(.lngZip) + Mki(.intCountryCode) + Mki(.byteTimeZone) + _
      Chr$(IIf(.bEmailHide, 1, 0)))
  End With
End Function

Function SendMeta_UpdateMore(UserDetail As typContactInfo) As Integer
  With UserDetail
    SendMeta_UpdateMore = Asm_String(UDP_CMD_META_USER, _
      Mki(META_CMD_SET_MORE) + Chr$(.intAge) + Mki(&H200) + StrAppend(.strHomepageURL) + _
      Chr$(.byteBirthYear) + Chr$(.byteBirthMonth) + Chr$(.byteBirthDay) + String$(3, &HFF))
  End With
End Function

Function SendMeta_UpdateWork(UserDetail As typContactInfo) As Integer
  With UserDetail
    SendMeta_UpdateWork = Asm_String(UDP_CMD_META_USER, _
      Mki(META_CMD_SET_WORK) + _
      StrAppend(.strWorkCity, .strWorkState, .strWorkPhone, .strWorkFax, .strWorkAddress) + _
      Mkl(.lngWorkZip) + Mki(.intWorkCountry) + _
      StrAppend(.strWorkName, .strWorkDepartment, .strWorkPosition) + _
      Mki(.intWorkOccupation) + StrAppend(.strWorkWebURL))
  End With
End Function

Function SendMeta_UpdateInterest(UserDetail As typContactInfo) As Integer
  Dim Parameter As String
  Dim i As Integer
  With UserDetail
    Parameter = Mki(META_CMD_SET_INTEREST) + Chr$(.byteInterestTotal)
    
    If .byteInterestTotal > 4 Then .byteInterestTotal = 4
    For i = 0 To .byteInterestTotal - 1
      Parameter = Parameter + Mki(.intInterestCategory(i)) + StrAppend(.strInterestName(i))
    Next i
  End With
  SendMeta_UpdateInterest = Asm_String(UDP_CMD_META_USER, Parameter)
End Function

Function SendMeta_UpdateAffiliations(UserDetail As typContactInfo) As Integer
  Dim Parameter As String
  Dim i As Integer
  With UserDetail
    Parameter = Mki(META_CMD_SET_AFFILIATIONS)
    
    Parameter = Parameter + Chr$(.byteBackgroundTotal)
    If .byteBackgroundTotal > 4 Then .byteBackgroundTotal = 4
    For i = 0 To .byteBackgroundTotal - 1
      Parameter = Parameter + Mki(.intBackgroundCategory(i)) + StrAppend(.strBackgroundName(i))
    Next i
    
    Parameter = Parameter + Chr$(.byteOrganizationTotal)
    If .byteOrganizationTotal > 4 Then .byteOrganizationTotal = 4
    For i = 0 To .byteOrganizationTotal - 1
      Parameter = Parameter + Mki(.intOrganizationCategory(i)) + StrAppend(.strOrganizationName(i))
    Next i
  End With
  SendMeta_UpdateAffiliations = Asm_String(UDP_CMD_META_USER, Parameter)
End Function
Function SendMeta_UpdateAbout(UserDetail As typContactInfo) As Integer
  With UserDetail
    SendMeta_UpdateAbout = _
      Asm_String(UDP_CMD_META_USER, Mki(META_CMD_SET_ABOUT) + StrAppend(.strAboutInfo))
  End With
End Function

Function SendMeta_UpdateSecurity(UserDetail As typContactInfo) As Integer
  With UserDetail
    SendMeta_UpdateSecurity = Asm_String(UDP_CMD_META_USER, _
      Mki(META_CMD_SET_SECURE) + _
      Chr$(IIf(.bAuthorize, 0, 1)) + _
      Chr$(IIf(.bWebPresence, 1, 0)) + _
      Chr$(IIf(.bPublishIP, 0, 1)))
  End With
End Function


'-- Search Function --
Function SendMeta_SearchUIN(ByVal uin As Long) As Integer
  SendMeta_SearchUIN = Asm_UIN(UDP_CMD_META_USER, uin, Mki(META_CMD_SEARCH_UIN))
End Function

Function SendMeta_SearchName(ByVal strNick As String, ByVal strFirst As String, ByVal strLast As String) As Integer
  SendMeta_SearchName = _
    Asm_String(UDP_CMD_META_USER, Mki(META_CMD_SEARCH_NAME) + _
    StrAppend(strFirst, strLast, strNick))
End Function

Function SendMeta_SearchEmail(ByVal strEmail As String) As Integer
  SendMeta_SearchEmail = _
    Asm_String(UDP_CMD_META_USER, Mki(META_CMD_SEARCH_NAME) + _
    StrAppend(strEmail))
End Function

