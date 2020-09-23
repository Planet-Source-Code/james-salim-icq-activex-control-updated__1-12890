Attribute VB_Name = "udpMetaRecv"
Option Explicit

Sub RecvMeta_Handle(ByVal Parameter As String)
  Dim UserDetail As typContactInfo
  Dim SubCMD As Integer
  Dim MetaResult As Byte
  Dim i As Integer

  SubCMD = PacketRead(Parameter, vbInteger)
  MetaResult = PacketRead(Parameter, vbByte)
  
  If MetaResult = META_SRV_FAILURE Then
    'DebugTxt "RcvMETA", "Failed. SubCMD " + Hex$(SubCMD)
  Else
    With UserDetail
    Select Case SubCMD
      Case META_SRV_SEARCH_FOUND
        .lngUIN = PacketRead(Parameter, vbLong)
        .strNickname = PacketRead(Parameter, vbString)
        .strFirstName = PacketRead(Parameter, vbString)
        .strLastName = PacketRead(Parameter, vbString)
        .strEmail = PacketRead(Parameter, vbString)
        .bAuthorize = IIf(PacketRead(Parameter, vbByte), False, True)
        .bWebPresence = IIf(PacketRead(Parameter, vbByte), True, False)
        Events_SearchFound UserDetail
        
      Case META_SRV_SEARCH_LAST
        .lngUIN = PacketRead(Parameter, vbLong)
        .strNickname = PacketRead(Parameter, vbString)
        .strFirstName = PacketRead(Parameter, vbString)
        .strLastName = PacketRead(Parameter, vbString)
        .strEmail = PacketRead(Parameter, vbString)
        .bAuthorize = IIf(PacketRead(Parameter, vbByte), False, True)
        .bWebPresence = IIf(PacketRead(Parameter, vbByte), True, False)
        i = PacketRead(Parameter, vbLong)     '??? = Users Left (0=Done; ?? = TooMany)
        If .lngUIN <> 0 Then Events_SearchFound UserDetail
        If i = 0 Then
          Events_SearchDone False
        Else
          Events_SearchDone True
        End If
        
      '--------------------'
      ' Info Request Reply '
      '--------------------'
      Case META_SRV_USER_INFO
        .strNickname = PacketRead(Parameter, vbString)
        .strFirstName = PacketRead(Parameter, vbString)
        .strLastName = PacketRead(Parameter, vbString)
        .strEmail = PacketRead(Parameter, vbString)
        .strEmail2 = PacketRead(Parameter, vbString)
        .strEmail3 = PacketRead(Parameter, vbString)
        
        .strCity = PacketRead(Parameter, vbString)
        .strState = PacketRead(Parameter, vbString)
        .strPhone = PacketRead(Parameter, vbString)
        .strFax = PacketRead(Parameter, vbString)
        .strStreet = PacketRead(Parameter, vbString)
        .strCellular = PacketRead(Parameter, vbString)
        .lngZip = PacketRead(Parameter, vbLong)
        .intCountryCode = PacketRead(Parameter, vbInteger)
        .byteTimeZone = PacketRead(Parameter, vbByte)

        .bAuthorize = IIf(PacketRead(Parameter, vbByte), False, True)
        .bWebPresence = IIf(PacketRead(Parameter, vbByte), True, False)
        .bPublishIP = IIf(PacketRead(Parameter, vbByte), False, True)
        Events_InfoReply icqMain, UserDetail
        
      Case META_SRV_USER_WORK
        .strWorkCity = PacketRead(Parameter, vbString)
        .strWorkState = PacketRead(Parameter, vbString)
        .strWorkPhone = PacketRead(Parameter, vbString)
        .strWorkFax = PacketRead(Parameter, vbString)
        .strWorkAddress = PacketRead(Parameter, vbString)
        .lngWorkZip = PacketRead(Parameter, vbLong)
        .intWorkCountry = PacketRead(Parameter, vbInteger)
        .strWorkName = PacketRead(Parameter, vbString)
        .strWorkDepartment = PacketRead(Parameter, vbString)
        .strWorkPosition = PacketRead(Parameter, vbString)
        .intWorkOccupation = PacketRead(Parameter, vbInteger)
        .strWorkWebURL = PacketRead(Parameter, vbString)
        Events_InfoReply icqWork, UserDetail
        
      Case META_SRV_USER_MORE
        .intAge = PacketRead(Parameter, vbInteger)
        .byteGender = PacketRead(Parameter, vbByte)
        .strHomepageURL = PacketRead(Parameter, vbString)
        .byteBirthYear = PacketRead(Parameter, vbByte)
        .byteBirthMonth = PacketRead(Parameter, vbByte)
        .byteBirthDay = PacketRead(Parameter, vbByte)
        .byteLanguage1 = PacketRead(Parameter, vbByte)
        .byteLanguage2 = PacketRead(Parameter, vbByte)
        .byteLanguage3 = PacketRead(Parameter, vbByte)
        Events_InfoReply icqMetaMore, UserDetail
      
      Case META_SRV_USER_ABOUT
        .strAboutInfo = PacketRead(Parameter, vbString)
        Events_InfoReply icqAbout, UserDetail
      
      Case META_SRV_USER_INTERESTS
        .byteInterestTotal = PacketRead(Parameter, vbByte)
        If .byteInterestTotal > 4 Then .byteInterestTotal = 4
        For i = 0 To .byteInterestTotal - 1
          .intInterestCategory(i) = PacketRead(Parameter, vbInteger)
          .strInterestName(i) = PacketRead(Parameter, vbString)
        Next i
        Events_InfoReply icqInterest, UserDetail
      
      Case META_SRV_USER_AFFILIATIONS
        .byteBackgroundTotal = PacketRead(Parameter, vbByte)
        If .byteBackgroundTotal > 4 Then .byteBackgroundTotal = 4
        For i = 0 To .byteBackgroundTotal - 1
          .intBackgroundCategory(i) = PacketRead(Parameter, vbInteger)
          .strBackgroundName(i) = PacketRead(Parameter, vbString)
        Next i
          
        .byteOrganizationTotal = PacketRead(Parameter, vbByte)
        If .byteOrganizationTotal > 4 Then .byteOrganizationTotal = 4
        For i = 0 To .byteOrganizationTotal - 1
          .intOrganizationCategory(i) = PacketRead(Parameter, vbInteger)
          .strOrganizationName(i) = PacketRead(Parameter, vbString)
        Next i
        Events_InfoReply icqAffiliations, UserDetail
    
      Case META_SRV_USER_HPCATEGORY
        .byteHPCategoryTotal = PacketRead(Parameter, vbByte)
        If .byteHPCategoryTotal > 4 Then .byteHPCategoryTotal = 4
        For i = 0 To .byteHPCategoryTotal - 1
          .intHPCategoryCategory(i) = PacketRead(Parameter, vbInteger)
          .strHPCategoryName(i) = PacketRead(Parameter, vbString)
        Next i
        Events_InfoReply icqHPCategory, UserDetail

      Case META_SRV_ACK_INFO
      Case META_SRV_ACK_HOMEPAGE
      Case META_SRV_ACK_ABOUT
      Case META_SRV_ACK_SECURE
      Case META_SRV_ACK_PASS
      Case Else
    End Select
    End With
  End If
End Sub

