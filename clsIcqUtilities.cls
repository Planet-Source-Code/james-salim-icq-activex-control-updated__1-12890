VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIcqUtilities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum enumInfoGender
  icqMale = 2
  icqFemale = 1
  icqNotSpecified = 0
End Enum

Public Enum enumMessageType
  icqMsgText = &H1
  icqMsgChatReq = &H2
  icqMsgFile = &H3
  icqMsgURL = &H4
  icqMsgAuthReq = &H6
  icqMsgAuthDecline = &H7
  icqMsgAuthAccept = &H8
  icqMsgAdded = &HC
  icqMsgWebpager = &HD
  icqMsgExpress = &HE
  icqMsgContact = &H13
End Enum

Public Enum enumOnlineState
  icqOnline = &H0
  icqAway = &H1
  icqNa = &H5
  icqOccupied = &H11
  icqDND = &H13
  icqChat = &H20
  icqInvisible = &H100
End Enum

Public Enum enumUseTCP
  icqNoTCP = 1
  icqTCPSendOnly = 2
  icqTCPSendRecv = 4
End Enum

Public Enum enumRandomGroup
  icqGrpGeneral = 1
  icqGrpRomance = 2
  icqGrpGames = 3
  icqGrpStudents = 4
  icqGrp20Something = 6
  icqGrp30Something = 7
  icqGrp40Something = 8
  icqGrp50Over = 9
  icqGrpManRequestWoman = 10
  icqGrpWomanRequestMan = 11
End Enum

Public Enum enumConnectionSate
  icqDisconnected = 0
  icqRegistering = 1
  icqLogin = 2
  icqConnected = 3
End Enum

Public Enum enumSearchResult
  icqSearchUserFound = 0
  icqSearchDone = 1
  icqSearchTooMany = 2
End Enum

Public Enum enumInfoType
  icqNewUser = 0
  icqBasic = 1
  icqMain = 2
  icqMore = 3
  icqMetaMore = 11
  icqWork = 4
  icqInterest = 5
  icqAffiliations = 6
  icqAbout = 7
  icqSecurity = 8
  icqHPCategory = 9
  icqall = 10
End Enum

Public Enum enumErrorConstant
  icqErrNotConnected = 1
  icqErrWrongPassword = 2
  icqErrTryAgain = 3
  icqErrGoAway = 4
  icqErrInvalidUIN
End Enum

Dim CountryName As Variant
Dim CountryCode As Variant
Dim Language As Variant
Dim Occupation As Variant
Dim PastBackground As Variant
Dim Affiliations As Variant
Private Sub Class_Initialize()
  'CountryName holds the Country name, while CountryCode holds the country code
  'number. Total is 122 (0 to 121)
  CountryName = Array( _
    "USA", "Russia", "Egypt", "South Africa", "Greece", "Netherlands", "Belgium", _
    "France", "Monaco", "Spain", "Hungary", "Yugoslavia", "Italy", "San Marino", _
    "Vatican City", "Romania", "Liechtenstein", "Switzerland", "Czech Republic", _
    "Austria", "UK", "Denmark", "Sweden", "Norway", "Poland", "Germany", "Peru", _
    "Mexico", "Guantanomo Bay", "Argentina", "Brazil", "Chile", "Columbia", _
    "Venezuela", "Malaysia", "Australia", "Indonesia", "Philippines", "New Zealand", _
    "Singapore", "Thailand", "Japan", "South Korea", "Vietnam", "China", "Turkey", _
    "India", "Pakistan", "Sri Lanka", "Iran", "Canada", "Morocco", "Algeria", _
    "Tunisia", "Libya", "Senegal", "Mali", "Ivory Coast", "Liberia", "Ghana", "Nigeria", _
    "Cameroon", "Gabon", "Zaire", "Ethiopia", "Kenya", "Tanzania", "Zimbabwe", "Namibia", _
    "Malawi", "Aruba", "Portugal", "Luxembourg", "Ireland", "Iceland", "Malta", "Cyprus", _
    "Finland", "Bulgaria", "Ukraine", "Belize", "Guatemala", "El Salvador", "Honduras", _
    "Nicaragua", "Costa Rice", "Panama", "Haiti", "Guadeloupe", "Bolivia", "Guyana", "Ecuador", _
    "Paraguay", "French Antilles", "Suriname", "Uruguay", "Netherlands Antilles", "Saipan", _
    "Saipan", "Guam", "Papua New Guinea", "Fiji", "American Samoa", "New Caledonia", _
    "French Polynesia", "Hong Kong", "Trinidad and Tobago", "Bangladesh", "Taiwan", _
    "Jordan", "Iraq", "Kuwait", "Saudia Arabia", "Yemen", "Oman", "United Arab Emirates", _
    "Israel", "Bahrain", "Qatar", "Nepal", "Slovak Republic", "Not specified")
  
  CountryCode = Array( _
    1, 7, 20, 27, 30, 31, 32, 33, 33, 34, 36, 38, 39, 39, 39, 40, _
    41, 41, 42, 43, 44, 45, 46, 47, 48, 49, 51, 52, 53, 54, 55, 56, _
    57, 58, 60, 61, 62, 63, 64, 65, 66, 81, 82, 84, 86, 90, 91, 92, _
    94, 98, 107, 212, 213, 216, 218, 221, 223, 225, 231, 233, 234, _
    237, 241, 243, 251, 254, 255, 263, 264, 265, 297, 351, 352, 353, _
    354, 356, 357, 358, 359, 380, 501, 502, 503, 504, 505, 506, 507, _
    509, 590, 591, 592, 593, 595, 596, 597, 598, 599, 670, 670, 671, _
    675, 679, 684, 687, 689, 852, 868, 880, 886, 962, 964, 965, 966, _
    967, 968, 971, 972, 973, 974, 977, 4201, &HFFFF)
  
  'Occupation - Total 28
  Occupation = Array( _
    "Academic", 1, "Administrative", 2, "Art/Entertainment", 3, "College Student", 4, _
    "Computers", 5, "Community & Social", 6, "Education", 7, "Engineering", 8, _
    "Financial Services", 9, "Government", 10, "High School Student", 11, "Home", 12, _
    "ICQ - Providing Help", 13, "Law", 14, "Managerial", 15, "Manufacturing", 16, _
    "Medical/Health", 17, "Military", 18, "Non-Government Organization", 19, "Professional", 20, _
    "Retail", 21, "Retired", 22, "Science & Research", 23, "Sports", 24, _
    "Technical", 25, "University Student", 26, "Web Building", 27, "Other Services", 99)
    
  'Past Background - Total 8
  PastBackground = Array( _
    "Elementary School", 300, "High School", 301, "College", 302, "University", 303, _
    "Military", 304, "Past Work Place", 305, "Past Organization", 306, "Other", 399)

  'Affiliations - Total 20
  Affiliations = Array( _
    "Alumni Org.", 200, "Charity Org.", 201, "Club/Social Org.", 202, "Community Org.", 203, _
    "Cultural Org.", 204, "Fan Clubs", 205, "Fraternity/Sorority", 206, "Hobbyists Org.", 207, _
    "International Org.", 208, "Nature and Environment Org.", 209, "Professional Org.", 210, _
    "Scientific/Technical Org.", 211, "Self Improvement Group", 212, "Spiritual/Religious Org.", 213, _
    "Sports Org.", 214, "Support Org.", 215, "Trade and Business Org.", 216, "Union", 217, _
    "Voluntary Org.", 218, "Other", 299)

  'Language - Total 34
  Language = Array( _
    "Not Entered", "Arabic", "Bhojpuri", "Bulgarian", "Burmese", "Cantonese", "Catalan", "Chinese", _
    "Croatian", "Czech", "Danish", "Dutch", "English", "Esperanto", "Estonian", "Farsi", "Finnish", _
    "French", "Gaelic", "German", "Greek", "Hebrew", "Hindi", "Hungarian", "Icelandic", "Indonesian", _
    "Italian", "Japanese", "Khmer", "Korean", "Lao", "Latvian", "Lithuanian", "Malay", "Norwegian", _
    "Polish", "Portuguese", "Romanian", "Russian", "Serbian", "Slovak", "Slovenian", "Somali", _
    "Spanish", "Swahili", "Swedish", "Tagalog", "Tatar", "Thai", "Turkish", "Ukrainian", "Urdu", _
    "Vietnamese", "Yiddish", "Yoruba", "Afrikaans", "Bosnian", "Persian", "Albanian", "Armenian", _
    "Punjabi", "Chamorro", "Mongolian", "Mandarin", "Taiwaness", "Macedonian", "Sindhi", "Welsh")
End Sub

'--- Country Conversion ---
Public Function GetCountryName(Code As Integer) As String
  Dim Index As Integer
  On Error GoTo ErrorProc
  Index = GetCountryIndex(Code)
  GetCountryName = CountryName(Index)
ErrorProc:
End Function
Public Function GetCountryCode(Index As Integer) As Integer
  On Error GoTo ErrorProc
  GetCountryCode = CountryCode(Index)
ErrorProc:
End Function
Public Function GetCountryIndex(CntryCode As Integer) As Integer
  Dim i As Integer
  GetCountryIndex = &HFFFF

  For i = 0 To UBound(CountryCode)
    If CountryCode(i) = CntryCode Then
      GetCountryIndex = i
      Exit For
    End If
  Next i
End Function

'--- Language Conversion ---
Public Function GetLangName(Index As Integer) As String
  On Error GoTo ErrorProc
  GetLangName = Language(Index)
ErrorProc:
End Function

'--- Occupation Conversion ---
Public Function GetOccupationName(Code As Integer) As String
  Dim Index As Integer
  On Error GoTo ErrorProc
  Index = GetOccupationIndex(Code)
  GetOccupationName = Occupation(Index * 2)
ErrorProc:
End Function
Public Function GetOccupationCode(Index As Integer) As Integer
  On Error GoTo ErrorProc
  GetOccupationCode = Occupation(Index * 2 + 1)
ErrorProc:
End Function
Public Function GetOccupationIndex(OccupationCode As Integer) As Integer
  Dim i As Integer
  GetOccupationIndex = -1
  
  For i = 1 To UBound(Occupation) Step 2
    If Occupation(i) = OccupationCode Then
      GetOccupationIndex = (i - 1) / 2
      Exit For
    End If
  Next i
End Function

'--- PastBackground Conversion ---
Public Function GetPastBackgroundName(Code As Integer) As String
  Dim Index As Integer
  On Error GoTo ErrorProc
  Index = GetPastBackgroundIndex(Code)
  GetPastBackgroundName = PastBackground(Index * 2)
ErrorProc:
End Function
Public Function GetPastBackgroundCode(Index As Integer) As Integer
  On Error GoTo ErrorProc
  GetPastBackgroundCode = PastBackground(Index * 2 + 1)
ErrorProc:
End Function
Public Function GetPastBackgroundIndex(PastBackgroundCode As Integer) As Integer
  Dim i As Integer
  GetPastBackgroundIndex = -1
  
  For i = 1 To UBound(PastBackground) Step 2
    If PastBackground(i) = PastBackgroundCode Then
      GetPastBackgroundIndex = (i - 1) / 2
      Exit For
    End If
  Next i
End Function

'--- Affiliations Conversion ---
Public Function GetAffiliationsName(Code As Integer) As String
  Dim Index As Integer
  On Error GoTo ErrorProc
  Index = GetAffiliationsIndex(Code)
  GetAffiliationsName = Affiliations(Index * 2)
ErrorProc:
End Function
Public Function GetAffiliationsCode(Index As Integer) As Integer
  On Error GoTo ErrorProc
  GetAffiliationsCode = Affiliations(Index * 2 + 1)
ErrorProc:
End Function
Public Function GetAffiliationsIndex(AffiliationsCode As Integer) As Integer
  Dim i As Integer
  GetAffiliationsIndex = -1
  
  For i = 1 To UBound(Affiliations) Step 2
    If Affiliations(i) = AffiliationsCode Then
      GetAffiliationsIndex = (i - 1) / 2
      Exit For
    End If
  Next i
End Function


'--- Time Zone Conversion ---
Public Function GetTimeZone(ByVal Code As Byte) As String
  If Code <= 24 Then
    GetTimeZone = "GMT -"
    GetTimeZone = GetTimeZone & Format(Int(Code / 2), "00") & ":"
    GetTimeZone = GetTimeZone & IIf((Code Mod 2) = 1, "30", "00")
  End If
  
  If Code >= 230 Then
    Code = 256 - Code
    GetTimeZone = "GMT +"
    GetTimeZone = GetTimeZone & Format(Int(Code / 2), "00") & ":"
    GetTimeZone = GetTimeZone & IIf((Code Mod 2) = 1, "30", "00")
  End If
End Function

' Input Format is as GetTimeZone output format
Public Function GetTimeCode(ByVal strTime As String) As Byte
  Dim bPositive As Boolean
  Dim TimeHour As Byte
  Dim bHalfHour As Boolean
  
  bPositive = IIf(Mid$(strTime, 5, 1) = "+", True, False)
  TimeHour = CByte(Mid$(strTime, 6, 2))
  bHalfHour = IIf(Mid$(strTime, 8, 2) = "00", False, True)
  
  GetTimeCode = TimeHour * 2 + IIf(bHalfHour, 1, 0)
  If bPositive = False Then GetTimeCode = 256 - GetTimeCode
End Function
