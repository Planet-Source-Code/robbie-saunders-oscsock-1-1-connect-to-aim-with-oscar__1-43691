Attribute VB_Name = "OSCAR"
Dim A1, A2, A3, A4, A5, A6, A7, A8, A9

'login authorization
Public Function AuthLogin(strUser As String, strPassword As String) As String

    A1 = ""
    AuthLogin = A1 & ChrA("0 0 0 1") & TLV(1, strUser) & TLV(2, EncryptPW(strPassword)) & _
        TLV(3, "AOL Instant Messenger, version 5.2.3074/WIN32") & _
        TLV(22, ChrA("0 1")) & TLV(23, ChrA("0 5")) & TLV(24, ChrA("0 1")) & _
        TLV(25, ChrA("9 236")) & TLV(14, "us") & TLV(15, "en") & TLV(9, ChrA("0 9"))

End Function

'login authorization step 2
Public Function authLogin2(strCookie As String) As String

    A1 = ""
    authLogin2 = A1 & ChrA("0 0 0 1") & TLV(6, strCookie)

End Function

'required for login
Public Function addICBMParam() As String

    A1 = ChrA("0 4 0 2 0 0 0 0 0 2")
    addICBMParam = A1 & ChrA("0 0 0 0 0 11 31 64 3 231 3 231 0 0 0 0")
    
End Function

'some mysterious packet
Public Function loginPacket1() As String

    A1 = ChrA("0 1 0 23 0 0 0 0 0 23")
    loginPacket1 = A1 & ChrA("0 1 0 3 0 19 0 3 0 2 0 1 0 3 0 1 0 4 0 1 0 6 0 1 0 8 0 1 0 9 0 1 0 10 0 1 0 11 0 1")
    
End Function

'some mysterious packet
Public Function formatPacket1() As String

    A1 = ChrA("0 1 0 23 0 0 0 0 0 23")
    formatPacket1 = A1 & ChrA("0 1 0 3 0 7 0 1")
    
End Function

'some request packet
Public Function requestPacket1() As String

    A1 = ChrA("0 19 0 2 0 0 0 0 0 2")
    requestPacket1 = A1
    
End Function

'some request packet
Public Function requestPacket2() As String

    A1 = ChrA("0 19 0 5 0 0 0 32 0 5")
    requestPacket2 = A1 & ChrA("62 37 14 13 1 101")
    
End Function

'some request packet
Public Function requestPacket3() As String

    A1 = ChrA("0 2 0 2 0 0 0 0 0 2")
    requestPacket3 = A1
    
End Function

'some request packet
Public Function requestPacket4() As String

    A1 = ChrA("0 4 0 4 0 0 0 0 0 4")
    requestPacket4 = A1
    
End Function

'some request packet
Public Function requestPacket5() As String

    A1 = ChrA("0 19 0 7 0 0 0 0 0 7")
    requestPacket5 = A1
    
End Function

'request rate
Public Function requestRate() As String

    A1 = ChrA("0 1 0 6 0 0 0 0 0 6")
    requestRate = A1
    
End Function

'request personal info
Public Function requestPInfo() As String
    
    A1 = ChrA("0 1 0 14 0 0 0 0 0 14")
    requestPInfo = A1

End Function

'some login packet
Public Function someThing() As String

    A1 = ChrA("0 9 0 2 0 0 0 0 0 2")
    someThing = A1
    
End Function

'request buddy list?
Public Function requestList() As String

    A1 = ChrA("0 3 0 2 0 0 0 0 0 2")
    requestList = A1

End Function

'acknowledge rate
Public Function rateAck() As String

    A1 = ChrA("0 1 0 8 0 0 0 0 0 8")
    rateAck = A1 & ChrA("0 1 0 2 0 3 0 4 0 5")
    
End Function

'request privacy
Public Function requestPrivacy() As String

    A1 = ChrA("0 1 0 20 0 0 0 0 0 0")
    requestPrivacy = A1 & ChrA("0 0 0 3")
    
End Function

'stuff for the search? i dunno
Public Function watcherPacket1() As String

    A1 = ChrA("0 2 0 9 0 0 0 1 0 9")
    watcherPacket1 = A1
    
End Function

'stuff for the search? i dunno
Public Function watcherPacket2() As String

    A1 = ChrA("0 2 0 15 0 0 0 2 0 15")
    watcherPacket2 = A1
    
End Function

'stuff for the search? i dunno
Public Function watcherPacket3(strName As String) As String

    A1 = ChrA("0 2 0 11 0 0 0 3 0 11")
    watcherPacket3 = A1 & Chr(Len(strName)) & strName
    
End Function

'set a profile/capabilities
Public Function changeProfile(strProfileText As String) As String

    A1 = ChrA("0 2 0 4 0 0 0 0 0 0")
    changeProfile = A1 & TLV(1, "text/x-aolrtf; charset=""us-ascii""") & _
        TLV(2, strProfileText) & TLV(5, FULLCAP)

End Function

'set an away message
Public Function awayMessage(strMessage As String) As String
    
    A1 = ChrA("0 2 0 4 0 0 0 0 0 4")
    awayMessage = A1 & TLV(3, "text/aolrtf; charset=" & Chr(34) & "us-ascii" & Chr(34)) & TLV(4, strMessage)

End Function

'send an im
Public Function instantMessage(strRequestID As String, strUserName As String, strMessage As String) As String

    A1 = ChrA("0 4 0 6 0 0 0 10 0 6")
    instantMessage = A1 & strRequestID & ChrA("0 1") & Chr(Len(strUserName)) & strUserName & _
        TLV(2, ChrA("5 1 0 3 1 1 2 1 1") & TwoByteLen(ChrA("0 0 0 0") & strMessage))

End Function

'send a unicode im
Public Function unicodeMessage(strRequestID As String, strUserName As String, strMessage As String) As String

    A1 = ChrA("0 4 0 6 0 0 0 10 0 6")
    unicodeMessage = A1 & strRequestID & ChrA("0 1") & Chr(Len(strUserName)) & strUserName & _
        TLV(2, ChrA("5 1 0 3 1 1 2 1 1") & TwoByteLen(ChrA("0 2 0 0") & strMessage))

End Function

'send a file
Public Function fileSend(strRequestID As String, strUserName As String, strFileName As String) As String

    A1 = ChrA("0 4 0 6 0 0 0 10 0 6")
    fileSend = A1 & strRequestID & ChrA("0 2") & Chr(Len(strUserName)) & strUserName & _
        TLV(5, ChrA("0 0") & strRequestID & ChrA("9 70 19 67 76 127 17 209 130 34 68 69 83 84 0 0 0 10 0 2 0 1 0 15 0 0 0 3 0 4 65 13 107 17 0 5 0 2 20 70 39 17") & TwoByteLen(ChrA("0 2 255 255 255 255 255 110") & strFileName & ChrA("0 1 2 3 4 5 7"))) & ChrA("0 3 0 0")

End Function

'direct connect request
Public Function dcRequest(strRequestID As String, strUserName As String) As String

    A1 = ChrA("0 4 0 6 0 0 0 10 0 6")
    dcRequest = A1 & strRequestID & ChrA("0 2") & Chr(Len(strUserName)) & strUserName & _
        TLV(5, ChrA("0 0") & strRequestID & ChrA("9 70 19 69 76 127 17 209 130 34 68 69 83 84 0 0 0 10 0 2 0 1 0 3 0 4 24 16 172 135 0 5 0 2 20 70 0 15 0 0")) & ChrA("0 3 0 0")

End Function

'game invite
Public Function inviteGame(strRequestID As String, strUserName As String, strGameURL As String, strGameName As String, strMessage As String) As String

    A1 = ChrA("0 4 0 6 0 0 0 10 0 6")
    inviteGame = A1 & strRequestID & ChrA("0 2") & Chr(Len(strUserName)) & strUserName & _
        TLV(5, ChrA("0 0") & strRequestID & ChrA("9 70 19 71 76 127 17 209 130 34 68 69 83 84 0 0 0 10 0 2 0 1 0 15 0 0 0 14") & TwoByteLen("en") & ChrA("0 13") & TwoByteLen("us-ascii") & ChrA("0 12") & TwoByteLen(strMessage) & ChrA("0 3 0 4 64 163 30 79 0 5 0 2 20 70 0 7") & TwoByteLen(strGameURL) & ChrA("39 17") & TwoByteLen(ChrA("0 0 2 0 5 7 76 127 17 209 130 34 68 69 83 84 0 0 0 11 0 9") & strGameName & Chr(0) & "Fuck you" & ChrA("0 0 0 0 0"))) & ChrA("0 3 0 0")

End Function

'chat invite
Public Function inviteChat(strRequestID As String, strUserName As String, strChatRoomURL As String, strMessage As String) As String

    A1 = ChrA("0 4 0 6 0 0 0 10 0 6")
    inviteChat = A1 & strRequestID & ChrA("0 2") & Chr(Len(strUserName)) & strUserName & _
        TLV(5, ChrA("0 0") & strRequestID & ChrA("116 143 36 32 98 135 17 209 130 34 68 69 83 84 0 0 0 10 0 2 0 1 0 15 0 0 0 14") & TwoByteLen("en") & ChrA("0 13") & TwoByteLen("us-ascii") & ChrA("0 12") & TwoByteLen(strMessage) & ChrA("39 17") & TwoByteLen(ChrA("0 4") & Chr(Len(strChatRoomURL)) & strChatRoomURL & ChrA("0 0"))) & ChrA("0 3 0 0")

End Function

'buddy list send
Public Function buddyList(strRequestID As String, strUserName As String, strBuddyList As String) As String

    A1 = ChrA("0 4 0 6 0 0 0 10 0 6")
    buddyList = A1 & strRequestID & ChrA("0 2") & Chr(Len(strUserName)) & strUserName & _
        TLV(5, ChrA("0 0") & strRequestID & ChrA("9 70 19 75 76 127 17 209 130 34 68 69 83 84 0 0 0 10 0 2 0 1 0 15 0 0 39 17") & TwoByteLen(strBuddyList)) & ChrA("0 3 0 0")

End Function

'warn a user
Public Function UserWarning(strUserName As String) As String
    
    A1 = ChrA("0 4 0 8 0 0 0 9 0 8 0 0")
    UserWarning = A1 & Chr(Len(strUserName)) & strUserName

End Function

'block/unblock a user
Public Function userBlock(strUserName As String, intType As Integer) As String
    
    A1 = ChrA("0 19 0 " & intType & " 0 0 0 6 0 " & intType)
    userBlock = A1 & TwoByteLen(strUserName) & ChrA("0 0 11 17 0 3 0 0")

End Function

'talk request
Public Function requestTalk(strRequestID As String, strUserName As String) As String

    A1 = ChrA("0 4 0 6 0 0 0 10 0 6")
    requestTalk = A1 & strRequestID & ChrA("0 2") & Chr(Len(strUserName)) & strUserName & _
        TLV(5, ChrA("0 0") & strRequestID & ChrA("9 70 19 65 76 127 17 209 130 34 68 69 83 84 0 0 0 10 0 2 0 1 0 3 0 4 24 16 172 135 0 15 0 0 39 17 0 4 0 0 0 1")) & ChrA("0 3 0 0")

End Function

'send a theme
Public Function SendTheme(strRequestID As String, strUserName As String, strMessage As String, strThemeName As String) As String

    A1 = ChrA("0 4 0 6 0 0 0 10 0 6")
    SendTheme = A1 & strRequestID & ChrA("0 1") & Chr(Len(strUserName)) & strUserName & _
        TwoByteLen(ChrA("5 1 0 3 1 1 2 1 1") & TwoByteLen(ChrA("0 0 0 0") & strMessage)) & ChrA("0 13") & TwoByteLen(ChrA("0 128") & TwoByteLen(strThemeName) & ChrA("0 130") & TwoByteLen(strThemeName)) & ChrA("0 8 0 12 0 0 12 157 0 1 40 0 60 49 66 80")

End Function

'istyping send
Public Function setTalk(intType As Integer, strUserName As String) As String

    A1 = ChrA("0 4 0 20 0 0 0 0 0 20")
    setTalk = A1 & ChrA("0 0 0 0 0 0 0 0 0 1") & Chr(Len(strUserName)) & strUserName & Chr(0) & Chr(intType)

End Function

'store a comment for a user
Public Function setComment(strName As String, strComment As String) As String

    A1 = ChrA("0 19 0 9 0 0 0 6")
    setComment = A1 & TLV(9, strName) & ChrA("96 147 93 230 0 0 0 7") & TLV(316, strComment)

End Function

'send a stupid email to someone
Public Function inviteFriend(strEmail As String, strMessage As String) As String

    A1 = ChrA("0 6 0 2 0 0 0 1 0 2")
    inviteFriend = A1 & TLV(17, strName) & TLV(21, strMessage)

End Function

'that little `createÿÿ` packet
Public Function chatCreate(strChannel As String) As String

    A1 = ChrA("0 13 0 8 0 0 0 1 0 8 0 4 6")
    chatCreate = A1 & "create" & ChrA("255 255 1 0 3") & TLV(215, "en") & TLV(214, "us-ascii") & TLV(211, strChannel)

End Function

'send something to a chat room
Public Function chatSend(strRequestID As String, strMessage As String) As String

    A1 = ChrA("0 14 0 5 0 0 0 0 0 5")
    chatSend = A1 & strRequestID & ChrA("0 3 0 1 0 0 0 6 0 0") & TLV(5, TLV(2, "us-ascii") & TLV(3, "en") & TLV(1, strMessage))

End Function

'get member info
Public Function getMInfo(strUserName As String) As String
    
    A1 = ChrA("0 2 0 21 0 0 0 21 0 21")
    getMInfo = A1 & ChrA("0 0 0 1") & Chr(Len(strUserName)) & strUserName
    
End Function

'get member info
Public Function nameFormat(strFormattedName As String) As String
    
    A1 = ChrA("0 7 0 4 0 0 0 1 0 4")
    nameFormat = A1 & TLV(1, strFormattedName)
    
End Function

'get member info
Public Function emailChange(strEmail As String) As String
    
    A1 = ChrA("0 7 0 4 0 0 0 1 0 4")
    emailChange = A1 & TLV(17, strEmail)
    
End Function

'format name server change
Public Function servFormat() As String

    A1 = ChrA("0 1 0 4 0 0 0 9 0 4 0 7")
    servFormat = A1

End Function

'new email server change
Public Function servEmail() As String

    A1 = ChrA("0 1 0 4 0 0 0 7 0 4 0 7")
    servEmail = A1

End Function

'chatting server change (1st one)
Public Function servChat() As String

    A1 = ChrA("0 1 0 4 0 0 0 3 0 4 0 13")
    servChat = A1

End Function

'chatting server change (2nd one)
Public Function servChat2(strChatURL As String) As String

    A1 = ChrA("0 1 0 4 0 0 0 4 0 4 0 14")
    servChat2 = A1 & TLV(1, ChrA("0 4") & Chr(Len(strChatURL)) & strChatURL & ChrA("0 0"))

End Function

'final login step
Public Function clientReady() As String

    A1 = ChrA("0 1 0 2 0 0 0 0 2 0")
    clientReady = A1 & ChrA("0 1 0 3 1 16 6 41 0 19 0 3 1 16 6 41 0 2 0 1 1 16 6 41 0 3 0 1 1 16 6 41 0 4 0 1 1 16 6 41 0 6 0 1 1 16 6 41 0 8 0 1 1 16 6 41 0 9 0 1 1 16 6 41 0 10 0 1 1 16 6 41 0 11 0 1 1 4 0 1")
    
End Function

'final format handshake step
Public Function formatReady() As String

    A1 = ChrA("0 1 0 2 0 0 0 0 0 2")
    formatReady = A1 & ChrA("0 1 0 3 0 16 6 41 0 7 0 1 0 16 6 41")
    
End Function

'final email handshake step
Public Function emailReady() As String

    A1 = ChrA("0 1 0 2 0 0 0 0 0 2")
    emailReady = A1 & ChrA("0 1 0 3 0 16 6 208 0 7 0 1 0 16 6 208")
    
End Function

'final chat handshake step
Public Function chatReady() As String

    A1 = ChrA("0 1 0 2 0 0 0 0 0 2")
    chatReady = A1 & ChrA("0 1 0 3 0 16 6 208 0 13 0 1 0 16 6 208")
    
End Function

'final chat handshake step 2
Public Function chatReady2() As String

    A1 = ChrA("0 1 0 2 0 0 0 0 0 2")
    chatReady2 = A1 & ChrA("0 1 0 3 0 16 6 208 0 14 0 1 0 16 6 208")
    
End Function

'final login step(jacked for a ghosting effect)
Public Function clientReadyJacked() As String

    clientReadyJacked = ChrA("0 1 0 2 0 4 2 51 0 2 0 1 0 4 0 1 0 3 0 1 0 4 0 1 0 4 0 1 0 4 0 1 0 6 0 1 0 4 0 1 0 8 0 1 0 4 0 1 0 9 0 1 0 4 0 1 0 10 0 1 0 4 0 1 0 11 0 1 0 4 0 1 0 12 0 1 0 4 0 1")
    
End Function

'add a buddy
Public Function buddyAdd(strUserName As String) As String
    
    A1 = ChrA("0 3 0 4 0 0 0 0 0 0")
    addBuddy = A1 & Chr(Len(strUserName)) & strUserName
    
End Function

'password encryption
Public Function EncryptPW(ByRef strPass As String) As String
    Dim arrTable() As Variant
    Dim strEncrypted As String
    Dim lngX As Long
    Dim strHex As String
    
    arrTable = Array(243, 179, 108, 153, 149, 63, 172, 182, 197, 250, 107, 99, 105, 108, 195, 154)
    
    For lngX = 0 To Len(strPass$) - 1
        strHex = Chr(Asc(Mid(strPass, lngX + 1, 1)) Xor CLng(arrTable((lngX Mod 16))))
        strEncrypted = strEncrypted & strHex
    Next
    
    EncryptPW = strEncrypted
End Function

'capability block
Public Function FULLCAP()
FULLCAP = ChrA("9 70 19 65 76 127 17 209 130 34 68 69 83 84 0 0 9 70 19 67 76 127 17 209 130 34 68 69 83 84 0 0 116 143 36 32 98 135 17 209 130 34 68 69 83 84 0 0 9 70 19 69 76 127 17 209 130 34 68 69 83 84 0 0 9 70 19 70 76 127 17 209 130 34 68 69 83 84 0 0 9 70 19 71 76 127 17 209 130 34 68 69 83 84 0 0 9 70 19 72 76 127 17 209 130 34 68 69 83 84 0 0 9 70 19 72 76 127 17 209 130 34 68 69 83 84 0 0")
'Buddy Icon: 9 70 19 70 76 127 17 209 130 34 68 69 83 84 0 0
'File Send:  9 70 19 67 76 127 17 209 130 34 68 69 83 84 0 0
'Buddy List: 9 70 19 75 76 127 17 209 130 34 68 69 83 84 0 0
'Talk:       9 70 19 65 76 127 17 209 130 34 68 69 83 84 0 0
'Games:      9 70 19 71 76 127 17 209 130 34 68 69 83 84 0 0
'Chat:       116 143 36 32 98 135 17 209 130 34 68 69 83 84 0 0
'IM Image:   9 70 19 69 76 127 17 209 130 34 68 69 83 84 0 0
'Get File:   9 70 19 72 76 127 17 209 130 34 68 69 83 84 0 0
'65
'67
'32
'69
'70
'71
End Function

Public Function TLV(intType As Integer, strData As String) As String
    TLV = IntegerToBase256(intType) & TwoByteLen(strData)
End Function

