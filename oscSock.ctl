VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.UserControl oscSock 
   ClientHeight    =   4830
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6165
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   4830
   ScaleWidth      =   6165
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   30000
      Left            =   3120
      Top             =   1560
   End
   Begin VB.TextBox SBYTE 
      Height          =   285
      Index           =   1
      Left            =   2040
      TabIndex        =   1
      Text            =   "0"
      Top             =   2280
      Width           =   615
   End
   Begin VB.TextBox SBYTE 
      Height          =   285
      Index           =   0
      Left            =   1320
      TabIndex        =   0
      Text            =   "0"
      Top             =   2280
      Width           =   615
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   1
      Left            =   2160
      Top             =   2760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   0
      Left            =   1680
      Top             =   2760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label3 
      Height          =   375
      Index           =   0
      Left            =   3240
      TabIndex        =   4
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Label Label2 
      Height          =   255
      Index           =   0
      Left            =   3240
      TabIndex        =   3
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "1"
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   1320
      Width           =   495
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00E0E0E0&
      X1              =   1500
      X2              =   1500
      Y1              =   0
      Y2              =   360
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00E0E0E0&
      X1              =   0
      X2              =   1560
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   1500
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   405
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   0
      Picture         =   "oscSock.ctx":0000
      Top             =   0
      Width           =   1500
   End
End
Attribute VB_Name = "oscSock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'all code except the password encryption by robbie saunders
'version 1.1: login recoded to mimick latest aim (5.2), too much other stuff added

'Data Storage Index Chart
'0 - username
'1 - password
'2 - formatted username
'3 - authorization cookie
'4 - email address
'5 - temp channel storage
'8 - type of add-on service
'9 - add-on auth. cookie

Dim TheData(255) As String, IncomingStuff(255) As String, PacketLen(255), LeftOver(255) As Boolean
Dim SHAFT(255), A1, A2, A3, A4, A5, A6, A7, A8, A9, LoggedOn As Boolean
Public Ghosted As Boolean
Public Event loggedIn(strFormattedName As String, strEmailAddy As String)
Public Event incomingProfile(strFormattedName As String, lngSecondsOnline As Long, intWarningLevel As Integer, strProfile As String)
Public Event incomingIM(strName As String, strMessage As String)
Public Event incomingChat(intIndex As Integer, strName As String, strMessage As String)
Public Event incomingPacket(intIndex As Integer, strPacket As String)
Public Event buddySignedOn(strName As String, strCapabilities As String)
Public Event buddySignedOff(strName As String)
Public Event chatReady(strChannel As String, intIndex As Integer)
Public Event chatBuddyEntered(strName As String)
Public Event chatBuddyLeft(strName As String)

'begin the login process
Public Sub loginUser(strUserName As String, strPassword As String)
    
    'store username and pass
    pData 0, strUserName
    pData 1, strPassword
    'we're not logged on yet
    LoggedOn = False
    'connect
    Winsock1(0).Close
    Winsock1(0).Connect "login.oscar.aol.com", "5190"
    'stop the anti-idle
    Timer1.Enabled = False
    
End Sub

'logout
Public Sub logOut()
    
    'close the connection
    Winsock1(0).Close
    Winsock1(1).Close
    'stop the anti-idle
    Timer1.Enabled = False
    
End Sub

'let user change profile
Public Sub setProfile(strProfileText As String)
    
    sendPacket 1, changeProfile(strProfileText)
    
End Sub

'set away status
Public Sub setAway(strAwayMessage As String)

    sendPacket 1, awayMessage(strAwayMessage)

End Sub

'send a basic im
Public Sub sendIM(strUserName As String, strMessage As String)

    sendPacket 1, instantMessage(String(8, Chr(GRInteger(0, 255))), strUserName, strMessage)

End Sub

'send a unicode im
Public Sub sendUnicodeIM(strUserName As String, strMessage As String)

    sendPacket 1, unicodeMessage(String(8, Chr(GRInteger(0, 255))), strUserName, strMessage)

End Sub

'send a file?
Public Sub sendFile(strUserName As String, strFileName As String)

    sendPacket 1, fileSend(String(6, GRInteger(0, 255)) & ChrA("0 0"), strUserName, strFileName)

End Sub

'direct connect?
Public Sub directConnect(strUserName As String)

    sendPacket 1, dcRequest(String(6, GRInteger(0, 255)) & ChrA("0 0"), strUserName)

End Sub

'invite to a game?
Public Sub gameInvite(strUserName As String, strGameName As String, strMessage As String)

    Dim gameURL
    gameURL = "aim:AddGame?name=NetMeeting&go1st=true&multiplayer=true&url=http://www.microsoft.com/windows/netmeeting/&cmd=msconf.dll,CallToProtocolHandler+%25i&servercmd=c:%5Cprogra~1%5CNetMeeting%5Cconf.exe&hint=If+it+takes+a+long+time+for+your+buddy+to+connect%3CBR%3Eto+you,+disable+the+following+preference+in+NetMeeting:%3CBR%3E'Log+on+to+a+directory+server+when+NetMeeting+starts'.'"
    sendPacket 1, inviteGame(String(6, GRInteger(0, 255)) & ChrA("0 0"), strUserName, CStr(gameURL), strGameName, strMessage)

End Sub

'invite to a chat room?
Public Sub chatInvite(strUserName As String, intChatExchange As Integer, strChatName As String, strMessage As String)

    Dim chatURL
    chatURL = "!aol://2719:10-" & intChatExchange & "-" & strChatName
    sendPacket 1, inviteChat(String(6, GRInteger(0, 255)) & ChrA("0 0"), strUserName, CStr(chatURL), strMessage)

End Sub

'send your buddy list
Public Sub blistSend(strUserName As String, strBuddyList As String)

    sendPacket 1, buddyList(String(6, GRInteger(0, 255)) & ChrA("0 0"), strUserName, strBuddyList)

End Sub

'talk?
Public Sub talkRequest(strUserName As String)

    sendPacket 1, requestTalk(String(6, GRInteger(0, 255)) & ChrA("0 0"), strUserName)

End Sub

'block/unblock user
Public Sub blockUser(strUserName As String, boolBlock As Boolean)
    
    If boolBlock Then
        sendPacket 1, userBlock(strUserName, 8)
    Else
        sendPacket 1, userBlock(strUserName, 10)
    End If
    
End Sub

'warn user
Public Sub warnUser(strUserName As String)
    
    sendPacket 1, UserWarning(strUserName)
        
End Sub

'get someone's profile
Public Sub getProfile(strUserName As String)
    
    sendPacket 1, getMInfo(strUserName)
        
End Sub

'set 5.0+ `typing` attribute
Public Sub isTyping(strUserName As String, boolTyping As Boolean)
    
    If boolTyping = True Then
        sendPacket 1, setTalk(2, strUserName)
    Else
        sendPacket 1, setTalk(0, strUserName)
    End If
        
End Sub

'send an aim expression
Public Sub sendExpression(strUserName As String, strMessage As String, strThemeName As String)

    sendPacket 1, SendTheme(String(6, GRInteger(0, 255)) & ChrA("0 0"), strUserName, strMessage, strThemeName)

End Sub

'send something to a chat room
Public Sub sendChat(intIndex As Integer, strMessage As String)

    sendPacket2 2, intIndex, chatSend(String(8, Chr(GRInteger(0, 255))), strMessage)

End Sub

'store a buddy comment
Public Sub storeComment(strUserName As String, strComment As String)

    sendPacket 1, setComment(strUserName, strComment)
    
End Sub

'send one of those "invitational" emails
Public Sub signUpFriend(strEmail As String, strMessage As String)

    sendPacket 1, inviteFriend(strEmail, strMessage)
    
End Sub

'format your screen name
Public Sub formatSN(strFormattedName As String)

    sendPacket 1, servFormat
    pData 8, "format"
    pData 2, strFormattedName
    
End Sub

'change your email
Public Sub changeEmail(strEmail As String)

    sendPacket 1, servEmail
    pData 8, "email"
    pData 4, strEmail
    
End Sub

'begin the long process of joining a chat room
Public Sub joinChat(strChannel As String)

    sendPacket 1, servChat
    pData 8, "chat"
    pData 5, strChannel
    
End Sub

'end an add in (like chatting)
Public Sub endAddIn(intIndex As Integer)

    On Error Resume Next 'make sure dumb people don't try and close controls that don't exist
    Winsock1(intIndex).Close
    Unload Winsock1(intIndex)
    
End Sub

'add buddy
Public Sub addBuddy(strUserName As String)

    sendPacket 1, buddyAdd(strUserName)

End Sub

'send a packet
Public Sub sendPacket(Index As Integer, theStuff As String)
    
    If Winsock1(Index).State = sckConnected Then
    
        SBYTE(Index) = SBYTE(Index) + 1
        If SBYTE(Index) > 65535 Then SBYTE(Index) = 0
        
        Winsock1(Index).SendData "*" & Chr(Index + 1) & IntegerToBase256(SBYTE(Index)) & IntegerToBase256(Len(theStuff)) & theStuff
        
    Else
        'uh oh
    End If

End Sub

'send another gay one
Public Sub sendPacket2(intType As Integer, Index, theStuff As String)
    
    If Winsock1(Index).State = sckConnected Then
    
        SBYTE(Index) = SBYTE(Index) + 1
        If SBYTE(Index) > 65535 Then SBYTE(Index) = 0
        
        Winsock1(Index).SendData "*" & Chr(intType) & IntegerToBase256(SBYTE(Index)) & IntegerToBase256(Len(theStuff)) & theStuff
        DoEvents
        
    Else
        'uh oh
    End If

End Sub

'little anti-idler
Private Sub Timer1_Timer()

    sendPacket2 5, 1, ""
    DoEvents
    
End Sub

Private Sub UserControl_Resize()
UserControl.Width = 1515
UserControl.Height = 495
End Sub

'parse all the packets apart
Private Sub Winsock1_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    
    On Error Resume Next
    
    Winsock1(Index).GetData IncomingStuff(Index) 'grab new info
    
    TheData(Index) = TheData(Index) & IncomingStuff(Index) 'add it to the buffer

NextOne:

    If LeftOver(Index) = False Then PacketLen(Index) = GetLength(Chr(0) & Mid(TheData(Index), 5, 2)) + 6 'grab the instruction length
    
    If PacketLen(Index) > Len(TheData(Index)) Then
        LeftOver(Index) = True 'make sure we don't regrab the length bytes and slow us down
        Exit Sub
    End If
    
    processData Index, Left(TheData(Index), PacketLen(Index)) 'process the instruction
    TheData(Index) = Right(TheData(Index), Len(TheData(Index)) - PacketLen(Index)) 'remove it from the buffer
    PacketLen(Index) = 0
    LeftOver(Index) = False 'enable length grabbing
    
    If Len(TheData(Index)) >= 6 Then GoTo NextOne: 'if we have the complete header of the next packet then keep going

End Sub

Public Sub processData(intIndex, strData)
    
    RaiseEvent incomingPacket(CInt(intIndex), CStr(strData))

    Select Case intIndex
    
        Case 0 'login stuff
            
            Select Case Asc(Mid(strData, 2, 1))
            
                Case 1 'auth request

                    sendPacket 0, AuthLogin(CStr(gData(0)), CStr(gData(1)))
                
                Case 4 'disconnect packet
                    
                    pData 2, getTLV(strData, 1) 'store formatted sn
                    A1 = Split(getTLV(strData, 5), ":", 2) 'get server/port
                    pData 3, getTLV(strData, 6) 'store auth cookie
                    pData 4, getTLV(strData, 17) 'store email address
                    Winsock1(1).Close
                    Winsock1(1).Connect A1(0), A1(1)
                    
            End Select
            
        Case 1 'BOS stuff
            
            'give the server your auth. cookie
            If Asc(Mid(strData, 2, 1)) = 1 Then
                sendPacket2 1, 1, authLogin2(CStr(gData(3)))
            End If
        
            'generate a unique long value from the SNAC header
            Select Case GetLength(Chr(0) & Mid(strData, 8, 1) & Mid(strData, 10, 1))
            
                Case 259 'something weird
                    
                    sendPacket 1, loginPacket1
                    
                Case 261 'server redirect
                
                    Label1 = Label1 + 1
                    Load Winsock1(Label1)
                    Load Label2(Label1) 'store service type
                    Load Label3(Label1) 'i can't remember hah
                    Load SBYTE(Label1) 'store sbyte
                    SBYTE(Label1) = "0"
                    strData = Right(strData, Len(strData) - 30)
                    Winsock1(Label1).Close
                    Winsock1(Label1).Connect getTLV(strData, 5), "5190"
                    pData 9, getTLV(strData, 6)
                    Label2(Label1) = gData(8)
                    Label3(Label1) = "0"
                    
                Case 263 'request more crap
                
                    sendPacket 1, rateAck
                    DoEvents
                    sendPacket 1, requestPacket1
                    sendPacket 1, requestPacket2
                    sendPacket 1, requestPacket3
                    sendPacket 1, requestPacket4
                    sendPacket 1, requestPInfo
                    sendPacket 1, someThing
                    sendPacket 1, requestList
                    
                Case 280 'request rate information
                
                    sendPacket 1, requestRate
                    
                Case 518 'incoming profile
                
                    A1 = Asc(Mid(strData, 17, 1))
                    A2 = Mid(strData, 18, A1) 'name
                    A3 = getTLV(Right(strData, Len(strData) - (40 + A1)), 2) 'profile
                    A4 = GetLength(Mid(strData, 33 + A1, 3)) 'time online
                    A5 = GetLength(Chr(0) & Mid(strData, 18 + A1, 2)) / 10 'warning level
                    RaiseEvent incomingProfile(CStr(A2), CLng(A4), CInt(A5), CStr(A3))
                    
                Case 779 'buddy signing on
                
                    A1 = Asc(Mid(strData, 17, 1))
                    A2 = Mid(strData, 18, A1) 'name
                    A3 = getTLV(Right(strData, Len(strData) - (19 + A1)), 13) 'capability block
                    RaiseEvent buddySignedOn(CStr(A2), CStr(A3))
                    
                Case 780 'buddy signing off
                
                    A1 = Asc(Mid(strData, 17, 1))
                    A2 = Mid(strData, 18, A1) 'name
                    RaiseEvent buddySignedOff(CStr(A2))
                    
                Case 1031 'incoming im/whatever
                
                    A1 = Asc(Mid(strData, 27, 1))
                    A2 = Mid(strData, 28, A1) 'name
                    If Mid(strData, 26, 1) = Chr(1) Then 'no attachment (regular im)
                        A3 = InStr(1, strData, ChrA("3 1 1 2 1 1"))
                        If A3 <> 0 Then
                            A4 = Right(strData, Len(strData) - (A3 + 5))
                            A5 = GetLength(Chr(0) & Mid(A4, 1, 2))
                            A6 = Mid(A4, 7, A5 - 4)
                        End If
                        RaiseEvent incomingIM(CStr(A2), CStr(A6))
                    ElseIf Mid(strData, 26, 1) = Chr(2) Then 'attachment
                    
                    End If
            
                Case 4867 'some response from some request packet
            
                    sendPacket 1, requestPacket5
                    
                    If LoggedOn = False Then
                        setProfile "oscSock version 1.1 beta"
                        DoEvents
                        
                        sendPacket 1, addICBMParam
                        
                        If Ghosted = True Then
                            sendPacket 1, clientReadyJacked 'ghosted
                        Else
                            sendPacket 1, clientReady 'client ready
                        End If
                        
                        sendPacket 1, watcherPacket1
                        sendPacket 1, watcherPacket2
                        sendPacket 1, watcherPacket3(gData(0))
                        
                        Timer1.Enabled = True
                        RaiseEvent loggedIn(CStr(gData(2)), CStr(gData(4)))
                        LoggedOn = True
                    End If

                    
            End Select
            
        Case Else 'add-on services
        
            'give the server your auth. cookie
            If Asc(Mid(strData, 2, 1)) = 1 Then
                sendPacket2 1, intIndex, authLogin2(CStr(gData(9)))
            End If
        
            'generate a unique long value from the SNAC header
            Select Case GetLength(Chr(0) & Mid(strData, 8, 1) & Mid(strData, 10, 1))
            
                Case 259 'step 1, request rate information
                    
                    sendPacket2 2, intIndex, formatPacket1
                    sendPacket2 5, intIndex, ""
                    
                Case 263 'step 3, handshaking/format
                
                    If Label3(intIndex) = 0 Then
                        Label3(intIndex) = 1
                        sendPacket2 2, intIndex, rateAck
                        DoEvents
                        
                        Select Case Label2(intIndex)
                        
                            Case "format"
                            
                                sendPacket2 2, intIndex, formatReady
                                sendPacket2 2, intIndex, nameFormat(gData(2))
                                
                            Case "email"
                            
                                sendPacket2 2, intIndex, emailReady
                                sendPacket2 2, intIndex, emailChange(gData(8))
                        
                            Case "chat"
                                
                                sendPacket2 2, intIndex, chatReady
                                sendPacket2 2, intIndex, chatCreate(gData(5))
                                
                            Case "chat2"
                            
                                sendPacket2 2, intIndex, chatReady2
                                RaiseEvent chatReady(CStr(gData(5)), CInt(intIndex))
                            
                        End Select
                        
                    End If
                    
                Case 280 'step 2, some more handshake bullshit
                    
                    sendPacket2 2, intIndex, requestRate
                    
                Case 3337 'chat room information
                
                    If Label2(intIndex) = "chat" Then
                        A1 = Asc(Mid(strData, 23, 1))
                        A2 = Mid(strData, 24, A1)
                        pData 8, "chat2"
                        sendPacket 1, servChat2(CStr(A2))
                    End If
                    
                Case 3587 'buddy entering chat room
                
                    If Label2(intIndex) = "chat2" Then
                        A1 = Asc(Mid(strData, 17, 1))
                        A2 = Mid(strData, 18, A1)
                        RaiseEvent chatBuddyEntered(CStr(A2))
                    End If
                    
                Case 3588 'buddy leaving chat room
                
                    If Label2(intIndex) = "chat2" Then
                        A1 = Asc(Mid(strData, 17, 1))
                        A2 = Mid(strData, 18, A1)
                        RaiseEvent chatBuddyLeft(CStr(A2))
                    End If
                
                Case 3590 'incoming chat send
                
                    If Label2(intIndex) = "chat2" Then
                        A1 = Asc(Mid(strData, 31, 1))
                        A2 = Mid(strData, 32, A1)
                        A3 = GetLength(Chr(0) & Mid(strData, 86 + A1, 2))
                        A4 = Mid(strData, 88 + A1, A3)
                        RaiseEvent incomingChat(CInt(intIndex), CStr(A2), CStr(A4))
                    End If
                    
            End Select

    End Select
    
DoEvents
    
End Sub

'easier than variables IMO
Private Function gData(intIndex)
    gData = SHAFT(intIndex)
End Function

Private Sub pData(intIndex, TheData)
    SHAFT(intIndex) = TheData
End Sub

Private Function getTLV(strData, intType As Integer) As String
    Dim strStartIt As String
    Dim i
    
    If InStr(1, strData, IntegerToBase256(intType)) <> 0 Then
        For i = 1 To Len(strData)
            If Mid(strData, i, 2) = IntegerToBase256(intType) Then
                strStartIt = Mid(strData, i)
                getTLV = Mid(strStartIt, 5, GetLength(Chr(0) & Mid(strStartIt, 3, 2)))
                Exit Function
            End If
        Next i
    Else
        getTLV = ""
    End If
End Function

