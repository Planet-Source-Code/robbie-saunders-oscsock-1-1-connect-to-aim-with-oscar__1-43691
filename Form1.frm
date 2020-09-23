VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "OSCAR Bot"
   ClientHeight    =   4710
   ClientLeft      =   3585
   ClientTop       =   3555
   ClientWidth     =   5775
   LinkTopic       =   "Form1"
   ScaleHeight     =   4710
   ScaleWidth      =   5775
   Begin VB.CommandButton Command17 
      Caption         =   "Leave Chat"
      Height          =   375
      Left            =   4920
      TabIndex        =   24
      Top             =   3720
      Width           =   735
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Send Chat"
      Height          =   495
      Left            =   4440
      TabIndex        =   23
      Top             =   4080
      Width           =   1095
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   5040
      TabIndex        =   22
      Text            =   "0"
      Top             =   3360
      Width           =   495
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Join Chat"
      Height          =   495
      Left            =   3120
      TabIndex        =   21
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Store Comment"
      Height          =   375
      Left            =   3000
      TabIndex        =   20
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Sign On Friend"
      Height          =   375
      Left            =   4080
      TabIndex        =   19
      Top             =   360
      Width           =   1335
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Change Email"
      Height          =   375
      Left            =   4080
      TabIndex        =   18
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton Command11 
      Caption         =   "FormatSN"
      Height          =   375
      Left            =   2640
      TabIndex        =   17
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Empty BList"
      Height          =   375
      Left            =   4080
      TabIndex        =   16
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Empty Icon"
      Height          =   375
      Left            =   4080
      TabIndex        =   15
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Expression"
      Height          =   375
      Left            =   4080
      TabIndex        =   14
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Ghosted Login"
      Height          =   255
      Left            =   2640
      TabIndex        =   13
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Get Profile"
      Height          =   375
      Left            =   2400
      TabIndex        =   12
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Height          =   1215
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Text            =   "Form1.frx":0000
      Top             =   3240
      Width           =   2535
   End
   Begin VB.CheckBox Check1 
      Caption         =   "is typing"
      Height          =   375
      Left            =   2760
      TabIndex        =   10
      Top             =   1920
      Width           =   1575
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Modal Attack"
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Not Away"
      Height          =   375
      Left            =   960
      TabIndex        =   9
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Away"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   2640
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "File"
      Height          =   375
      Left            =   960
      TabIndex        =   7
      Top             =   2160
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "IM"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2160
      Width           =   735
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Text            =   "i'm alive!!!"
      Top             =   1800
      Width           =   2295
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Text            =   "robbieshit"
      Top             =   1440
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Login"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   2
      Text            =   "w0rdup1t"
      Top             =   480
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "oscsockbot"
      Top             =   120
      Width           =   2295
   End
   Begin Project1.oscSock oscSock 
      Left            =   3600
      Top             =   3720
      _ExtentX        =   2672
      _ExtentY        =   873
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If Check1.Value = 1 Then
    oscSock.isTyping Text3, True
Else
    oscSock.isTyping Text3, False
End If
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
    oscSock.Ghosted = True
Else
    oscSock.Ghosted = False
End If
End Sub

Private Sub Command1_Click()
oscSock.loginUser Text1, Text2
End Sub

Private Sub Command10_Click()
oscSock.sendPacket 1, COCKBLOCK(String(8, Chr(GRInteger(0, 255))), Text3, ChrA("9 70 19 75 76 127 17 209 130 34 68 69 83 84 0 0"))
End Sub

Private Sub Command11_Click()
oscSock.formatSN Text3
End Sub

Private Sub Command12_Click()
oscSock.changeEmail Text4
End Sub

Private Sub Command13_Click()
oscSock.signUpFriend Text3, Text4
End Sub

Private Sub Command14_Click()
oscSock.storeComment Text3, Text4
End Sub

Private Sub Command15_Click()
oscSock.joinChat Text3
End Sub

Private Sub Command16_Click()
oscSock.sendChat Text6, Text3
End Sub

Private Sub Command17_Click()
oscSock.endAddIn Text6
End Sub

Private Sub Command2_Click()
oscSock.sendIM Text3, Text4
End Sub

Private Sub Command3_Click()
oscSock.sendFile Text3, Text4
End Sub

Private Sub Command4_Click()
oscSock.setAway Text4
End Sub

Private Sub Command5_Click()
oscSock.setAway ""
End Sub

Private Sub Command6_Click()
oscSock.directConnect Text3
oscSock.chatInvite Text3, "4", "modals", "modals hah"
oscSock.gameInvite Text3, "modals", "modals hah"
oscSock.sendFile Text3, "modals hah"
oscSock.talkRequest Text3
oscSock.blistSend Text3, BuddyListForm("modals hah")
End Sub

Private Sub Command7_Click()
oscSock.getProfile Text3
End Sub

Private Sub Command8_Click()
oscSock.sendExpression Text3, Text4, Text4
End Sub

Private Sub Command9_Click()
oscSock.sendPacket 1, COCKBLOCK(String(8, Chr(GRInteger(0, 255))), Text3, ChrA("9 70 19 70 76 127 17 209 130 34 68 69 83 84 0 0"))
End Sub

'blank attachment function like from aim filter
Public Function COCKBLOCK(strRequestID As String, strUserName As String, strCapability As String) As String

    A1 = ChrA("0 4 0 6 0 0 0 10 0 6")
    COCKBLOCK = A1 & strRequestID & ChrA("0 2") & Chr(Len(strUserName)) & strUserName & _
        TLV(5, ChrA("0 0") & strRequestID & strCapability & ChrA("0 10 0 2 0 1 0 3 0 4 24 16 172 135 0 5 0 2 20 70 0 15 0 0")) & ChrA("0 3 0 0")

End Function

Private Sub oscSock_buddySignedOn(strName As String, strCapabilities As String)
Text5 = strName
End Sub

Private Sub oscSock_chatBuddyLeft(strName As String)
Text5 = strName & " left!"
End Sub

Private Sub oscSock_chatReady(strChannel As String, intIndex As Integer)
Text5 = "In " & strChannel & "!!"
Text6 = intIndex
End Sub

Private Sub oscSock_incomingChat(intIndex As Integer, strName As String, strMessage As String)
Text5 = "(" & intIndex & ")" & strName & ";" & strMessage
End Sub

Private Sub oscSock_incomingIM(strName As String, strMessage As String)
Text5 = strName & ";" & strMessage
End Sub

Private Sub oscSock_incomingProfile(strFormattedName As String, lngSecondsOnline As Long, intWarningLevel As Integer, strProfile As String)
Text5 = strFormattedName & ";" & intWarningLevel & ";" & lngSecondsOnline & ";" & strProfile
End Sub

Private Sub oscSock_loggedIn(strFormattedName As String, strEmailAddy As String)
MsgBox strFormattedName & " LOGGED ON YAY!"
End Sub

Function BuddyListForm(strBuddyList)
Dim K1, K2, K3, K4, K5() As String
K5 = Split(strBuddyList, ";")
BuddyListForm = TwoByteLen(K5(0))
BuddyListForm = BuddyListForm & IntegerToBase256(UBound(K5))
For i = 1 To UBound(K5)
    BuddyListForm = BuddyListForm & TwoByteLen(K5(i))
    DoEvents
Next i
End Function
