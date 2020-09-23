Attribute VB_Name = "HTTP2"
'An HTTP Client Module Written By Robbie Saunders

Function Post(thePage, theReferer, theHost, theCookie, theContent)
Post = ""
Post = Post & "POST " & thePage & " HTTP/1.1" & vbCrLf
Post = Post & "Accept: image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, application/msword, */*" & vbCrLf
If theReferer <> "" Then Post = Post & "Referer: " & theReferer & vbCrLf
Post = Post & "Accept-Language: ie-ee,en-us;q=0.5" & vbCrLf
Post = Post & "Content-Type: application/x-www-form-urlencoded" & vbCrLf
Post = Post & "Accept-Encoding: gzip, deflate" & vbCrLf
Post = Post & "User-Agent: Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1)" & vbCrLf
If theHost <> "" Then Post = Post & "Host: " & theHost & vbCrLf
Post = Post & "Content-Length: " & Len(theContent) & vbCrLf
Post = Post & "Cache-Control: no-cache" & vbCrLf
If theCookie <> "" Then Post = Post & "Cookie: " & theCookie & vbCrLf
Post = Post & "" & vbCrLf
Post = Post & theContent & vbCrLf
End Function

Function Post2(thePage, theHost, theCookie, theContent)
Post2 = ""
Post2 = Post2 & "POST " & thePage & " HTTP/1.0" & vbCrLf
Post2 = Post2 & "Accept-Language: en" & vbCrLf
Post2 = Post2 & "ETag: " & Chr(34) & "772d65e-65c-ae5e664c6" & Chr(34) & vbCrLf
Post2 = Post2 & "Content-Type: application/x-www-form-urlencoded" & vbCrLf
Post2 = Post2 & "Authorization: Basic Z2FtZTo5391545685994" & vbCrLf
Post2 = Post2 & "Accept: text/html, image/gif, image/jpeg, *; q=.2, */*; q=.2" & vbCrLf
Post2 = Post2 & "User-Agent: Mozilla/4.0 (compatible; MSIE 6.0; Win32)" & vbCrLf
If theHost <> "" Then Post2 = Post2 & "Host: " & theHost & vbCrLf
Post2 = Post2 & "Content-Length: " & Len(theContent) & vbCrLf
Post2 = Post2 & "Connection: Keep -Alive" & vbCrLf
Post2 = Post2 & "Cache-Control: no-cache" & vbCrLf
If theCookie <> "" Then Post2 = Post2 & "Cookie: " & theCookie & vbCrLf
Post2 = Post2 & "" & vbCrLf
Post2 = Post2 & theContent & vbCrLf
End Function

Function Gett(thePage, theReferer, theHost, theCookie)
Gett = ""
Gett = Gett & "GET " & thePage & " HTTP/1.0" & vbCrLf
Gett = Gett & "Accept: image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, application/msword, application/vnd.ms-powerpoint, application/vnd.ms-excel, */*" & vbCrLf
If theReferer <> "" Then Gett = Gett & "Referer: " & theReferer & vbCrLf
Gett = Gett & "Accept -Language: en -us" & vbCrLf
Gett = Gett & "Accept -Encoding: gzip , deflate" & vbCrLf
Gett = Gett & "User-Agent: Mozilla/4.0 (compatible; MSIE 5.01; Windows NT; MSOCD; AtHome020)" & vbCrLf
If theHost <> "" Then Gett = Gett & "Host: " & theHost & vbCrLf
Gett = Gett & "Proxy -Connection: Keep -Alive" & vbCrLf
Gett = Gett & "Pragma: no -cache" & vbCrLf
If theCookie <> "" Then Gett = Gett & "Cookie: " & theCookie & vbCrLf
Gett = Gett & "" & vbCrLf
End Function

Function AIMGet(thePage, theReferer, theHost)
AIMGet = ""
AIMGet = AIMGet & "GET " & thePage & " HTTP/1.0" & vbCrLf
If theHost <> "" Then AIMGet = AIMGet & "Host: " & theHost & vbCrLf
AIMGet = AIMGet & "If-Modified-Since:" & vbCrLf
AIMGet = AIMGet & "Accept:*/*" & vbCrLf
If theReferer <> "" Then AIMGet = AIMGet & "Referer: " & theReferer & vbCrLf
AIMGet = AIMGet & "User-Agent: AIM/30 (Mozilla 1.24b; Windows; I; 32-bit)" & vbCrLf
AIMGet = AIMGet & vbCrLf
End Function

Function XAIMGet(thePage, theReferer, theHost, theCookie, theSN)
XAIMGet = ""
XAIMGet = XAIMGet & "GET " & thePage & " HTTP/1.0" & vbCrLf
XAIMGet = XAIMGet & "Accept: image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, application/msword, application/vnd.ms-powerpoint, application/vnd.ms-excel, */*" & vbCrLf
If theReferer <> "" Then XAIMGet = XAIMGet & "Referer: " & theReferer & vbCrLf
XAIMGet = XAIMGet & "Accept -Language: en -us" & vbCrLf
XAIMGet = XAIMGet & "Accept -Encoding: gzip , deflate" & vbCrLf
XAIMGet = XAIMGet & "User-Agent: Mozilla/4.0 (compatible; MSIE 5.01; Windows NT; MSOCD; AtHome020)" & vbCrLf
XAIMGet = XAIMGet & "X-Aim: product=9&platform=1&channel=179&build=2646&SN=" & theSN & "&CC=BHNH&PC=HLLGDICGBF&UTC=1015547980&LT=1015519180" & vbCrLf
If theHost <> "" Then XAIMGet = XAIMGet & "Host: " & theHost & vbCrLf
XAIMGet = XAIMGet & "Proxy -Connection: Keep -Alive" & vbCrLf
XAIMGet = XAIMGet & "Pragma: no -cache" & vbCrLf
If theCookie <> "" Then XAIMGet = XAIMGet & "Cookie: " & theCookie & vbCrLf
XAIMGet = XAIMGet & "" & vbCrLf
End Function

Function GrabCookie(theStuff, theCookie)
Dim A1, A2
A1 = Split(theStuff, vbCrLf)
For i = 0 To UBound(A1)
    If Len(A1(i)) >= 12 + Len(theCookie) Then
        If Left(A1(i), 12 + Len(theCookie)) = "Set-Cookie: " & theCookie Then
            A2 = Split(A1(i), ";")
            GrabCookie = Right(A2(0), Len(A2(0)) - 12)
            Exit Function
        End If
    End If
    DoEvents
Next i
GrabCookie = False
End Function
