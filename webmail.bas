Attribute VB_Name = "webmail"
'This code is used for the Web Mail. The interface is based VERY LOOSELY off of that
'in Ashley's Mailserver.

Private Const htmlopen  As String = _
    "Content-type: text/html" & vbCrLf & vbCrLf & _
    "<?xml version=""1.0"" encoding=""iso-8859-1""?>" & _
    "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"">" & _
    "<html xmlns=""http://www.w3.org/1999/xhtml"">" & _
        "<head>" & _
            "<title>"

Private Const headclose As String = "</title>" & _
            "<style type=""text/css""><!--" & _
                "body{" & _
                    "color: lime;" & _
                    "background-color: black;" & _
                    "font-family: Monospace;" & _
                "}" & _
                "a{" & _
                    "text-decoration: none;" & _
                    "color: white;" & _
                "}" & _
                "a:hover{" & _
                    "border-top: thin solid lime;" & _
                    "border-bottom: thin solid lime;" & _
                "}" & _
                "td.hoveron:hover{" & _
                    "color: white;" & _
                "}" & _
            "--></style>" & _
        "</head>"

Private Const bodystart = headclose & _
        "<body>" & _
            "<table border=""0"" width=""100%"" summary=""you need tables to view this page"">" & _
                "<tr>" & _
                    "<td width=""10%"" style=""border:thin solid lime; text-align:center"">" & _
                        "<a href=""/mail/inbox.webmail"">Inbox</a><br />" & _
                        "<a href=""/mail/compose.webmail"">Compose</a><br />" & _
                        "<a href=""/mail/settings.webmail"">Settings</a><br />" & _
                        "<a href=""/mail/logout.webmail"">Login/Logout</a><br />" & _
                        "<a href=""/mail/signup.webmail"">Signup</a>" & _
                    "</td>" & _
                    "<td width=""1%"">&nbsp;</td>" & _
                    "<td valign=""top"">"
                    
Private Const bodyend = _
                    "</td>" & _
                "</tr>" & _
            "</table>" & _
        "</body>" & _
    "</html>"
                    
Public Function DoWebSite(FileName, vars As Dictionary, headers As Dictionary, ByRef pageheader) As String
    If UserPassword(vars("un")) = vars("pw") And Len(vars("pw")) > 0 Then
        pageheader = pageheader & "Set-Cookie: un=" & vars("un") & "; expires=Fri 28-Jun-2012 13:25:03 GMT;  path=/; domain=" & hostname & ";" & vbCrLf
        pageheader = pageheader & "Set-Cookie: pw=" & vars("pw") & "; expires=Fri 28-Jun-2012 13:25:03 GMT;  path=/; domain=" & hostname & ";" & vbCrLf
    Else
        If Not (FileName = "login.webmail" Or FileName = "signup.webmail") Then
            pageheader = "HTTP/1.0 302 FOUND" & vbNewLine & "Server: NerdMail POP3/SMTP Server" & vbNewLine & "Host: " & _
            hostname & vbNewLine & "Url: /mail/inbox.webmail" & vbNewLine & "Location: /mail/inbox.webmail" & vbNewLine & "Connection: close" & vbNewLine & _
            vbNewLine
        End If
    End If
    
    If FileName = "login.webmail" Then
        DoWebSite = LoginPage
        Exit Function
    End If
    
    If FileName = "inbox.webmail" Then
        If vars("Delete") = "Delete" Then
            For Each f In vars.Keys
                If vars(f) = "dm" Then
                    On Error Resume Next
                    modMySQL.Execute "DELETE FROM email WHERE msgID='" & f & "'"
                    On Error GoTo 0
                End If
            Next
        End If
        
        accmsgcount = modMySQL.MessageCount(vars("un"))
        
        DoWebSite = htmlopen & "Inbox for " & vars("un") & "@" & frmMain.Domain & bodystart & "There " & IIf(accmsgcount = 1, "is ", "are ") & accmsgcount & _
        " message" & IIf(accmsgcount = 1, " ", "s ") & "in your inbox."
        
        DoWebSite = DoWebSite & "<p><form action=""inbox.webmail"" method=""post""><table border=""1"" summary=""you still need tables to view this page""><tr><td align=""center"">&nbsp;</td><td align=""center"" width=""50"">From:</td><td align=""center"" width=""300"">Subject:</td></tr>"
        Dim msgs As MYSQL_RS, msgID As String
        Dim from As String, subject As String
        Set msgs = modMySQL.GetMessages(vars("un"))
        
        msgs.MoveFirst
        Do Until msgs.EOF
            msgID = msgs.Fields("msgID").value
            from = getmailheader(msgs.Fields("message").value, "from")
            subject = getmailheader(msgs.Fields("message").value, "subject")
            DoWebSite = DoWebSite & "<tr class=""hoveron""><td><input type=""checkbox"" name=""" & msgs.Fields("msgID").value & """ value=""dm"" /><td>" & from & "</td><td><a href=""showmsg.webmail?msg=" & msgID & """>" & subject & "</a></td></tr>"
            msgs.MoveNext
        Loop
        DoWebSite = DoWebSite & "</table><input type=""submit"" name=""Delete"" value=""Delete"" /></form>" & bodyend
    End If
    
    If FileName = "showmsg.webmail" Then
        DoWebSite = htmlopen & getmailheader(GetMessage(vars("msg")), "subject") & "-" & getmailheader(GetMessage(vars("msg")), "from") & bodystart
        Source = GetMessage(vars("msg"))
        mh = Left(Source, InStr(1, Source, vbCrLf & vbCrLf) - 1)
        body = Mid(Source, InStr(1, Source, vbCrLf & vbCrLf) + 3)
        body = Left(body, Len(body) - 3)
        
        Set hlist = parseheaders(CStr(mh))
        DoWebSite = DoWebSite & "<table summary=""you still need tables"" border=""1"" bordercolor=""#000000"" width=""550""><tr><td>"
        For Each a In Array("From", "To", "Subject", "Date")
            DoWebSite = DoWebSite & "<span style=""font-weight: bold;"">" & a & ":</span> " & Replace(Replace(hlist(a), ">,<", ">, <"), "<", "&lt;") & "<br />"
        Next
        DoWebSite = DoWebSite & "<p>" & Replace(Replace(body, vbCrLf & vbCrLf, "</p><p>"), vbCrLf, "<br />")
        DoWebSite = DoWebSite & "</table><p> &nbsp; </p><p> &nbsp; </p><p><span style=""font-weight: bold"">Message source follows:</span></p><p><pre>" & Replace(Source, "<", "&lt;") & "</pre></p>" & bodyend
    End If
    
    If FileName = "logout.webmail" Then
        pageheader = pageheader & "Set-Cookie: un=" & vars("un") & "; expires=Thu 28-Jun-2000 13:25:03 GMT;  path=/; domain=" & hostname & ";" & vbCrLf
        pageheader = pageheader & "Set-Cookie: pw=" & vars("pw") & "; expires=Thu 28-Jun-2000 13:25:03 GMT;  path=/; domain=" & hostname & ";" & vbCrLf
        DoWebSite = htmlopen & "You have been logged out" & bodystart & "<a href=""/"">Sign In</a>" & bodyend
    End If
    
    If FileName = "compose.webmail" Then
        DoWebSite = htmlopen & "Compose a new email" & bodystart & "<p style=""font-weight: bold"">Compose a new email</p><p><form action=""send.webmail"" method=""post"">" & _
        "<div align=""center""><table summary=""yep.. still""><tr><td align=""right""><span style=""font-weight: bold"">To:</span></td><td><input size=""60"" name=""to"" value=""" & vars("to") & """ /></td></tr>" & _
        "<tr><td align=""right""><span style=""font-weight: bold"">From:</span></td><td><input size=""60"" name=""from"" value=""" & vars("un") & "@" & frmMain.Domain & """ /></td></tr>" & _
        "<tr><td align=""right""><span style=""font-weight: bold"">Subject:</span></td><td><input size=""60"" name=""subject"" value=""" & vars("subject") & """ /></td></tr>" & _
        "</table><br />"
        
        DoWebSite = DoWebSite & "<textarea rows=""16"" cols=""70"" name=""body""></textarea><p><input type=""submit"" value=""Send"" name=""Send"" /></div>"
        DoWebSite = DoWebSite & bodyend
    End If
    
    If FileName = "send.webmail" Then
        from = vars("from")
        ato = vars("to")
        subject = vars("subject")
        body = vars("body")
        
        Data = "Recieved: " & frmMain.Domain & " webmail, user=" & vars("un") & vbCrLf & _
        "From: " & from & vbCrLf & _
        "To: " & ato & vbCrLf & _
        "Date: " & Now & vbCrLf & _
        "Subject: " & subject & vbCrLf & vbCrLf & body & vbCrLf & "."
        
        For Each aato In Split(ato, ",")
            modMySQL.NewQueueMail modMySQL.GenMsgID, from, CStr(aato), CStr(Data)
        Next
        DoWebSite = htmlopen & "Mail has been sent" & bodystart & "<p style=""font-weight: bold"">Your message has been sent!</p>" & bodyend
    End If
    
    If FileName = "signup.webmail" Then
        If vars("Signup") = "Signup" Then
            If Len(vars("pw1")) < 3 Then
                er = "Password is too short!"
                GoTo nope
            ElseIf Len(vars("unp")) < 3 Then
                er = "Username is too short!"
                GoTo nope
            ElseIf UserExists(vars("unp")) Then
                er = "Account already exists!"
                GoTo nope
            ElseIf vars("pw1") <> vars("pw2") Then
                er = "Passwords do not match!"
                GoTo nope
            End If
            
            'Go ahead and create account
            NewUser vars("unp"), vars("pw1")
            
            modMySQL.NewWelcomeMessage vars("unp"), vars("pw1")
            DoWebSite = htmlopen & "Signup successful!" & bodystart & "<p>Signup successful, <a href=""login.webmail"">login</a> to continue.</p>" & bodyend
        Else
nope:
            DoWebSite = htmlopen & "Signup for " & frmMain.Domain & " email" & bodystart & _
            "<p>Signup for the " & frmMain.Domain & " email service and get:</p><p><ul>"
            
            For Each x In Array("Full POP3 access", "Webmail access", "Autoresponder")
                DoWebSite = DoWebSite & "<li>" & x & "</li>"
            Next
            
            DoWebSite = DoWebSite & "</ul></p><p>To get all that, simply sign up below:</p><p>" & _
            "<form action=""signup.webmail"" method=""post"">" & _
            "<span style=""color: red; font-size: larger; font-weight: bold"">" & er & "</span>" & _
            "<table summary=""get a new browser already"">" & _
            "<tr><td>Username:</td><td><input name=""unp"" /></td></tr>" & _
            "<tr><td>Password:</td><td><input type=""password"" name=""pw1"" /></td></tr>" & _
            "<tr><td>Verify Password:</td><td><input type=""password"" name=""pw2"" /></td></tr>" & _
            "</table><input type=""submit"" name=""Signup"" value=""Signup"" /></form>" & bodyend
        End If
    End If
    
    If FileName = "settings.webmail" Then
        Dim sql As String
        DoWebSite = htmlopen & "Modify your account" & bodystart & "<p>Modify your account</p><p>"
        
        If vars("Save") = "Save" Then
            If vars("pw1") <> vars("pw") Then
                DoWebSite = DoWebSite & "<span style=""color: red; size: large; font-weight: bold"">Enter your current password!</span>"
                GoTo no
            ElseIf vars("pw2") <> vars("pw3") Then
                DoWebSite = DoWebSite & "<span style=""color: red; size: large; font-weight: bold"">Passwords do not match!</span>"
                GoTo no
            ElseIf Len(vars("pw2")) < 3 Then
                DoWebSite = DoWebSite & "<span style=""color: red; size: large; font-weight: bold"">Password is too short!</span>"
                GoTo no
            End If
                
            sql = "UPDATE users SET password='" & vars("pw2") & "'"
            modMySQL.Execute sql
            DoWebSite = DoWebSite & "<span style=""color: red; size: large; font-weight: bold"">Password changed, please log back in</span>"
        Else
no:
            DoWebSite = DoWebSite & "<form action=""settings.webmail"" method=""post""><table summary=""i give up..."">"
            DoWebSite = DoWebSite & "<tr><td>Old Password:</td><td><input type=""password"" name=""pw1"" /></td></tr>"
            DoWebSite = DoWebSite & "<tr><td>New Password:</td><td><input type=""password"" name=""pw2"" /></td></tr>"
            DoWebSite = DoWebSite & "<tr><td>Confirm New Password:</td><td><input type=""password"" name=""pw3"" /></td></tr>"
            DoWebSite = DoWebSite & "</table><input type=""submit"" name=""Save"" value=""Save"" /></form>" & bodyend
        End If
    End If
End Function

Function LoginPage() As String
    Dim t As String, m As String, b As String
    
    t = htmlopen & "&lt;Check Email for " & frmMain.Domain & "&gt;" & bodystart
    
    m = "<p>" & _
        "Any computer. Any internet connection. One email address. Introducing NerdMail, " & _
        "the ONLY open source MySQL-based POP3/SMTP Server*. NerdMail is coded entirely " & _
        "in Visual Basic, utilizing the MyVbQl library to access the powerful MySQL " & _
        "database. Completely free for anyone, and completely open source." & _
        "</p>"
        
    m = m & _
        "<p><form action=""/mail/inbox.webmail"" method=""post"">" & _
        "Username: <input type=""text"" name=""un"" /><br />" & _
        "Password: <input type=""password"" name=""pw"" /><br />" & _
        "<input type=""submit"" value=""Login"" />" & _
        "</form>"
        
    b = bodyend
    
    LoginPage = t & m & b
End Function

Public Function getmailheader(Content, headername As String) As String
    Dim b As Dictionary
    Set b = parseheaders(CStr(Mid(Content, 1, InStr(1, Content, vbCrLf & vbCrLf) - 1)))
    getmailheader = b(headername)
End Function

'These functions were taken from Ashley's Mailserver, which in turn were taken from
'his web server (called "Be Your Own Geocities" on PSCode). Check it out.

Public Function tophpvariables(getdata As String, postdata As String, cookiedata As String) As Dictionary
    'takes the get data (?message=1), the post data (Feedback=Hi%2E+Im+Ashley etc), and cookies (session=1414)
    'and converts it all to one nice data dictonary
    Dim back As New Dictionary
    getdata = Replace(cookiedata, ";", "&") & "&" & getdata & "&" & postdata
    getdata = Replace(getdata, " ", "")
    getdata = Replace(getdata, vbCrLf, "")
    
    Keys = Split(getdata, "&")
    For a = 0 To UBound(Keys)
        If InStr(1, Keys(a), "=") = 0 Then Keys(a) = Keys(a) & "="
        K = Mid(Keys(a), 1, InStr(1, Keys(a), "=") - 1)
        v = Mid(Keys(a), InStr(1, Keys(a), "=") + 1)
        K = fromhttpstringtostring(CStr(K))
        v = fromhttpstringtostring(CStr(v))
        If K <> "" Then back(K) = v
    Next a
    Set tophpvariables = back
End Function

Public Function fromhttpstringtostring(httpstring As String) As String
    'turns 'This%20is%20cool' into 'This is cool'
    httpstring = Replace(httpstring, "+", " ")
    While InStr(1, httpstring, "%")
        fromhttpstringtostring = fromhttpstringtostring & Mid(httpstring, 1, InStr(1, httpstring, "%") - 1)
        httpstring = Mid(httpstring, InStr(1, httpstring, "%"))
        esc = Mid(httpstring, 1, 3)
        ch = Chr(hexdiget(Mid(esc, 2, 1)) * 16 + hexdiget(Mid(esc, 3, 1)))
        httpstring = Replace(httpstring, esc, ch)
    Wend
    fromhttpstringtostring = fromhttpstringtostring & httpstring
End Function

Public Function hexdiget(d) As Integer
    'converts a number from 0-15 into a hexeqiverlant (ie a=10)
    If d = Val(CStr(d)) Then hexdiget = d: Exit Function
    Select Case LCase(d)
    Case "a"
        hexdiget = 10
    Case "b"
        hexdiget = 11
    Case "c"
        hexdiget = 12
    Case "d"
        hexdiget = 13
    Case "e"
        hexdiget = 14
    Case "f"
        hexdiget = 15
    End Select
End Function

Public Function parseheaders(h As String) As Dictionary
    'turn
    'cookie: name=ashley
    'referer: www.pornrus.com
    'accept: all the stuff that goes here.
    'langauge: en-au
    'etc.
    'into a datadictonary.
    
    'would also be really usefull for email header parsing, but, I
    'only use it for http request parsing
    Dim K As String, v As String
    Set p = New Dictionary
    p.CompareMode = TextCompare
    h = h & vbNewLine
    h = Replace(h, ": ", ":")
    h = Replace(h, vbCrLf & " ", "")
    h = Replace(h, vbCrLf & vbTab, "")
    While h <> vbNewLine And h <> ""
        K = LCase(Mid(h, 1, InStr(1, h, ":") - 1))
        h = Mid(h, Len(K) + 2)
        v = Mid(h, 1, InStr(1, h, vbNewLine) - 1)

        h = Mid(h, Len(v) + 3)
        p(K) = v
    Wend
    Set parseheaders = p
End Function

Function hostname() As String
    hostname = frmMain.Domain ' & ":74"
End Function
