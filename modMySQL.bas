Attribute VB_Name = "modMySQL"
Private host     As String
Private name     As String
Private pass     As String
Private base     As String

Private mysql    As New MYSQL_CONNECTION

Public queue     As MYSQL_RS
Public users     As MYSQL_RS

Sub InitMySQL()
    Dim connect As MYSQL_CONNECTION_STATE
    
    host = GetSetting(App.Title, "Database Settings", "Hostname", "localhost")
    name = GetSetting(App.Title, "Database Settings", "Username", "nerdmail")
    pass = GetSetting(App.Title, "Database Settings", "Password", "nerd1")
    base = GetSetting(App.Title, "Database Settings", "Database", "nerdmail")
    
    connect = mysql.OpenConnection(host, name, pass, base)
    
    If connect = MY_CONN_CLOSED Then
        MsgBox mysql.Error.Description
        End
    End If
    
    Set queue = mysql.Execute("SELECT * FROM queue")
    Set users = mysql.Execute("SELECT * FROM users")
End Sub

Sub Update()
    Set queue = mysql.Execute("SELECT * FROM queue")
    Set users = mysql.Execute("SELECT * FROM users")
End Sub

Function UserExists(username) As Boolean
    Dim ue As MYSQL_RS
    Set ue = mysql.Execute("SELECT * FROM users WHERE username='" & username & "'")
    If ue.AffectedRecords > 0 Then UserExists = True
End Function

Function UserPassword(username) As String
    If Not UserExists(username) Then Exit Function
    
    Dim up As MYSQL_RS
    Set up = mysql.Execute("SELECT * FROM users WHERE username='" & username & "'")
    up.MoveFirst
    UserPassword = up.Fields("password").value
End Function

Sub NewLocalMail(msgID As String, sender As String, recipient As String, message As String)
    Dim sql As String
    sql = "INSERT INTO email VALUES('"
    sql = sql & msgID & "', '"
    sql = sql & sender & "', '"
    sql = sql & LCase(recipient) & "', '"
    sql = sql & message & "')"
    
    mysql.Execute sql
End Sub

Sub RemoveFromQueue(msgID As String)
    Dim sql As String
    sql = "DELETE FROM queue WHERE msgID='" & msgID & "'"
    mysql.Execute sql
End Sub

Sub NewQueueMail(msgID As String, sender As String, recipient As String, message As String)
    Dim sql As String
    sql = "INSERT INTO queue VALUES('"
    sql = sql & msgID & "', '"
    sql = sql & sender & "', '"
    sql = sql & recipient & "', '"
    sql = sql & message & "')"
    
    mysql.Execute sql
End Sub

Function mysqlVer() As String
    mysqlVer = mysql.Execute("SELECT Version()").Fields(0).value
End Function

Sub SaveSettings()
    SaveSetting App.Title, "Database Settings", "Hostname", host
    SaveSetting App.Title, "Database Settings", "Username", name
    SaveSetting App.Title, "Database Settings", "Password", pass
    SaveSetting App.Title, "Database Settings", "Database", base
End Sub

Sub MailerDaemon(sender As String, recipient As String, message As String)
    Dim message2 As String
    
    message2 = "To: " & sender & vbCrLf
    message2 = message2 & "From: mailerdaemon@" & frmMain.Domain & vbCrLf
    message2 = message2 & "Subject: Bounced Email" & vbCrLf & vbCrLf
    message2 = message2 & "Your email to " & recipient & " could not be delivered."
    message2 = message2 & vbCrLf & vbCrLf
    message2 = message2 & "The original message follows:" & vbCrLf
    message2 = message2 & message & vbCrLf & vbCrLf
    message2 = message2 & "We apologize for the inconvenience." & vbCrLf
    message2 = message2 & frmMain.Domain & " Mailer Daemon"
    
    NewQueueMail GenMsgID, "mailerdaemon@" & frmMain.Domain, sender, message2
End Sub

Function GenMsgID() As String
    Dim sHour    As String
    Dim sMinute  As String
    Dim sSecond  As String
    Dim sNow     As String
    
    sNow = Time$
    
    sHour = Hour(sNow)
    sMinute = Minute(sNow)
    sSecond = Second(sNow)
    
    If Len(sHour) = 1 Then sHour = "0" & sHour
    If Len(sMinute) = 1 Then sHour = "0" & sMinute
    If Len(sSecond) = 1 Then sHour = "0" & sSecond
    
    GenMsgID = sHour & sMinute & sSecond
    Randomize Timer
    
    Dim RandStr As String
    RandStr = Trim(Str(Int(Rnd * 9999)))
    Do Until Len(RandStr) = 4
        RandStr = "0" & RandStr
    Loop
    
    GenMsgID = GenMsgID & RandStr
End Function

Function Execute(sql As String) As MYSQL_RS
    Set Execute = mysql.Execute(sql)
End Function

Function GetMessage(msgID) As String
    Dim gm As MYSQL_RS
    Set gm = mysql.Execute("SELECT * FROM email WHERE msgID='" & msgID & "'")
    
    gm.MoveFirst
    GetMessage = gm.Fields("message").value
End Function

Function GetMessages(username As String) As MYSQL_RS
    Set GetMessages = mysql.Execute("SELECT * FROM email WHERE lcase(recipient)='" & LCase(username) & "@" & (frmMain.Domain) & "'")
End Function

Sub NewUser(username As String, password As String)
    mysql.Execute "INSERT INTO users VALUES('" & username & "', '" & password & "')"
    Update
End Sub

Function InputBox2(Caption As String, Prompt As String, Optional Default As String = "", Optional PasswordChar As String = "") As String
    Dim ib As New frmInputBox
    ib.Caption = Caption
    ib.lblPrompt.Caption = Prompt
    ib.txtInputBox.Text = Default
    ib.txtInputBox.PasswordChar = PasswordChar
    If ib.txtInputBox.Text <> "" Then
        ib.txtInputBox.SelStart = 0
        ib.txtInputBox.SelLength = Len(ib.txtInputBox.Text)
    End If
    
    ib.Show vbModal
    If ib.Canceled Then Exit Function
    InputBox2 = ib.txtInputBox.Text
End Function

Function MessageCount(username) As Long
    MessageCount = GetMessages(CStr(username)).AffectedRecords
End Function

Sub NewWelcomeMessage(username, password)
    Dim message As String
    
    message = "To: " & username & "@" & frmMain.Domain & vbCrLf & _
              "From: welcome@" & frmMain.Domain & vbCrLf & _
              "Subject: Welcome to " & frmMain.Domain & " Email!" & vbCrLf & _
              "Date: " & Now & vbCrLf & vbCrLf & _
              "Welcome to the " & frmMain.Domain & " email service!" & vbCrLf & vbCrLf & _
              "Username: " & username & vbCrLf & _
              "Password: " & password & vbCrLf & _
              "POP3 Server: " & frmMain.Domain & vbCrLf & _
              "SMTP Server: " & frmMain.Domain & vbCrLf & _
              "Also, you can access webmail at http://" & hostname & vbCrLf & _
              "Thank you for signing up with " & frmMain.Domain & " email!" & vbCrLf & "."
    
    NewQueueMail GenMsgID, "welcome@" & frmMain.Domain, username & "@" & frmMain.Domain, message
End Sub
