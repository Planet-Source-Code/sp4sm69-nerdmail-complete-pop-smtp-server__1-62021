VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NerdMail v1.0"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5895
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   5895
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAbout 
      Caption         =   "< ABOUT >"
      Height          =   375
      Left            =   4440
      TabIndex        =   3
      Top             =   3120
      Width           =   1335
   End
   Begin MSWinsockLib.Winsock http 
      Index           =   0
      Left            =   120
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   81
   End
   Begin VB.CommandButton cmdUsers 
      Caption         =   "< USERS >"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3120
      Width           =   1335
   End
   Begin NerdMail.MX mx 
      Left            =   2640
      Top             =   1440
      _ExtentX        =   714
      _ExtentY        =   450
   End
   Begin MSWinsockLib.Winsock mta 
      Left            =   120
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock pop3 
      Index           =   0
      Left            =   120
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   110
   End
   Begin MSWinsockLib.Winsock smtp 
      Index           =   0
      Left            =   120
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   25
   End
   Begin VB.Timer timMySQL 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   120
      Top             =   2520
   End
   Begin VB.Timer timMTA 
      Interval        =   10000
      Left            =   120
      Top             =   2040
   End
   Begin ComctlLib.StatusBar status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   3615
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Object.Width           =   212
            MinWidth        =   9
            Key             =   "popIP"
            Object.Tag             =   ""
            Object.ToolTipText     =   "The IP Address of the POP Server"
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Object.Width           =   212
            MinWidth        =   9
            Key             =   "smtpIP"
            Object.Tag             =   ""
            Object.ToolTipText     =   "The IP Address of the SMTP Server"
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   4957
            MinWidth        =   9
            Key             =   "domain"
            Object.Tag             =   ""
            Object.ToolTipText     =   "The root domain of the server"
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   4957
            MinWidth        =   9
            Key             =   "mysqlVer"
            Object.Tag             =   ""
            Object.ToolTipText     =   "The MySQL Database server version"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ComctlLib.ListView lvwQueue 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   5106
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   4
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Message ID"
         Object.Width           =   1965
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Sender"
         Object.Width           =   1965
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Recipient"
         Object.Width           =   1965
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Message"
         Object.Width           =   2099
      EndProperty
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This code is written based on the MySQL Mail server written by me
'Some parts of the code may be copied or derived from code orignally by
'Ashley Harris. Permission WAS obtained

'Debug mode was used during development of the MTA to check progress of sending
'To disable debug mode, change this to False:
Private Const DebugON        As Boolean = False

'POP Connection Variables
Private pop_login()          As Boolean          'Is the connection authenticated?
Private pop_username()       As String           'The POP connection username
Private pop_password()       As String           'The POP connection password
Private pop_messages()       As MYSQL_RS         'Messages for the given user
Private pop_query()          As New Collection   'Queries to be execute on close

'SMTP Connection Variables
Private smtp_state()         As Integer          'State of the SMTP transaction
Private smtp_sender()        As String           'SMTP mail sender
Private smtp_mailtos()       As New Collection   'SMTP mail recipients
Private smtp_message()       As String           'SMTP mail message

'MTA Connection Variables
Private mta_state            As Integer          'State of MTA transaction
Private mta_sender           As String           'MTA mail sender
Private mta_recipient        As String           'MTA mail recipient
Private mta_message          As String           'MTA mail message
Private mta_msgID            As String           'MTA mail message ID

'vbAccelerator System Tray form (Ben Baird <psyborg@cyberhighway.com>)
Private WithEvents SysTray   As frmSysTray
Attribute SysTray.VB_VarHelpID = -1

'MD5 Class for APOP
Private md5                  As New md5

'Holds the root domain (username@DOMAIN)
Public Domain                As String

'Checks for completion of the setup wizard, if needed
Public SetupDone             As Boolean

Private Sub cmdAbout_Click()
    Dim ab As New frmAbout
    ab.Show vbModal, Me
End Sub

'Show the user manager
Private Sub cmdUsers_Click()
    frmUsers.Show vbModal, Me
End Sub

'Called on application startup
Private Sub Form_Load()
    'Redim arrays to allow later Redims
    ReDim pop_login(0)
    ReDim pop_username(0)
    ReDim pop_password(0)
    ReDim pop_messages(0)
    ReDim pop_query(0)
    
    ReDim smtp_state(0)
    ReDim smtp_sender(0)
    ReDim smtp_mailtos(0)
    ReDim smtp_message(0)
    
    'Change "/setup" to "--setup", for example
    cmd = LCase(Command)
    cmd = Replace(cmd, "/", "--")
    
    'Run setup wizard on "--setup" (or "/setup")
    If InStr(1, cmd, "--setup") > 0 Then
        frmSetup.Show vbModal
        If SetupDone = False Then Unload Me
        Exit Sub
    End If
    
    'Run setup wizard if missing configuration
    If GetSetting(App.Title, "Server Settings", "POP IP", "NoCfg") = "NoCfg" Then
        frmSetup.Show vbModal
        If SetupDone = False Then Unload Me
        Exit Sub
    End If
    
    'Flag setup as complete if setup was not necessary
    SetupDone = True
    
    'Initialize the MySQL connection
    modMySQL.InitMySQL
    
    'Add the queued emails to the queue ListView in the window
    Dim lItem As ListItem
    queue.MoveFirst
    Do Until queue.EOF
        Set lItem = lvwQueue.ListItems.Add()
        lItem.Text = queue.Fields("msgID").value
        lItem.SubItems(1) = queue.Fields("sender").value
        lItem.SubItems(2) = queue.Fields("recipient").value
        lItem.SubItems(3) = queue.Fields("message").value
        queue.MoveNext
    Loop
    
    'Load settings!!
    Dim popIP   As String
    Dim smtpIP  As String
    
    popIP = GetSetting(App.Title, "Server Settings", "POP IP", pop3(0).LocalIP)
    smtpIP = GetSetting(App.Title, "Server Settings", "SMTP IP", smtp(0).LocalIP)
    Domain = GetSetting(App.Title, "Server Settings", "Domain", pop3(0).LocalHostName)
    
    'Bind and listen
    On Error GoTo PortInUse
    'pop3(0).bind 110, popIP
    pop3(0).listen
    'smtp(0).bind 25, smtpIP
    smtp(0).listen
    http(0).bind 81 'Do not use port 80, so that we don't interfere with mail servers
    http(0).listen
    On Error GoTo 0
    
    'Setup the status bar
    status.Panels("popIP").Text = popIP
    status.Panels("smtpIP").Text = smtpIP
    status.Panels("mysqlVer").Text = mysqlVer
    status.Panels("domain").Text = Domain
    
    'Initialize the System Tray icon
    Set SysTray = New frmSysTray
    Load SysTray
    
    'Hide window unless told otherwise
    If InStr(1, cmd, "--show") = 0 Then
        Me.Hide
    End If
    Exit Sub
    
    'If port 110, port 25, or port 81 is in use, error and exit
PortInUse:
    MsgBox "The NerdMail server requires the following ports:" & vbCrLf & "110: POP" & vbCrLf & "25: SMTP" & vbCrLf & "81:Webmail"
    End
End Sub

'Perform cleanup, and exit
Sub Quit()
    'Save settings!!
    SaveSetting App.Title, "Server Settings", "POP IP", pop3(0).LocalIP
    SaveSetting App.Title, "Server Settings", "SMTP IP", smtp(0).LocalIP
    SaveSetting App.Title, "Server Settings", "Domain", Domain
    
    'Save MySQL Database settings
    modMySQL.SaveSettings
    
    'Hide the system tray icon
    Unload SysTray
    
    'Exit
    End
End Sub

'Called when the form is unloaded (unlikc Form_Unload, called BEFORE unloading)
'This is used to hide the application, leaving it in the system tray
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'If setup is not complete then we're exiting, so end
    If SetupDone = False Then End
    
    'Don't unload the form
    Cancel = 1
    
    'Hide the window
    Me.Hide
End Sub

'Webmail connection
Private Sub http_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    'We may need to load a new socket... if so, we have a boolean for that
    Dim loadnew As Boolean: loadnew = True
    
    'Loop through all sockets to see if there's on available
    For I = 1 To http.UBound
        If http(I).State <> 7 Then 'Yep, there's one free
            'Close the socket
            http(I).Close
            
            'We're not loading a new socket
            loadnew = False
            
            'Exit the loop
            Exit For
        End If
    Next
    
    'If we didn't find an open socket, load a new one
    If loadnew = True Then Load http(I)
    
    'Accept the request
    http(I).accept requestID
End Sub

'Data arrived on webmail.. I won't explain this routine, it's pretty well documented,
'but it's not my place to explain, as its not mine
Private Sub http_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    'This code is based HEAVILY on the webmail code from Ashley Harris's POP3/SMTP
    'server, found on Planet Source Code. Permission was obtained. Any comments are
    'directly from his Mail Server project, in case he was explaining anything
    Dim fso As New FileSystemObject
'-------------------------------------------------------------------------------------'
    
    'All my fans will probibly notice that this entire sub has been copy/pasted from
    'my webserver that ALLMOST won code of the month on PSC (FULL perl support, acceptable
    'ASP support, some PHP support, resume downloads, etc.) Damn hacker taking out
    'the voting charts. This mail server is my second attempt, please vote for it.
    
    If docroot = "\" Then docroot = App.Path
    
    Dim a As String, filedata As String, headers As Dictionary
    If http(Index).State <> 7 Then http(Index).Close: Exit Sub
    http(Index).getdata a
    'Debug.Print a
    
    'ok, the way I've done this is wrong, because, say someone uploads a file via a cgi that's, say 500kb.
    'the browser will split it into packets of (say) 8kb each, and will send them here.
    'which means this function will be called first with the request and the first 7.5kb of the file.
    'then again, with the second chunk of the file, but no headers, which will crash this because theres
    'no headers to parse. I was stuck, and, cause it was 2AM in the morning, I thought of this stupid
    'and unreliable solution. This is a FIXME! (actually, it works rather good, and prob follows the protocol)
    If bytesTotal = 8192 Or InStr(1, a, "multipart/form-data", TextCompare) Then
        http(Index).Tag = http(Index).Tag & a
        stats.List(Index) = "Incomming File..."
        Exit Sub
    Else
        a = http(Index).Tag & a
        http(Index).Tag = ""
    End If
    
    If a = "" Then
        http(Index).Close
        Exit Sub
    End If
    
    otherheaders = Mid(a & vbNewLine, InStr(1, a, vbNewLine) + 2)
    otherheaders = Mid(otherheaders, 1, InStr(1, otherheaders, vbNewLine & vbNewLine) - 1)
    
    Set headers = parseheaders(CStr(otherheaders))
    
    If InStr(1, a, vbNewLine & vbNewLine) > 0 Then postdata = Mid(a, InStr(1, a, vbNewLine & vbNewLine) + 4) Else postdata = ""
    
    'ADDED BY THE NERD COUNCIL
    'Firefox (the official browser of The Nerd Council, www.mozilla.org), send the POST
    'data seperately from the initial request. Thus, we must handle this
    If LCase(Left(a, 4)) = "post" And postdata = "" Then
        http(Index).Tag = a
        Exit Sub
    End If
    
    If CLng(headers("Content-length")) > Len(postdata) Then
        'ok, there are more packets comming, were executing too early
        'ie5 for mac is the cause of this code
        stats.List(Index) = "Awaiting POST data"
        http(Index).Tag = http(Index).Tag & a
        Exit Sub
    End If
    
    If headers("Content-type") = "application/x-www-form-urlencoded" And IsEmpty(headers("Content-length")) Then
        'my mac did this while posting the feedback form. makes no sense, but, it works.
        ' (it splits the request in 2)
        'stats.List(Index) = "Awaiting POST data"
        http(Index).Tag = http(Index).Tag & a
        Exit Sub
    End If
    
    'get the request, and then take the first line of it, then take just the request page
    a = Left(a, InStr(1, a, vbNewLine) - 1)
    a = Mid(a, InStr(1, a, " ") + 1)
    a = Left(a, InStr(1, a, " ") - 1)

    
    While Mid(a, 1, 3) = "/.."
        a = Mid(a, 4)
    Wend
    
    If Right(a, 1) = "/" Then a = a & Default
    
    'seperated the request string into filename and GET data
    If Not CBool(InStr(1, a, "?")) Then
        a = a & "?"
    End If
    cmd = Left(a, InStr(1, a, "?") - 1)
    Data = Mid(a, InStr(1, a, "?") + 1)
    cmd = Replace(cmd, "/", "\")
    cmd = Replace(cmd, "%20", " ")
    
    header = "HTTP/1.0 200 OK" & vbNewLine & "Server: NerdMail POP3/SMTP Server" & vbNewLine & "Host: " & _
    hostname & vbNewLine & "Connection: close" & vbNewLine
    
    Select Case LCase(fso.GetParentFolderName(cmd))
    Case "\img"
        Open fso.BuildPath(fso.BuildPath(App.Path, "graphics"), fso.GetFileName(cmd)) For Binary As #1
        Dim I As String
        I = Space(LOF(1))
        Get #1, , I
        Close #1
        back = "Content-type: image/png" & vbCrLf & vbCrLf & I
    Case "\mail"
        back = DoWebSite(fso.GetFileName(cmd), tophpvariables(CStr(Data), CStr(postdata), headers("cookie")), headers, header)
    Case Else
        header = ""
        back = "HTTP/1.0 302 FOUND" & vbNewLine & "Server: NerdMail POP3/SMTP Server" & vbNewLine & "Host: " & _
        hostname & vbNewLine & "Url: /mail/inbox.webmail" & vbNewLine & "Location: /mail/inbox.webmail" & vbNewLine & "Connection: close" & vbNewLine & _
        vbNewLine
    End Select
    
    back = header & back
    On Error GoTo outofhere
    'While Len(back) > 0
        'http(Index).SendData Mid(back, 1, 10000)
        'back = Mid(back, 10001)
        't = Timer + 3
        'While t > Timer
            'DoEvents
        'Wend
    'Wend
    http(Index).SendData back
outofhere:
    t = Timer + 1
    While t > Timer
        DoEvents
    Wend
    
    On Error Resume Next
    http(Index).Close
End Sub

'MTA data arrved
Private Sub mta_DataArrival(ByVal bytesTotal As Long)
    On Error GoTo ErrorHandle
    
    'Temporary variable
    Dim a As String
    
    'Check for a linefeed (ie end of the line)
    mta.PeekData a, , bytesTotal
    If Right(a, 2) <> vbCrLf Then Exit Sub
    
    'Pull the data
    mta.getdata a, , bytesTotal
    
    If Right(a, 2) <> vbCrLf Then
        a = Left(a, Len(a) - 1)
        a = a & vbCrLf
    End If
    
    If DebugON = True Then MsgBox ":DEBUG:" & vbCrLf & "MTA recieved data" & vbCrLf & a
    
    'Check for an error (non-error numbers are 220, 250, 351, and 354
    Select Case Left(a, 1)
        Case "2", "3"
            mta_state = mta_state + 1
        Case Else 'ERROR.. give mailer daemon
            If DebugON = True Then
                MsgBox ":DEBUG:" & vbCrLf & "Error from remote server. MTA failed" & vbCrLf & Left(a, Len(a) - 2)
            End If
            modMySQL.MailerDaemon mta_sender, mta_recipient, mta_message
            Exit Sub
    End Select
    
    'Respond to the message (we don't really need to know the response...)
    Select Case mta_state
        Case 1 'Send HELO
            If DebugON = True Then
                MsgBox ":DEBUG:" & vbCrLf & "MTA sending..." & vbCrLf & "HELO " & Domain
            End If
            send "HELO " & Domain, mta
        Case 2 'Send MAIL FROM:
            If DebugON = True Then
                MsgBox ":DEBUG:" & vbCrLf & "MTA sending..." & vbCrLf & "MAIL FROM: <" & mta_sender & ">"
            End If
            send "MAIL FROM: <" & mta_sender & ">", mta
        Case 3 'Send RCPT TO:
            If DebugON = True Then
                MsgBox ":DEBUG:" & vbCrLf & "MTA sending..." & vbCrLf & "RCPT TO: <" & mta_recipient & ">"
            End If
            send "RCPT TO: <" & mta_recipient & ">", mta
        Case 4 'Send DATA
            If DebugON = True Then
                MsgBox ":DEBUG:" & vbCrLf & "MTA sending..." & vbCrLf & "DATA"
            End If
            send "DATA", mta
        Case 5 'Send message
            Dim mtam() As String
            mtam = Split(mta_message, vbCrLf)
            For I = 0 To UBound(mtam)
                If DebugON = True Then
                    MsgBox ":DEBUG:" & vbCrLf & "MTA sending (message)..." & vbCrLf & mtam(I)
                End If
                send mtam(I), mta
            Next
        Case 6 'Send QUIT
            If DebugON = True Then
                MsgBox ":DEBUG:" & vbCrLf & vbCrLf & "MTA sending complete" & vbCrLf & "QUIT"
            End If
            send "QUIT", mta
            modMySQL.RemoveFromQueue mta_msgID
    End Select
    Exit Sub
    
ErrorHandle:
    MsgBox "MTA Error Encountered!" & vbCrLf & Err.Description & vbCrLf & Err.Source & vbCrLf & mta_state
End Sub

'POP Connection request
Private Sub pop3_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    'We may need to load a new socket... if so, we have a boolean for that
    Dim loadnew As Boolean: loadnew = True
    
    'Loop through all sockets to see if there's on available
    For I = 1 To pop3.UBound
        If pop3(I).State <> 7 Then 'Yep there's one available
            'Close the socket
            pop3(I).Close
            
            'We're not loading a new socket
            loadnew = False
            
            'Exit loop
            Exit For
        End If
    Next
    
    'Load a new socket if needed
    If loadnew = True Then Load pop3(I)
    
    If loadnew = True Then
        'We need to increase the array's upper bound so we can use it
        ReDim Preserve pop_login(I)
        ReDim Preserve pop_username(I)
        ReDim Preserve pop_password(I)
        ReDim Preserve pop_messages(I)
        ReDim Preserve pop_query(I)
    Else
        'Reset the variables for our use
        pop_login(I) = False
        pop_username(I) = ""
        pop_password(I) = ""
        Set pop_messages(I) = Nothing
        Set pop_query(I) = Nothing: Set pop_query(I) = New Collection
    End If
    
    'Timestamp for APOP
    ts = "<" & Int(Rnd() * 100000000000#) & Int(Rnd() * 100000000000#) & Int(Rnd() * 100000000000#) & Int(Rnd() * 100000000000#) & ">"
    pop3(I).Tag = ts
    
    'Accept connection
    pop3(I).accept requestID: DoEvents
    
    'Send welcome message
    send "+OK Welcome to the NerdMail POP Server", pop3(I)
End Sub

Private Sub smtp_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    'We may need to load a new socket... if so, we have a boolean for that
    Dim loadnew As Boolean: loadnew = True
    
    'Loop through all sockets to see if there's on available
    For I = 1 To smtp.UBound
        If smtp(I).State <> 7 Then 'Yep there's one available
            'Close the socket
            smtp(I).Close
            
            'We're not loading a new one
            loadnew = False
            
            'Exit loop
            Exit For
        End If
    Next
    
    'Load a new one, if needed
    If loadnew = True Then Load smtp(I)
        
    If loadnew = True Then
        'Redim variables for use
        ReDim Preserve smtp_state(I)
        ReDim Preserve smtp_sender(I)
        ReDim Preserve smtp_mailtos(I)
        ReDim Preserve smtp_message(I)
    Else
        'Reset variables
        smtp_state(I) = 0
        smtp_sender(I) = ""
        Set smtp_mailtos(I) = Nothing: Set smtp_mailtos(I) = New Collection
        smtp_message(I) = ""
    End If
    
    'Accept the connection
    smtp(I).accept requestID: DoEvents
    
    'Send welcome message
    send "220 Welcome to the NerdMail SMTP Server", smtp(I)
End Sub

'POP Data has arrived!
Private Sub pop3_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim a As String
    Dim b As String
    Dim c As Variant, out As String
    Dim messages As MYSQL_RS
    
    'Check for linefeed
    pop3(Index).PeekData a, , bytesTotal
    If Right(a, 1) <> vbLf Then Exit Sub
    
    pop3(Index).getdata a, , bytesTotal
    b = UCase(a)
    
    If pop_login(Index) = False Then
        'Still logging in
        If Left(b, 4) = "QUIT" Then 'Client is terminating
            send "+OK So long", pop3(Index)
            pop3(Index).Close
            Exit Sub
        ElseIf Left(b, 4) = "APOP" Then 'APOP is now supported!
            v = Mid(a, InStr(1, a, " ") + 1)
            digest = Mid(v, InStr(1, v, " ") + 1)
            v = Mid(v, 1, InStr(1, v, " ") - 1)
            username = v
            
            If LCase(digest) = LCase(checksum(pop3(Index).Tag & UserPassword(v))) Then
                pop_login(Index) = True
                send "+OK Secure login ok.", pop3(Index)
                Exit Sub
            Else
                send "-ERR Error in secure login!", pop3(Index)
                Exit Sub
            End If
        ElseIf Left(b, 4) = "USER" Then 'Client sent username
            c = Mid(a, 6, Len(a) - 7)
            If c = "" Then 'Blank username
                send "-ERR Blank username", pop3(Index)
                Exit Sub
            Else 'Non-blank username
                modMySQL.Update
                users.MoveFirst
                Do Until users.EOF
                    If LCase(users.Fields("username").value) = LCase(c) Then
                        pop_username(Index) = users.Fields("username").value
                        pop_password(Index) = users.Fields("password").value
                        
                        send "+OK That user is valid", pop3(Index)
                        Exit Sub
                    End If
                    users.MoveNext
                Loop
                send "-ERR Username not found", pop3(Index)
                Exit Sub
            End If
        ElseIf Left(b, 4) = "PASS" Then 'Client sent password
            If pop_username(Index) = "" Then
                send "-ERR Please specify username", pop3(Index)
                Exit Sub
            Else
                c = Mid(a, 6, Len(a) - 7)
                If c = pop_password(Index) Then
                    pop_login(Index) = True
                    Set pop_messages(Index) = modMySQL.GetMessages(pop_username(Index))
                    
                    send "+OK You are now logged in", pop3(Index)
                    Exit Sub
                Else
                    pop_username(Index) = ""
                    pop_password(Index) = ""
                    
                    send "-ERR Password incorrect", pop3(Index)
                    Exit Sub
                End If
            End If
        Else
            send "-ERR Invalid command during authentication", pop3(Index)
            Exit Sub
        End If
    Else
        'Logged in
        Set pop_messages(Index) = modMySQL.GetMessages(pop_username(Index))
        Set messages = pop_messages(Index)
        
        If Left(b, 4) = "NOOP" Then 'POP's version of PING
            send "+OK I ping you back!", pop3(Index)
            Exit Sub
        ElseIf Left(b, 4) = "DELE" Then 'Delete a message
            c = Mid(a, 6, Len(a) - 7)
            messages.MoveFirst
            Do Until messages.AbsolutePosition = c
                messages.MoveNext
            Loop
            
            Dim msgID As String
            msgID = messages.Fields("msgID").value
            
            pop_query(Index).Add "DELETE FROM email WHERE recipient='" & pop_username(Index) & "@" & Domain & "' AND msgID='" & msgID & "'"
            send "+OK Message will be deleted", pop3(Index)
            Exit Sub
        ElseIf Left(b, 4) = "STAT" Then 'Show space usage, and get message count
            If messages.AffectedRecords = 0 Then
                send "+OK 0 0", pop3(Index)
                Exit Sub
            Else
                Dim size As Long
                messages.MoveFirst
                Do Until messages.EOF
                    size = size + Len(messages.Fields("message").value)
                    messages.MoveNext
                Loop
                
                send "+OK " & messages.AffectedRecords & " " & size, pop3(Index)
                Exit Sub
            End If
        ElseIf Left(b, 4) = "RETR" Then 'Retrieve message
            c = Mid(a, 6, Len(a) - 7)
            messages.MoveFirst
            Do Until messages.AbsolutePosition = c
                messages.MoveNext
            Loop
            
            send "+OK " & vbCrLf & messages.Fields("message").value, pop3(Index)
            Exit Sub
        ElseIf Left(b, 4) = "LIST" Then 'Show the size of all messages or, optionally, one message
            c = Mid(Left(a, Len(a) - 2), 6)
            If c = "" Then
                out = "+OK " & vbCrLf
                messages.MoveFirst
                Do Until messages.EOF
                    out = out & messages.AbsolutePosition & " " & Len(messages.Fields("message").value) & vbCrLf
                    messages.MoveNext
                Loop
                out = out & "."
                send out, pop3(Index)
            Else
                messages.MoveFirst
                Do Until messages.AbsolutePosition = c
                    messages.MoveNext
                Loop
                
                send "+OK " & vbCrLf & c & " " & Len(messages.Fields("message").value) & vbCrLf & ".", pop3(Index)
            End If
        ElseIf Left(b, 4) = "TOP " Then 'Get the top v lines of messsage #c
            Dim l() As String
            Dim z As String, x As Integer
            
            c = Mid(a, 5, Len(a) - 6)
            v = Mid(c, InStr(1, c, " ") + 1)
            c = Mid(c, 1, InStr(1, c, " ") - 1)
            
            messages.MoveFirst
            Do Until messages.AbsolutePosition = c
                messages.MoveNext
            Loop
            
            out = "+OK" & vbCrLf
            l = Split(messages.Fields("message").value, vbCrLf)
            
            For x = 1 To UBound(l)
                If l(I) = "" Then Exit For
                out = out & l(I) & vbCrLf
            Next
            
            If v > UBound(l) Then v = UBound(l)
            
            For I = x To v
                out = out & l(x) & vbCrLf
            Next
            
            out = out & "."
            send out, pop3(Index)
        ElseIf Left(b, 4) = "QUIT" Then 'Client is terminating, delete messages
            send "+OK So long", pop3(Index)
            pop3(Index).Close
            For I = 1 To pop_query(Index).Count
                modMySQL.Execute pop_query(Index)(I)
            Next
            Exit Sub
        Else
            send "-ERR Invalid command (" & Left(Right(a, Len(a) - 2), 4) & ")", pop3(Index)
            Exit Sub
        End If
    End If
End Sub

'SMTP data has arrived!
Private Sub smtp_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim a As String
    Dim b As String
        
    smtp(Index).PeekData a, , bytesTotal
    If Right(a, 2) <> vbCrLf Then GoTo ExitSub
    smtp(Index).getdata a, , bytesTotal
    
    b = UCase(Left(a, 4))
    
    Select Case b
        Case "NOOP"
            send "250 I ping you back!", smtp(Index)
            Exit Sub
        Case "RSET"
            smtp_state(Index) = 0
            smtp_sender(Index) = ""
            Set smtp_mailtos(Index) = New Collection
            smtp_message(Index) = ""
            send "250 Ok, everything's reset", smtp(Index)
            Exit Sub
        Case "EHLO"
            smtp_state(Index) = 1
            smtp_sender(Index) = ""
            Set smtp_mailtos(Index) = New Collection
            smtp_message(Index) = ""
            send "250 " & Domain & " It's nice to meet you", smtp(Index)
            Exit Sub
        Case "QUIT"
            If smtp_state(Index) <> 4 Then
                send "221 Ok, goodbye", smtp(Index)
                smtp(Index).Close
                Exit Sub
            End If
        Case "EXPN"
            send "551 Mailing lists are not supported on this server.", smtp(Index)
            Exit Sub
        Case "VRFY"
            c = Mid(a, 6, Len(a) - 2)
            d = Mid(c, InStr(1, c, "@") + 1)
            c = Mid(c, 1, InStr(1, c, "@") - 1)
            If LCase(d) <> LCase(Domain) Then
                If ExtractEmail(a) = "" Then
                    send "551 Invalid syntax!", smtp(Index)
                Else
                    send "252 Syntax ok, but outside domain!", smtp(Index)
                End If
                Exit Sub
            End If
            If UserExists(c) Then
                send "250 Yes, that address exists", smtp(Index)
            Else
                send "551 No, that address does not exist!", smtp(Index)
            End If
            Exit Sub
    End Select
    
    If smtp_state(Index) = 0 Then 'Waiting for HELO
        Select Case b
            Case "HELO"
                smtp(Index).Tag = Mid(a, 6, Len(a) - 7)
                smtp_state(Index) = 1
                send "250 It's a pleasure to meet you", smtp(Index)
                GoTo ExitSub
            Case "HELP"
                send "214 Tell us your IP or Domain (ie 127.0.0.1 or yahoo.com)", smtp(Index)
                GoTo ExitSub
            Case "QUIT"
                send "250 Ok, sorry to see you go", smtp(Index)
                smtp(Index).Close
                GoTo ExitSub
            Case "EHLO"
                send "554 No ESMTP support", smtp(Index)
                GoTo ExitSub
            Case Else
                send "503 It's rude not to introduce yourself", smtp(Index)
                GoTo ExitSub
        End Select
    ElseIf smtp_state(Index) = 1 Then 'Waiting for MAIL FROM
        Select Case b
            Case "MAIL"
                If ExtractEmail(a) <> "" Then
                    smtp_sender(Index) = ExtractEmail(a)
                    smtp_state(Index) = 2
                    send "250 Sender accepted", smtp(Index)
                    GoTo ExitSub
                Else
                    send "501 Expected an email address (ie MAIL FROM: <admin@nerdcouncil.dev>)", smtp(Index)
                    GoTo ExitSub
                End If
            Case "HELP"
                send "214 Enter your email address", smtp(Index)
                GoTo ExitSub
            Case "QUIT"
                send "221 Ok, sorry to see you go", smtp(Index)
                smtp(Index).Close
                GoTo ExitSub
            Case Else
                send "503 Expected an email address (ie MAIL FROM: <admin@nerdcouncil.dev>)", smtp(Index)
                GoTo ExitSub
        End Select
    ElseIf smtp_state(Index) = 2 Then 'Waiting for RCPT TO, then DATA
        Select Case b
            Case "RCPT"
                If ExtractEmail(a) <> "" Then
                    If smtp_mailtos(Index).Count >= 60 Then
                        send "452 Sorry, limit of 60 recipients", smtp(Index)
                        GoTo ExitSub
                    End If
                    
                    If InStr(1, ExtractEmail(a), "@" & Domain) Then
                        Dim un As String
                        un = Mid(ExtractEmail(a), 1, InStr(1, ExtractEmail(a), "@") - 1)
                        If UserExists(un) Then
                            smtp_mailtos(Index).Add ExtractEmail(a)
                            send "250 Yes, that account exists", smtp(Index)
                            GoTo ExitSub
                        Else
                            send "550 That account does not exist", smtp(Index)
                            GoTo ExitSub
                        End If
                    Else
                        smtp_mailtos(Index).Add ExtractEmail(a)
                        send "251 Recipient added", smtp(Index)
                    End If
                Else
                    send "501 Uh-oh! Bad email address!", smtp(Index)
                    GoTo ExitSub
                End If
            Case "HELP"
                send "214 Enter recipient's emails, one per line (ie RCPT TO: email@domain.com), then DATA on a new line", smtp(Index)
                GoTo ExitSub
            Case "DATA"
                If smtp_mailtos(Index).Count > 0 Then
                    smtp_state(Index) = 3
                    send "354 " & smtp_mailtos(Index).Count & " recipient" & IIf(smtp_mailtos(Index).Count <> 1, "s ", " ") & "specified, enter mail data and end with . on new line", smtp(Index)
                    GoTo ExitSub
                Else
                    send "502 No recipients specified", smtp(Index)
                    GoTo ExitSub
                End If
            Case "QUIT"
                send "221 Ok, sorry to see you go", smtp(Index)
                smtp(Index).Close
                GoTo ExitSub
            Case Else
                send "503 Expected RCPT TO or DATA, not " & b, smtp(Index)
                GoTo ExitSub
        End Select
    ElseIf smtp_state(Index) = 3 Then 'Waiting for (vbCrLf).(vbCrLf)
        Dim body As String, p As Long
        body = smtp_message(Index) & a
        
        p = InStr(1, body, vbCrLf & "." & vbCrLf)
        If p > 0 Then
            smtp_state(Index) = 4
            smtp_message(Index) = body
            send "250 Body completed", smtp(Index)
            GoTo ExitSub
        Else
            smtp_message(Index) = body
            GoTo ExitSub
        End If
    ElseIf smtp_state(Index) = 4 Then 'Waiting for QUIT
        Select Case b
            Case "QUIT"
                For I = 1 To smtp_mailtos(Index).Count
                    NewQueueMail GenMsgID, smtp_sender(Index), smtp_mailtos(Index)(I), smtp_message(Index)
                Next
                send "221 Message sent", smtp(Index)
                smtp(Index).Close
                timMTA_Timer
                GoTo ExitSub
            Case "RSET"
                smtp(Index).Tag = ""
                smtp_sender(Index) = ""
                Set smtp_mailtos(Index) = Nothing: Set smtp_mailtos(Index) = New Collection
                smtp_message(Index) = ""
                
                send "250 Ok message reset", smtp(Index)
                GoTo ExitSub
            Case Else
                send "503 Expecting QUIT or RSET", smtp(Index)
                GoTo ExitSub
        End Select
    End If
    Exit Sub
    
ExitSub:
    Exit Sub
End Sub

Private Sub SysTray_MenuClick(ByVal lIndex As Long, ByVal sKey As String)
    If lIndex = 0 Then
        Me.Show
    ElseIf lIndex = 1 Then
        Dim ab As New frmAbout
        ab.Show vbModal, Me
    ElseIf lIndex = 2 Then
        Quit
    End If
End Sub

Private Sub SysTray_SysTrayDoubleClick(ByVal eButton As MouseButtonConstants)
    If eButton = vbLeftButton Then
        Me.Show
    End If
End Sub

Private Sub SysTray_SysTrayMouseUp(ByVal eButton As MouseButtonConstants)
    If eButton = vbRightButton Then
        SysTray.ShowMenu
    End If
End Sub

Private Sub timMTA_Timer()
    timMTA.Enabled = False
    timMySQL.Enabled = False
    
    modMySQL.Update
    queue.MoveFirst
    Do Until queue.EOF
        mta_msgID = queue.Fields("msgid").value
        mta_sender = queue.Fields("sender").value
        mta_recipient = queue.Fields("recipient").value
        mta_message = queue.Fields("message").value
        If LCase(ExtTarDom(mta_recipient)) = LCase(Domain) Then
            'Keep internal
            If UserExists(ExtUserName(mta_recipient)) = True Then
                NewLocalMail mta_msgID, mta_sender, mta_recipient, mta_message
                RemoveFromQueue mta_msgID
            Else
                MailerDaemon mta_sender, mta_recipient, mta_message
            End If
        Else
            'Forward to correct MTA
            mx.Domain = Mid(mta_recipient, InStr(1, mta_recipient, "@") + 1)
            Dim gmx As String
            gmx = mx.GetMX
            If DebugON = True Then
                MsgBox ":DEBUG:" & vbCrLf & "MTA Connecting to " & gmx & ":25"
            End If
            mta.Close
            mta.connect gmx, 25
            t = Timer + 60
            While Timer < t And mta.State <> 7: DoEvents: Wend
            If mta.State = 7 Then Exit Sub
            If t > Timer Then
                modMySQL.MailerDaemon mta_sender, mta_recipient, mta_message
                DoEvents
            Else
                mta_state = 0
            End If
        End If
        queue.MoveNext
    Loop
    
    timMySQL_Timer
    
    timMTA.Enabled = True
    timMySQL.Enabled = True
End Sub

Private Sub timMySQL_Timer()
    modMySQL.Update
    
    If queue.RecordCount = lvwQueue.ListItems.Count Then Exit Sub
    lvwQueue.ListItems.Clear
    Dim lItem As ListItem
    queue.MoveFirst
    Do Until queue.EOF
        Set lItem = lvwQueue.ListItems.Add()
        lItem.Text = queue.Fields("msgID").value
        lItem.SubItems(1) = queue.Fields("sender").value
        lItem.SubItems(2) = queue.Fields("recipient").value
        lItem.SubItems(3) = queue.Fields("message").value
        queue.MoveNext
    Loop
End Sub

Function ExtTarDom(Email As String) As String
    Dim a As Long
    
    a = InStr(1, Email, "@") + 1
    ExtTarDom = Mid(Email, a)
End Function

Function ExtUserName(Email As String) As String
    Dim a As Long
    
    a = InStr(1, Email, "@") - 1
    ExtUserName = Left(Email, a)
End Function

Sub send(message As String, ws As Winsock)
    If Right(message, 2) <> vbCrLf Then message = message & vbCrLf
    ws.SendData message
    DoEvents
End Sub

Function ExtractEmail(What As String) As String
    Dim re As New RegExp
    re.IgnoreCase = True
    
    re.Pattern = "[abcdefghijklmnopqrstuvwxyz_.-0123456789]{1,64}@[abcdefghijklmnopqrstuvwxyz_.-0123456789]{1,64}\.[abcdefghijklmnopqrstuvwxyz0123456789]{1,6}"
    On Error Resume Next
    ExtractEmail = re.Execute(What)(0)
End Function

Function ExtractIP(What As String) As String
    'extracts the ip address from: (using regexps)
    Dim re As New RegExp
    re.IgnoreCase = True
    
    re.Pattern = "[0123456789]{1,3}\.[0123456789]{1,3}\.[0123456789]{1,3}\.[0123456789]{1,3}"
    On Error Resume Next
    ExtractIP = re.Execute(What)(0)
End Function

Public Function checksum(InString) As String
    checksum = md5.DigestStrToHexStr(CStr(InString))
End Function
