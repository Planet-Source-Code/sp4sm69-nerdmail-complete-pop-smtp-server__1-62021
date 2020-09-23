VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmUsers 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NerdMail User Management"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4695
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUsers.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.ListView lvwUsers 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   5106
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Username"
         Object.Width           =   3351
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Password"
         Object.Width           =   3351
      EndProperty
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "POPUP MENU"
      Visible         =   0   'False
      Begin VB.Menu mnuAddUser 
         Caption         =   "&Add User..."
      End
      Begin VB.Menu mnuEditUser 
         Caption         =   "&Edit User..."
      End
      Begin VB.Menu mnuDeleteUser 
         Caption         =   "&Delete user..."
      End
   End
End
Attribute VB_Name = "frmUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    Dim li As ListItem
    
    modMySQL.Update
    
    users.MoveFirst
    Do Until users.EOF
        Set li = lvwUsers.ListItems.Add
        li.Text = users.Fields("username").value
        li.SubItems(1) = users.Fields("password").value
        users.MoveNext
    Loop
End Sub

Private Sub lvwUsers_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        PopupMenu mnuPopup
    End If
End Sub

Private Sub mnuAddUser_Click()
    Dim un As String, pw As String, pw2 As String
    
top:
    un = InputBox2("NerdMail User Setup", "Enter username (username@" & frmMain.Domain & "):")
    pw = InputBox2("NerdMail User Setup", "Enter password:", "", "*")
    pw2 = InputBox2("NerdMail User Setup", "Verify password:", "", "*")
    
    If pw <> pw2 Then
        If MsgBox("Passwords do not match! Try again?", vbYesNo) = vbYes Then
            GoTo top
            Exit Sub
        Else
            Exit Sub
        End If
    End If
    
    NewUser un, pw
    UpdateList
End Sub

Private Sub mnuDeleteUser_Click()
    If MsgBox("Are you sure?", vbYesNo, "NerdMail") = vbNo Then Exit Sub
    
    modMySQL.Execute "DELETE FROM users WHERE username='" & lvwUsers.SelectedItem.Text & "'"
    
    UpdateList
End Sub

Private Sub mnuEditUser_Click()
    Dim un As String, pw As String, pw2 As String
    
top:
    un = InputBox2("NerdMail User Setup", "Enter username (username@" & frmMain.Domain & "):", lvwUsers.SelectedItem.Text)
    If un = "" Then Exit Sub
    pw = InputBox2("NerdMail User Setup", "Enter password:", lvwUsers.SelectedItem.SubItems(1), "*")
    If pw = "" Then Exit Sub
    
    If pw <> lvwUsers.SelectedItem.SubItems(1) Then
        pw2 = InputBox2("NerdMail User Setup", "Verify password:", lvwUsers.SelectedItem.SubItems(1), "*")
        If pw2 = "" Then Exit Sub
    End If
    
    If pw <> pw2 Then
        If MsgBox("Passwords do not match. Try again?", vbYesNo, "NerdMail") = vbNo Then Exit Sub
        GoTo top
    End If
    
    users.MoveFirst
    Do Until users.Fields("username").value = lvwUsers.SelectedItem.Text Or users.EOF
        users.MoveNext
    Loop
    
    If users.Fields("username").value <> lvwUsers.SelectedItem.Text Then
        MsgBox "Username not found!"
        Exit Sub
    End If
    
    sql = "UPDATE users SET username='" & un & "', password='" & pw & "' WHERE username='" & lvwUsers.SelectedItem.Text & "'"
    modMySQL.Execute CStr(sql)
    DoEvents
    UpdateList
End Sub

Sub UpdateList()
    Dim li As ListItem
    modMySQL.Update
    
    lvwUsers.ListItems.Clear
    users.MoveFirst
    
    Do Until users.EOF
        Set li = lvwUsers.ListItems.Add
        li.Text = users.Fields("username").value
        li.SubItems(1) = users.Fields("password").value
        users.MoveNext
    Loop
End Sub
