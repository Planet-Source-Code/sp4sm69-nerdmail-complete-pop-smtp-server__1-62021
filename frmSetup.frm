VERSION 5.00
Begin VB.Form frmSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NerdMail Setup"
   ClientHeight    =   7095
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
   Icon            =   "frmSetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "< CANCEL >"
      Height          =   495
      Left            =   2400
      TabIndex        =   27
      Top             =   6480
      Width           =   2175
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "< OK >"
      Default         =   -1  'True
      Height          =   495
      Left            =   120
      TabIndex        =   26
      Top             =   6480
      Width           =   2175
   End
   Begin VB.CheckBox chkSetupUser 
      Caption         =   "Setup first user"
      Height          =   270
      Left            =   240
      TabIndex        =   18
      Top             =   4080
      UseMaskColor    =   -1  'True
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.Frame fraUser 
      Height          =   2295
      Left            =   120
      TabIndex        =   17
      Top             =   4080
      Width           =   4455
      Begin VB.TextBox password2 
         Height          =   390
         IMEMode         =   3  'DISABLE
         Left            =   1680
         PasswordChar    =   "*"
         TabIndex        =   23
         Top             =   1800
         Width           =   2655
      End
      Begin VB.TextBox password 
         Height          =   390
         IMEMode         =   3  'DISABLE
         Left            =   1680
         PasswordChar    =   "*"
         TabIndex        =   21
         Top             =   1440
         Width           =   2655
      End
      Begin VB.TextBox username 
         Height          =   390
         Left            =   1680
         TabIndex        =   19
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label lblEmail 
         Alignment       =   2  'Center
         Caption         =   "Admin Email will be:"
         Height          =   495
         Left            =   120
         TabIndex        =   25
         Tag             =   "Admin Email will be:"
         Top             =   840
         Width           =   4215
      End
      Begin VB.Label lblPassword2 
         Alignment       =   1  'Right Justify
         Caption         =   "Verify Password:"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   1845
         Width           =   1455
      End
      Begin VB.Label lblPassword 
         Alignment       =   1  'Right Justify
         Caption         =   "Password:"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   1485
         Width           =   1455
      End
      Begin VB.Label lblUsername 
         Alignment       =   1  'Right Justify
         Caption         =   "Username:"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   400
         Width           =   1455
      End
   End
   Begin VB.Frame fraDatabase 
      Caption         =   "Database Setup:"
      Height          =   2175
      Left            =   120
      TabIndex        =   7
      Top             =   1800
      Width           =   4455
      Begin VB.CheckBox chkDBSetup 
         Caption         =   "Database must be setup"
         Height          =   270
         Left            =   1800
         TabIndex        =   16
         Top             =   1800
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin VB.TextBox txtM_Name 
         Height          =   390
         Left            =   2040
         TabIndex        =   15
         Top             =   1320
         Width           =   2295
      End
      Begin VB.TextBox txtM_Pass 
         Height          =   390
         IMEMode         =   3  'DISABLE
         Left            =   2040
         PasswordChar    =   "*"
         TabIndex        =   13
         Top             =   960
         Width           =   2295
      End
      Begin VB.TextBox txtM_User 
         Height          =   390
         Left            =   2040
         TabIndex        =   11
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox txtM_Host 
         Height          =   390
         Left            =   2040
         TabIndex        =   9
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label lblM_Database 
         Alignment       =   1  'Right Justify
         Caption         =   "Database Name:"
         Height          =   255
         Left            =   480
         TabIndex        =   14
         Top             =   1365
         Width           =   1455
      End
      Begin VB.Label lblM_Password 
         Alignment       =   1  'Right Justify
         Caption         =   "MySQL Password:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1005
         Width           =   1815
      End
      Begin VB.Label lblM_Username 
         Alignment       =   1  'Right Justify
         Caption         =   "MySQL Username:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   645
         Width           =   1815
      End
      Begin VB.Label lblM_Server 
         Alignment       =   1  'Right Justify
         Caption         =   "MySQL Server:"
         Height          =   255
         Left            =   480
         TabIndex        =   8
         Top             =   285
         Width           =   1455
      End
   End
   Begin VB.Frame fraServer 
      Caption         =   "Server Setup:"
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.TextBox txtDomain 
         Height          =   390
         Left            =   1680
         TabIndex        =   6
         Top             =   1080
         Width           =   2655
      End
      Begin VB.ComboBox cboSMTPIP 
         Appearance      =   0  'Flat
         Height          =   390
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   720
         Width           =   2655
      End
      Begin VB.ComboBox cboPOPIP 
         Appearance      =   0  'Flat
         Height          =   390
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label lblDomain 
         Alignment       =   1  'Right Justify
         Caption         =   "Root Domain:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1120
         Width           =   1455
      End
      Begin VB.Label lblSMTPIP 
         Alignment       =   1  'Right Justify
         Caption         =   "SMTP Server IP:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   760
         Width           =   1455
      End
      Begin VB.Label lblPOPIP 
         Alignment       =   1  'Right Justify
         Caption         =   "POP Server IP:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   400
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public SetupDone As Boolean

Private Sub chkSetupUser_Click()
    If chkSetupUser.value = 0 Then
        fraUser.Enabled = False
    Else
        fraUser.Enabled = True
    End If
End Sub

Private Sub cmdCancel_Click()
    SetupDone = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    SetupDone = True
    
    If txtDomain.Text = "" Then GoTo NotComplete
    
    If txtM_Host.Text = "" Then GoTo NotComplete
    If txtM_User.Text = "" Then GoTo NotComplete
    If txtM_Pass.Text = "" Then GoTo NotComplete
    If txtM_Name.Text = "" Then GoTo NotComplete
    
    If chkSetupUser.value = 1 Then
        If username = "" Then GoTo NotComplete
        If password = "" Then GoTo NotComplete
        If password2 = "" Then GoTo NotComplete
        
        If password <> password Then
            MsgBox "Passwords do not match!"
            
            On Error Resume Next: password.SetFocus: On Error GoTo 0
            
            password.SelStart = 0
            password.SelLength = Len(password.Text)
        End If
    End If
    
    SaveSetting App.Title, "Server Settings", "POP IP", cboPOPIP.Text
    SaveSetting App.Title, "Server Settings", "SMTP IP", cboSMTPIP.Text
    SaveSetting App.Title, "Server Settings", "Domain", txtDomain.Text
    
    SaveSetting App.Title, "Database Settings", "Hostname", txtM_Host.Text
    SaveSetting App.Title, "Database Settings", "Username", txtM_User.Text
    SaveSetting App.Title, "Database Settings", "Password", txtM_Pass.Text
    SaveSetting App.Title, "Database Settings", "Database", txtM_Name.Text
    
    If chkDBSetup.value = 1 Then
        Dim db As MYSQL_CONNECTION
        Set db = New MYSQL_CONNECTION
        
        db.OpenConnection txtM_Host, txtM_User, txtM_Pass, ""
        If db.State <> MY_CONN_OPEN Then
            MsgBox "Database error!", vbOKOnly, "NerdMail Setup"
            Exit Sub
        End If
        
        db.Execute "CREATE DATABASE " & txtM_Name
        db.SelectDb txtM_Name
        db.Execute "CREATE TABLE users (username VARCHAR(20), password VARCHAR(20))"
        db.Execute "CREATE TABLE email (msgID VARCHAR(10), sender TINYTEXT, recipient TINYTEXT, message LONGTEXT)"
        db.Execute "CREATE TABLE queue LIKE email"
    End If
    
    If chkSetupUser.value = 1 Then NewUser username, password
    frmMain.SetupDone = True
    Unload Me
    Exit Sub
    
NotComplete:
    MsgBox "Please complete all fields.", vbOKOnly, "NerdMail Setup"
    Exit Sub
End Sub

Private Sub Form_Load()
    ip_array.GetIPArray
    For I = 0 To UBound(ip_array.IPArray)
        cboPOPIP.AddItem ip_array.IPArray(I)
        cboSMTPIP.AddItem ip_array.IPArray(I)
    Next
    
    cboPOPIP.ListIndex = 0
    cboSMTPIP.ListIndex = 0
    
    txtDomain.Text = frmMain.pop3(0).LocalHostName
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain.SetupDone = Me.SetupDone
End Sub

Private Sub txtDomain_Change()
    lblEmail.Caption = lblEmail.Tag & vbCrLf & username & "@" & txtDomain
End Sub

Private Sub username_Change()
    lblEmail.Caption = lblEmail.Tag & vbCrLf & username & "@" & txtDomain
End Sub
