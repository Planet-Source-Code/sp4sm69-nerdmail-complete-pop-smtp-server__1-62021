VERSION 5.00
Begin VB.Form frmInputBox 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "[Caption]"
   ClientHeight    =   1335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4950
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "CANCEL"
      Height          =   495
      Left            =   3840
      TabIndex        =   3
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   3840
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox txtInputBox 
      Height          =   390
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   3615
   End
   Begin VB.Label lblPrompt 
      Caption         =   "[Prompt]"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "frmInputBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This form simply serves as an InputBox substitute that supports PasswordChar. It is
'called using the InputBox2() routine located in modMySQL.bas

Public Canceled As Boolean

Private Sub cmdCancel_Click()
    Canceled = True
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    Canceled = False
    Me.Hide
End Sub
