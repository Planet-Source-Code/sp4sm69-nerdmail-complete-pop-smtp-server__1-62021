VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About NerdMail v1.0"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4080
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   4080
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "< OK >"
      Default         =   -1  'True
      Height          =   495
      Left            =   1313
      TabIndex        =   3
      Top             =   2640
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   0
      Picture         =   "frmAbout.frx":058A
      ScaleHeight     =   101
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   269
      TabIndex        =   0
      Top             =   0
      Width           =   4095
   End
   Begin VB.Label lblCodedBy 
      Alignment       =   2  'Center
      Caption         =   "< Coded by King of Geeks >"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Width           =   3735
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "< NerdMail v1.0 >"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1680
      Width           =   3855
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
    Unload Me
End Sub
