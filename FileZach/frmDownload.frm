VERSION 5.00
Begin VB.Form frmDownload 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Downloading File..."
   ClientHeight    =   1305
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   6375
   Icon            =   "frmDownload.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1305
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "No"
      Height          =   375
      Left            =   5520
      TabIndex        =   5
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Yes"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Would you like to download this file?"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   2580
   End
   Begin VB.Label lblFile 
      AutoSize        =   -1  'True
      Caption         =   "textdoc.txt"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3960
      TabIndex        =   2
      Top             =   120
      Width           =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "is attempting to send you the file"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   2370
   End
   Begin VB.Label lblIP 
      AutoSize        =   -1  'True
      Caption         =   "255.255.255.255"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1395
   End
End
Attribute VB_Name = "frmDownload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public intSocket As Integer

Private Sub Command1_Click()
frmMain.wsInfo(intSocket).SendData "SEND|SEND" & vbCrLf
DoEvents
Unload Me
End Sub


Private Sub Command2_Click()
frmMain.wsInfo(intSocket).SendData "REJECT|REJECT" & vbCrLf
DoEvents
Unload Me
End Sub


