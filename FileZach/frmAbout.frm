VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "FileZach: Super Fast File Sender v1.02"
   ClientHeight    =   6810
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   6765
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
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
   ScaleHeight     =   6810
   ScaleWidth      =   6765
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   2760
      Top             =   3120
   End
   Begin VB.CommandButton Command2 
      Caption         =   "I would like to vote for FileZach (and remove the wait period)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   6360
      Width           =   6495
   End
   Begin VB.CommandButton Command1 
      Caption         =   ":: Please read the terms of agreement ::"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   5880
      Width           =   6495
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2760
      Top             =   2640
   End
   Begin VB.TextBox Text1 
      Height          =   4695
      Left            =   4200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Text            =   "frmAbout.frx":0442
      Top             =   720
      Width           =   2415
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5775
      Left            =   0
      Picture         =   "frmAbout.frx":06A5
      ScaleHeight     =   5775
      ScaleWidth      =   4095
      TabIndex        =   0
      Top             =   0
      Width           =   4095
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "A picture of my friend, Zach, and his stupendoes speed."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   1320
         TabIndex        =   10
         Top             =   5280
         Width           =   2655
      End
      Begin VB.Label lblFiles 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   600
         MouseIcon       =   "frmAbout.frx":4E423
         MousePointer    =   99  'Custom
         TabIndex        =   8
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "- Select a File -"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1935
         Left            =   480
         TabIndex        =   7
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   """I'll take that file for you"""
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1560
         TabIndex        =   9
         Top             =   720
         Width           =   1935
      End
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0060D8FF&
      BorderWidth     =   5
      Height          =   375
      Left            =   120
      Top             =   6360
      Width           =   6495
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Visit our Website"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   4200
      TabIndex        =   4
      Top             =   5450
      Width           =   2415
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Version 1.02 created by Xomni Networks (www.xomni.net)"
      Height          =   495
      Left            =   4200
      TabIndex        =   2
      Top             =   240
      Width           =   2415
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "FileZach"
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
      Left            =   5040
      TabIndex        =   1
      Top             =   0
      Width           =   690
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim intCounter As Integer
Private Sub Command1_Click()
Hide
End Sub

Private Sub Form_Load()
intCounter = 10
lblFiles.Caption = "assorted.zip" & vbCrLf & "pictures.jpg" & vbCrLf & "game.exe" & vbCrLf & "movie.mpg" & vbCrLf & "music.mp3" & vbCrLf & "report.doc" & vbCrLf & "cursors.ico" & vbCrLf & "code.frm" & vbCrLf & "website.htm"
Load frmMain
End Sub

Private Sub lblFiles_Click()
MsgBox "FileZach's binary transfer allows you to send all types of files, despite their size! FileZach currently was tested on sending a 70mb file on the 'Medium' setting. If you encounter any problems please email me so future versions will be correct.", vbInformation, "Whats this?"
End Sub


Private Sub Timer1_Timer()

intCounter = intCounter - 1
Command1.Caption = "Waiting: " & intCounter
If intCounter = -1 Then
Command1.Caption = "I accept to the terms of agreement"
Command1.Enabled = True
Timer1.Enabled = False
Exit Sub
End If

End Sub


Private Sub Timer2_Timer()
Shape2.Visible = Not Shape2.Visible
End Sub


