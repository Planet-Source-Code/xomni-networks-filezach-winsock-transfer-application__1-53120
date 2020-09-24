VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSend 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Send a file..."
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   5850
   Icon            =   "frmSend.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   5850
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command3 
      Caption         =   "Send"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   11
      Top             =   2400
      Width           =   975
   End
   Begin MSWinsockLib.Winsock wsTransfer 
      Left            =   3720
      Top             =   840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   201
   End
   Begin MSWinsockLib.Winsock wsInfo 
      Left            =   3240
      Top             =   840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   200
   End
   Begin MSComDlg.CommonDialog opendialog 
      Left            =   2640
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Select file for transfer"
      Filter          =   "*.*"
      MaxFileSize     =   11520
   End
   Begin VB.CommandButton Command2 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5160
      TabIndex        =   4
      ToolTipText     =   "Add a friend"
      Top             =   480
      Width           =   615
   End
   Begin VB.ComboBox cbxHosts 
      Height          =   315
      Left            =   840
      TabIndex        =   3
      Top             =   480
      Width           =   4215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5160
      TabIndex        =   1
      ToolTipText     =   "Select a file"
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox txtFile 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "Select a File --->"
      Top             =   120
      Width           =   4935
   End
   Begin VB.Frame Frame1 
      Caption         =   "Progress"
      Height          =   1335
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   5655
      Begin MSComctlLib.ProgressBar pbTotal 
         Height          =   255
         Left            =   600
         TabIndex        =   8
         Top             =   480
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin MSComctlLib.ProgressBar pbUploaded 
         Height          =   255
         Left            =   600
         TabIndex        =   9
         Top             =   960
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label lblRecv 
         AutoSize        =   -1  'True
         Caption         =   "0 / 0 bytes"
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
         Left            =   2640
         TabIndex        =   13
         Top             =   720
         Width           =   780
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Recieved by Host Computer:"
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
         TabIndex        =   10
         Top             =   720
         Width           =   2400
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Sent:"
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
         TabIndex        =   7
         Top             =   240
         Width           =   435
      End
      Begin VB.Label lblBytes 
         AutoSize        =   -1  'True
         Caption         =   "0 / 0 bytes"
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
         Left            =   600
         TabIndex        =   6
         Top             =   240
         Width           =   4785
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Send to"
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
      TabIndex        =   2
      Top             =   525
      Width           =   645
   End
End
Attribute VB_Name = "frmSend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FileSize As Double
Dim FileSent As Double
Public Sub Restart()
FileSize = 0
FileSent = 0
pbTotal.Value = 0
pbTotal.Max = 0
pbUploaded.Value = 0
pbUploaded.Max = 0
lblBytes = "0 / 0 bytes"
lblRecv = "0 / 0 bytes"
Command3.Enabled = True
End Sub

Public Sub SendFile()
On Error GoTo 1

pbTotal.Value = 0
pbUploaded.Value = 0
pbUploaded.Max = FileSize
pbTotal.Max = FileSize
Caption = "Reading and Uploading File..."

Dim intFile As Integer
Dim strData As String
Dim strInput As String
Dim dblSent As Double
intFile = FreeFile

Open opendialog.FileName For Binary As #intFile
'Check for small file
If FileSize <= lngInputSize Then
strInput = Space(FileSize)
Get #intFile, , strInput
wsTransfer.SendData strInput
GoTo 1
End If
'More than 1KB - Start Getting Pieces in 1024 bytes

Do While EOF(intFile) = False
If dblSent + lngInputSize <= FileSize Then
strInput = Space(lngInputSize)
Get #intFile, , strInput
wsTransfer.SendData strInput
dblSent = dblSent + lngInputSize
DoEvents
Else
strInput = Space(FileSize - dblSent)
Get #intFile, , strInput
wsTransfer.SendData strInput
Exit Do
End If
DoEvents
Loop
1:
pbTotal.Value = FileSize
wsInfo.SendData "COMPLETE|COMPLETE" & vbCrLf
DoEvents
MsgBox opendialog.FileName & " was sent successfully!", vbInformation
DoEvents
Close #intFile
Unload Me
End Sub

Private Sub Command1_Click()
Dim strAdd As String

opendialog.ShowOpen
If opendialog.FileName <> "" And opendialog.CancelError = False Then
strAdd = opendialog.FileName
txtFile.Text = strAdd
Tag = Mid(strAdd, InStrRev(strAdd, "\") + 1)
Caption = "Ready to send " & Tag
End If

End Sub

Private Sub Command2_Click()
Dim strInput As String

strInput = InputBox("Enter host name or IP Address", "Add Connection", "127.0.0.1")
If strInput <> "" Then
cbxHosts.AddItem strInput
modMain.AddHost strInput
End If

End Sub


Private Sub Command3_Click()
Command3.Enabled = False
wsInfo.Connect cbxHosts.Text
wsTransfer.Connect cbxHosts.Text
End Sub


Private Sub Form_Load()
modMain.GetHost cbxHosts
End Sub

Private Sub wsInfo_Close()
Caption = "Upload failed: " & "connection closed!"
wsInfo.Close
End Sub

Private Sub wsInfo_Connect()
Caption = "Connected to " & wsInfo.RemoteHostIP
wsInfo.SendData "Filename|" & Tag & vbCrLf
DoEvents
FileSize = FileLen(opendialog.FileName)
wsInfo.SendData "Filesize|" & FileSize & vbCrLf
DoEvents
wsInfo.SendData "Filedate|" & FileDateTime(opendialog.FileName) & vbCrLf
End Sub


Private Sub wsInfo_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next

Dim strData As String
Dim strParse() As String
Dim strText() As String
Dim i As Integer
Dim strItem As String
wsInfo.GetData strData, , bytesTotal
strParse = Split(strData, vbCrLf)
For i = 0 To UBound(strParse)
strData = strParse(i)
If strData <> "" Then
strText = Split(strData, "|")
Select Case strText(0)
Case "ABORT"
MsgBox "Connection lost!", vbExclamation
Unload Me
Case "REJECT"
Caption = "File REJECTED by remote user!"
Restart
Case "SEND"
wsInfo.SendData "SEND|SENDING" & vbCrLf
SendFile
Case "INFO"
lblRecv.Caption = strText(1) & " / " & FileSize & " bytes"
pbUploaded.Value = strText(1)
If pbUploaded.Value = pbUploaded.Max Then Restart
DoEvents
End Select
End If
Next i

End Sub

Private Sub wsInfo_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Caption = "Upload failed: " & Description
wsInfo.Close
End Sub


Private Sub wsTransfer_SendComplete()
On Error Resume Next
DoEvents
End Sub


Private Sub wsTransfer_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
On Error Resume Next
FileSent = FileSent + bytesSent
pbTotal.Value = FileSent
lblBytes.Caption = FileSent & " / " & FileSize
DoEvents
End Sub


