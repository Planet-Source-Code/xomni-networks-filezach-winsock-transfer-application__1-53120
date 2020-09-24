VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   9660
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
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
   ScaleHeight     =   5550
   ScaleWidth      =   9660
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Options"
      Height          =   2415
      Left            =   120
      TabIndex        =   11
      Top             =   3120
      Width           =   2415
      Begin VB.OptionButton Option5 
         Caption         =   "Custom (in bytes)"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1680
         Width           =   2055
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Small (1kb)"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   720
         Width           =   1455
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Huge (10mb) *"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1440
         Width           =   1455
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Large (1mb)"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1200
         Width           =   1455
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Medium (10kb)"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   960
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.TextBox txtBytes 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   13
         Top             =   2040
         Width           =   2175
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "* Not Recommended"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   720
         MouseIcon       =   "frmMain.frx":1CFA
         MousePointer    =   99  'Custom
         TabIndex        =   19
         Top             =   120
         Width           =   1575
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "File Access Speed"
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
         Left            =   720
         TabIndex        =   12
         Top             =   360
         Width           =   1515
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         MouseIcon       =   "frmMain.frx":2004
         MousePointer    =   99  'Custom
         Picture         =   "frmMain.frx":230E
         ToolTipText     =   "Whats this?"
         Top             =   240
         Width           =   480
      End
   End
   Begin MSWinsockLib.Winsock wsListenInfo 
      Left            =   3120
      Top             =   960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   200
   End
   Begin MSWinsockLib.Winsock wsRecieve 
      Index           =   0
      Left            =   3600
      Top             =   1440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Caption         =   "Recieving files..."
      Height          =   5535
      Left            =   2640
      TabIndex        =   6
      Top             =   0
      Width           =   6975
      Begin MSWinsockLib.Winsock wsListenRecieve 
         Left            =   480
         Top             =   1440
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         LocalPort       =   201
      End
      Begin MSWinsockLib.Winsock wsInfo 
         Index           =   0
         Left            =   960
         Top             =   960
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Abort"
         Height          =   255
         Left            =   6120
         TabIndex        =   9
         Top             =   5160
         Width           =   735
      End
      Begin MSComctlLib.ProgressBar pbDownload 
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   5160
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSComctlLib.ListView lvFiles 
         Height          =   4575
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   8070
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Status"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Filename"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Downloaded"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Host Address"
            Object.Width           =   2646
         EndProperty
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "To update the progress bar on the file download, click on a file transfer"
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
         Width           =   6705
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "File Sharing"
      Enabled         =   0   'False
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2415
      Begin VB.TextBox Text1 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   960
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   2760
         Width           =   1335
      End
      Begin VB.CheckBox chkShare 
         Caption         =   "Share files in this folder"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   2520
         Width           =   2175
      End
      Begin VB.DirListBox Dir1 
         Height          =   1890
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   2175
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Password:"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   2820
         Width           =   750
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileEnable 
         Caption         =   "Enable"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFileDisable 
         Caption         =   "Disable"
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileQuit 
         Caption         =   "Quit"
      End
   End
   Begin VB.Menu mnuTransfers 
      Caption         =   "Transfers"
      Begin VB.Menu mnuFileTransfers 
         Caption         =   "Send..."
      End
      Begin VB.Menu mnuTransfersDisconnect 
         Caption         =   "Disconnect all"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuHelpSupport 
         Caption         =   "Support"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "About"
      End
      Begin VB.Menu mnuHelpSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpVote 
         Caption         =   "Vote / Add Feedback"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FileSize(999) As Double
Dim FileInput(999) As Double
Private Sub Command1_Click()

On Error Resume Next

Select Case UCase(Command1.Caption)
Case "CLEAR"
lvFiles.SelectedItem.Text = ""
lvFiles.SelectedItem.SubItems(1) = ""
lvFiles.SelectedItem.SubItems(2) = ""
lvFiles.SelectedItem.SubItems(3) = ""
Case "ABORT"
lvFiles.SelectedItem.Text = "Aborted"
wsInfo(lvFiles.SelectedItem.Index - 1).SendData "ABORT|ABORT" & vbCrLf
wsInfo(lvFiles.SelectedItem.Index - 1).Close
wsTransfer(lvFiles.SelectedItem.Index - 1).Close
End Select

End Sub

Private Sub Form_Load()
frmAbout.Show vbModal, Me 'About
strStatus = "Ready"
UpdateCaptions 'Set Form and Tray Captions
wsListenInfo.Listen
wsListenRecieve.Listen

lngInputSize = GetSetting(App.Title, "Settings", "UploadSpeed", "10240")
UpdateInputSize
End Sub


Private Sub Image1_Click()
MsgBox "This option selects how many bytes are read from the file at a time. A selection of 1KB (1024 bytes) would make FileZach read 1024 bytes at a time, and the send the file once the whole file has been read into memory. If you have a slow computer, selecting high read rates may cause crashes and freezes. You may want to work with the speeds finding one your computer can handle comfortably.", vbInformation, "What is File Access Speed?"
End Sub


Private Sub Label4_Click()
MsgBox "Such huge file input sizes can easily crash or freeze your system. I recommend you choose a slower input speed. In addition to your own computer speed, choosing a high input speed could also cause problems for the remote system.", vbInformation, "Whats this?"
End Sub

Private Sub lvFiles_ItemClick(ByVal item As MSComctlLib.ListItem)

On Error Resume Next
Select Case item.Text
Case "Aborted", "Completed"
Command1.Caption = "Clear"
Case "Downloading"
Command1.Caption = "Abort"
End Select

pbDownload.Max = FileSize(item.Index - 1)
pbDownload.Value = FileInput(item.Index - 1)
End Sub


Private Sub mnuFileDisable_Click()
wsListenInfo.Close
wsListenRecieve.Close
mnuFileDisable.Enabled = False
mnuFileEnable.Enabled = True
End Sub

Private Sub mnuFileEnable_Click()
wsListenInfo.Listen
wsListenRecieve.Listen
mnuFileDisable.Enabled = True
mnuFileEnable.Enabled = False
End Sub


Private Sub mnuFileQuit_Click()
If MsgBox("Are you sure you want to shutdown FileZach?", vbQuestion + vbYesNo) = vbYes Then End

End Sub

Private Sub mnuFileTransfers_Click()
Dim frmS As New frmSend
frmS.Show , Me
End Sub


Private Sub mnuHelpAbout_Click()
frmAbout.Show , Me
End Sub

Private Sub mnuHelpSupport_Click()
MsgBox "I hope to add online and offline support files at a future time. For now, please either email me (admin@xomni.net) or post feedback.", vbExclamation
End Sub


Private Sub mnuTransfersDisconnect_Click()
On Error Resume Next

Dim i As Integer
For i = 0 To wsInfo.Count - 1
wsInfo(i).Close
wsRecieve(i).Close
lvFiles.ListItems.Clear
Next i
End Sub

Private Sub Option1_Click()
SaveInputSize
UpdateInputSize
End Sub

Private Sub Option2_Click()
SaveInputSize
UpdateInputSize
End Sub


Private Sub Option3_Click()
SaveInputSize
UpdateInputSize
End Sub


Private Sub Option4_Click()
SaveInputSize
UpdateInputSize
End Sub


Private Sub Option5_Click()
SaveInputSize
txtBytes.Enabled = True
End Sub


Private Sub txtBytes_Change()
lngInputSize = txtBytes.Text
SaveInputSize
Option5.Value = True
End Sub


Private Sub wsInfo_Close(Index As Integer)
Dim item As ListItem
Set item = lvFiles.ListItems.item(Index + 1)
If item.Text <> "Complete" Then item.Text = "Aborted"
wsInfo(Index).Close
End Sub

Private Sub wsInfo_DataArrival(Index As Integer, ByVal bytesTotal As Long)
On Error Resume Next

Dim strInput As String
Dim strData() As String
Dim strParse() As String
Dim strItem As String
Dim strText As String
Dim item As ListItem
Dim i As Integer

wsInfo(Index).GetData strInput, , bytesTotal
strData = Split(strInput, vbCrLf)
For i = 0 To UBound(strData)
strText = ""
strText = strData(i)
If strText <> "" Then
strParse = Split(strText, "|")
strText = strParse(0)
strItem = strParse(1)
Set item = lvFiles.ListItems(Index + 1)
    Select Case UCase(strText)
    Case "FILENAME"
        item.SubItems(1) = strItem
        Kill App.Path & "\Downloads\" & lvFiles.ListItems(Index + 1).SubItems(1)
    Case "FILEDATE"
        item.Tag = strItem
        Dim frmDL As New frmDownload
        frmDL.lblFile.Caption = item.SubItems(1)
        frmDL.lblIP.Caption = item.SubItems(3)
        frmDL.intSocket = Index
        frmDL.Show , Me
    Case "FILESIZE"
        FileSize(Index) = strItem
        item.SubItems(2) = "0/" & strItem
    Case "SENDING"
        item.Text = "Downloading"
    Case "COMPLETE"
        item.Text = "Complete"
    Case "SEND"
        item.Text = "Downloading"
    End Select

End If
Next i
End Sub
    
Private Sub wsInfo_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Dim item As ListItem
Set item = lvFiles.ListItems.item(Index + 1)
item.Text = "Error"
wsInfo(Index).Close
wsInfo(Index).Close
End Sub

Private Sub wsListenInfo_ConnectionRequest(ByVal requestID As Long)
On Error Resume Next

Dim x As String
Dim i As Integer
For i = 0 To wsInfo.Count - 1
If wsInfo(i).State = sckClosed Then
wsInfo(i).Accept requestID
Dim item As ListItem

x = lvFiles.ListItems(i + 1)
If x = "" Then
Set item = lvFiles.ListItems.Add(i + 1)
Else
Set item = lvFiles.ListItems(i + 1)
End If

item.Text = "Waiting..."
item.SubItems(1) = "Unknown"
item.SubItems(2) = "Unknown"
item.SubItems(3) = wsListenInfo.RemoteHostIP
Exit Sub
End If
Next i
i = wsInfo.Count
Load wsInfo(i)
wsInfo(i).Accept requestID

x = lvFiles.ListItems(i + 1)
If x = "" Then
Set item = lvFiles.ListItems.Add(i + 1)
Else
Set item = lvFiles.ListItems(i + 1)
End If

Set item = lvFiles.ListItems.Add(i + 1)
item.Text = "Waiting..."
item.SubItems(1) = "Unknown"
item.SubItems(2) = "Unknown"
item.SubItems(3) = wsListenInfo.RemoteHostIP

End Sub

Private Sub wsListenRecieve_ConnectionRequest(ByVal requestID As Long)
On Error Resume Next

Dim x As String
Dim i As Integer
For i = 0 To wsRecieve.Count - 1
If wsRecieve(i).State = sckClosed Then
wsRecieve(i).Accept requestID
Exit Sub
End If
Next i

i = wsRecieve.Count
Load wsRecieve(i)
wsRecieve(i).Accept requestID

End Sub

Private Sub wsRecieve_Close(Index As Integer)
wsRecieve(Index).Close
End Sub

Private Sub wsRecieve_DataArrival(Index As Integer, ByVal bytesTotal As Long)
On Error Resume Next

Dim data As String
wsRecieve(Index).GetData data, , bytesTotal
Dim intFile As Integer
intFile = FreeFile
MkDir App.Path & "\Downloads"
Open App.Path & "\Downloads\" & lvFiles.ListItems(Index + 1).SubItems(1) For Binary As #intFile
Put #intFile, LOF(intFile) + 1, data
Close #intFile
FileInput(Index) = FileInput(Index) + Len(data)

Dim item As ListItem
Set item = lvFiles.ListItems(Index + 1)
item.SubItems(2) = FileInput(Index) & "/" & FileSize(Index)

wsInfo(Index).SendData "INFO|" & FileInput(Index) & vbCrLf
End Sub

Private Sub wsRecieve_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
wsRecieve(Index).Close
End Sub


