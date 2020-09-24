Attribute VB_Name = "modMain"
Option Explicit

Public strStatus As String
Public lngInputSize As Double
Public Sub AddHost(strHost As String)

Dim intFile As Integer
intFile = FreeFile
Open App.Path & "\Hosts.txt" For Append As #intFile
Print #intFile, strHost
Close #intFile

End Sub


Public Sub GetHost(cbxHosts As ComboBox)
On Error Resume Next

Dim intFile As Integer
Dim strHost As String

intFile = FreeFile

Open App.Path & "\Hosts.txt" For Input As #intFile
Do While EOF(intFile) = False
Line Input #intFile, strHost
If strHost <> "" Then cbxHosts.AddItem strHost Else Exit Do
DoEvents
Loop
Close #intFile

End Sub


Public Sub SaveInputSize()
On Error Resume Next
With frmMain
    If .Option1.Value = True Then lngInputSize = 1024
    If .Option2.Value = True Then lngInputSize = 10240
    If .Option3.Value = True Then lngInputSize = 1024000
    If .Option4.Value = True Then lngInputSize = 10240000
    If .Option5.Value = True Then lngInputSize = .txtBytes.Text
End With
SaveSetting App.Title, "Settings", "UploadSpeed", lngInputSize
End Sub

Public Sub UpdateCaptions()
frmMain.Caption = App.Title & " " & App.Major & "." & App.Minor & App.Revision & " - Super Fast File Transfer [" & strStatus & "]"
End Sub


Public Sub UpdateInputSize()
lngInputSize = GetSetting(App.Title, "Settings", "UploadSpeed", "10240")
On Error Resume Next
With frmMain
    .txtBytes.Enabled = False
    .txtBytes.Text = lngInputSize
    Select Case lngInputSize
    Case "1024"
        .Option1.Value = True
    Case "10240"
        .Option2.Value = True
    Case "1024000"
        .Option3.Value = True
    Case "10240000"
        .Option4.Value = True
    Case Else
        .Option5.Value = True
        .txtBytes.Enabled = True
        .txtBytes.Text = lngInputSize
    End Select
End With

End Sub


