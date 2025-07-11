VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   6645
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8175
   LinkTopic       =   "Form1"
   ScaleHeight     =   6645
   ScaleWidth      =   8175
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Left            =   1320
      Top             =   6120
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   840
      Top             =   6120
   End
   Begin VB.TextBox txtTime 
      Height          =   285
      Left            =   1800
      TabIndex        =   7
      Text            =   "0"
      Top             =   6240
      Width           =   495
   End
   Begin VB.Timer Timer3 
      Interval        =   1000
      Left            =   2400
      Top             =   6120
   End
   Begin VB.CheckBox chkPlay 
      Caption         =   "Auto Play"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   6480
      TabIndex        =   5
      Top             =   1560
      Width           =   1575
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Play"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6480
      TabIndex        =   4
      Top             =   1320
      Width           =   1575
   End
   Begin VB.CheckBox chkPatch 
      Caption         =   "Auto Patch"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4800
      TabIndex        =   3
      Top             =   1560
      Width           =   1575
   End
   Begin VB.CommandButton cmdPatch 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Update Tidel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4800
      Picture         =   "frmMainPatch.frx":0000
      TabIndex        =   2
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   1485
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   240
      Top             =   6000
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   3615
      Left            =   0
      TabIndex        =   1
      Top             =   1800
      Width           =   8175
      ExtentX         =   14420
      ExtentY         =   6376
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   1560
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Max             =   7
      Scrolling       =   1
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   5400
      Width           =   8085
      _ExtentX        =   14261
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6174
            MinWidth        =   6174
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6174
            MinWidth        =   6174
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private fso As New FileSystemObject
Dim FileName, LenFile, FileVer, GetFile As String
Private ontop As New clsOnTop



Private Sub Form_Load()

WebBrowser1.Navigate downloadurl & "/news.html"

  Set ontop = New clsOnTop
  ontop.MakeTopMost hWnd
  
  Status$ = "Idle"
UpdateTime = 0

On Error GoTo skip
Open App.Path & "\patcher.dat" For Input As #1
Input #1, mypatch
Input #1, myplay
Close #1

chkPatch.Value = mypatch
chkPlay.Value = myplay
skip:

If chkPatch.Value = 1 Then
   cmdPatch_Click
End If

End Sub


Private Sub chkPatch_Click()
Open App.Path & "\patcher.dat" For Output As #1
Print #1, chkPatch.Value
Print #1, chkPlay.Value
Close #1
End Sub

Private Sub chkPlay_Click()
Open App.Path & "\patcher.dat" For Output As #1
Print #1, chkPatch.Value
Print #1, chkPlay.Value
Close #1
txtTime = 0
StatusBar1.Panels(1).Text = "Status: Idle"
End Sub

Sub cmdPatch_Click()

    UpdateTime = 0
    Timer2.Interval = 1000
    cmdPatch.Enabled = False
StatusBar1.Panels(1).Text = "Status: Checking for updated files."

'update patcher
ProgressBar1.Value = 1
TransferSuccess = GetInternetFile(Inet1, downloadurl & "/patcher.pch", App.Path)
If TransferSuccess = True Then
Open App.Path & "\patcher.pch" For Input As #2
Do Until (EOF(2))
   Input #2, FileName, LenFile, FileVer, FileMod
   GetFile = App.Path & "\" & FileName
   Text1.Text = Text1.Text & vbCrLf & "Checking " & FileName & "... "
   If fso.FileExists(GetFile) = False Then
      Text1.Text = Text1.Text & "Bad."
      downloadfile ("/" & FileName)
   Else
      If FileName = fso.GetFileName(GetFile) And LenFile = FileLen(GetFile) And FileVer = fso.GetFileVersion(GetFile) Then
         Text1.Text = Text1.Text & "ok."
      Else
         Kill GetFile
         downloadfile ("/" & FileName)
      End If
   End If
Loop
Close #2
Kill App.Path & "\patcher.pch"
End If

'update client
ProgressBar1.Value = 2
TransferSuccess = GetInternetFile(Inet1, downloadurl & "/client.pch", App.Path)
If TransferSuccess = True Then
Open App.Path & "\client.pch" For Input As #3
Do Until (EOF(3))
   Input #3, FileName, LenFile, FileVer, FileMod
   GetFile = App.Path & "\" & FileName
   Text1.Text = Text1.Text & vbCrLf & "Checking " & FileName & "... "
   If fso.FileExists(GetFile) = False Then
      Text1.Text = Text1.Text & "Bad."
      downloadfile ("/" & FileName)
   Else
      If FileName = fso.GetFileName(GetFile) And LenFile = FileLen(GetFile) And FileVer = fso.GetFileVersion(GetFile) Then
         Text1.Text = Text1.Text & "ok."
      Else
      Kill GetFile
         downloadfile ("/" & FileName)
      End If
   End If
Loop
Close #3
Kill App.Path & "\client.pch"
End If

'update client
ProgressBar1.Value = 3
TransferSuccess = GetInternetFile(Inet1, downloadurl & "/client.pch", App.Path)
If TransferSuccess = True Then
Open App.Path & "\client.pch" For Input As #4
Do Until (EOF(4))
   Input #4, FileName, LenFile, FileVer, FileMod
   GetFile = App.Path & "\" & FileName
   Text1.Text = Text1.Text & vbCrLf & "Checking " & FileName & "... "
   If fso.FileExists(GetFile) = False Then
      Text1.Text = Text1.Text & "Bad."
      downloadfile ("/" & FileName)
   Else
      If FileName = fso.GetFileName(GetFile) And LenFile = FileLen(GetFile) And FileVer = fso.GetFileVersion(GetFile) Then
         Text1.Text = Text1.Text & "ok."
      Else
      Kill GetFile
         downloadfile ("/" & FileName)
      End If
   End If
Loop
Close #4
Kill App.Path & "\client.pch"
End If

'update maps
ProgressBar1.Value = 4
TransferSuccess = GetInternetFile(Inet1, downloadurl & "/maps.pch", App.Path)
If TransferSuccess = True Then
Open App.Path & "\maps.pch" For Input As #5
Do Until (EOF(5))
   Input #5, FileName, LenFile, FileVer, FileMod
   GetFile = App.Path & "\maps\" & FileName
   Text1.Text = Text1.Text & vbCrLf & "Checking " & FileName & "... "
   If fso.FileExists(GetFile) = False Then
      Text1.Text = Text1.Text & "Bad."
      downloadfile ("/maps/" & FileName)
   Else
      If FileName = fso.GetFileName(GetFile) And LenFile = FileLen(GetFile) And FileVer = fso.GetFileVersion(GetFile) Then
         Text1.Text = Text1.Text & "ok."
      Else
      Kill GetFile
         downloadfile ("/maps/" & FileName)
      End If
   End If
Loop
Close #5
Kill App.Path & "\maps.pch"
End If

'update midis
ProgressBar1.Value = 5
TransferSuccess = GetInternetFile(Inet1, downloadurl & "/midi.pch", App.Path)
If TransferSuccess = True Then
Open App.Path & "\midi.pch" For Input As #6
Do Until (EOF(6))
   Input #6, FileName, LenFile, FileVer, FileMod
   GetFile = App.Path & "\midis\" & FileName
   Text1.Text = Text1.Text & vbCrLf & "Checking " & FileName & "... "
   If fso.FileExists(GetFile) = False Then
      Text1.Text = Text1.Text & "Bad."
      downloadfile ("midis/" & FileName)
   Else
      If FileName = fso.GetFileName(GetFile) And LenFile = FileLen(GetFile) And FileVer = fso.GetFileVersion(GetFile) Then
         Text1.Text = Text1.Text & "ok."
      Else
      Kill GetFile
         downloadfile ("midis/" & FileName)
      End If
   End If
Loop
Close #6
Kill App.Path & "\midi.pch"
End If

'update skins
ProgressBar1.Value = 6



        ProgressBar1.Value = 7
        Timer2.Interval = 0
        Timer1.Enabled = False
        cmdPlay.Enabled = True
        
        StatusBar1.Panels(1).Text = "Status: Idle"
        
        If chkPlay.Value = 1 Then
            txtTime = 5
        End If

End Sub

Sub downloadfile(file As String)

Text1.Text = Text1.Text & vbCrLf & "Downloading updated " & file & "... "
StatusBar1.Panels(1).Text = "Status: Downloading for updated files."

TransferSuccess = GetInternetFile(Inet1, downloadurl & "/" & file, App.Path)
If TransferSuccess = True Then
      Text1.Text = Text1.Text & "Done."
Else
Text1.Text = Text1.Text & "Error."
End If

StatusBar1.Panels(1).Text = "Status: Checking for updated files."

End Sub


Private Sub Text1_Change()

End Sub

Private Sub Timer1_Timer()

    'StatusBar1.Panels(1).Text = "Status: " & Status$

End Sub

Private Sub Timer2_Timer()
    UpdateTime = UpdateTime + 1
    StatusBar1.Panels(2).Text = "Download Time:" & Str$(UpdateTime) & " Seconds"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub

Private Sub cmdPlay_Click()
msg = MsgBox("Start Game", vbOKOnly)
End Sub

Private Sub Timer3_Timer()
If txtTime <> 0 Then
If txtTime = 1 Then
If chkPlay.Value = 1 Then
txtTime = 0

msg = MsgBox("Start Game", vbOKOnly)

End If
Else

txtTime = txtTime - 1
StatusBar1.Panels(1).Text = "Status: Auto Play in " & txtTime & " second(s)."

End If
End If
End Sub
