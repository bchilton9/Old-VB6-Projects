VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmPatcher 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "The Land's of Tidel - Patcher"
   ClientHeight    =   5190
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   7005
   Icon            =   "liveu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   7005
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer3 
      Interval        =   1000
      Left            =   240
      Top             =   720
   End
   Begin VB.TextBox txtTime 
      Height          =   285
      Left            =   6000
      TabIndex        =   8
      Text            =   "0"
      Top             =   240
      Visible         =   0   'False
      Width           =   495
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   3495
      Left            =   0
      TabIndex        =   7
      Top             =   1320
      Width           =   6975
      ExtentX         =   12303
      ExtentY         =   6165
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
      Left            =   1440
      Picture         =   "liveu.frx":08CA
      TabIndex        =   6
      Top             =   720
      Width           =   1575
   End
   Begin VB.CheckBox chkPatch 
      Caption         =   "Auto Patch"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1440
      TabIndex        =   5
      Top             =   960
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
      Left            =   3840
      TabIndex        =   4
      Top             =   720
      Width           =   1575
   End
   Begin VB.CheckBox chkPlay 
      Caption         =   "Auto Play"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3840
      TabIndex        =   3
      Top             =   960
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   840
      Top             =   720
   End
   Begin VB.Timer Timer2 
      Left            =   5520
      Top             =   720
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   1080
      TabIndex        =   1
      Top             =   360
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Max             =   3
      Scrolling       =   1
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   4800
      Width           =   7005
      _ExtentX        =   12356
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
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   3120
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Click Update Tidel to start."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1080
      TabIndex        =   2
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "frmPatcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private ontop As New clsOnTop

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
status$ = "Idle"
End Sub

Private Sub cmdPatch_Click()

    Dim TransferSuccess As Boolean
    UpdateTime = 0
    Timer2.Interval = 1000
    cmdPatch.Enabled = False
    ProgressBar1.Value = 1
    status$ = "Checking for updated version."
    Label1.Caption = "Checking for updated version."
    TransferSuccess = GetInternetFile(Inet1, downloadurl & "/TidelUpdate.dat", App.Path)

    If TransferSuccess = False Then
        ProgressBar1.Value = 3
        Timer2.Interval = 0
        Label1.Caption = "Error connecting to patch server."
        cmdPlay.Enabled = True
        If chkPlay.Value = 1 Then
txtTime = 5
        End If
        Exit Sub
    End If
       
    ProgressBar1.Value = 2
    
    status$ = "Version check success."
    
    On Error GoTo pass
    Open App.Path & "\TidelUpdate.dat" For Input As #1
        Input #1, updatever$
pass:
    Close #1

On Error GoTo skip
    Kill App.Path & "\TidelUpdate.dat"
skip:
     
    If updatever$ > myVer Then
        Label1.Caption = "There is an update available to version " + updatever
    Else
        Label1.Caption = "There is no update available."
        ProgressBar1.Value = 3
        Timer2.Interval = 0
        cmdPlay.Enabled = True
        If chkPlay.Value = 1 Then
txtTime = 5
        End If
        Exit Sub
    End If

    status$ = "Getting updated file."

    TransferSuccess = GetInternetFile(Inet1, downloadurl & "/TidelUpdate.exe", App.Path)

    If TransferSuccess = False Then
        ProgressBar1.Value = 3
        Command1.Enabled = True
        Timer2.Interval = 0
        Label1.Caption = "Error connecting to download server."
        cmdPlay.Enabled = True
        If chkPlay.Value = 1 Then
txtTime = 5
        End If
        Exit Sub
    End If
    
    ProgressBar1.Value = 3
    Timer2.Interval = 0
    
    run = Shell(App.Path & "\Tidelupdate.exe", vbNormalFocus)
    End

End Sub

Private Sub cmdPlay_Click()
        Load frmLogin
        frmLogin.Show
        Unload Me
End Sub

Private Sub Form_Load()

WebBrowser1.Navigate downloadurl & "/news.html"

  Set ontop = New clsOnTop
  ontop.MakeTopMost hWnd

myVer = App.Major & "." & App.Minor & "." & App.Revision

status$ = "Idle"
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

Private Sub Timer1_Timer()

    StatusBar1.Panels(1).Text = "Status: " & status$

End Sub

Private Sub Timer2_Timer()
    UpdateTime = UpdateTime + 1
    StatusBar1.Panels(2).Text = "Download Time:" & Str$(UpdateTime) & " Seconds"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub

Private Sub Timer3_Timer()
If txtTime <> 0 Then
If txtTime = 1 Then
If chkPlay.Value = 1 Then
txtTime = 0
        Load frmLogin
        frmLogin.Show
        Unload Me
End If
Else
txtTime = txtTime - 1
status$ = "Auto Play in " & txtTime & " second(s)."
End If
End If
End Sub
