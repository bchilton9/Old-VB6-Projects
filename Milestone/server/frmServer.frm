VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmServer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Server"
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   9840
   Icon            =   "frmServer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   9840
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Accounts In Use"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   7320
      TabIndex        =   5
      Top             =   1920
      Width           =   2535
      Begin VB.ListBox LstAccounts 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   2220
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Current Online Players"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   7320
      TabIndex        =   3
      Top             =   0
      Width           =   2535
      Begin VB.ListBox LstPlayers 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1380
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   2250
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Active Server Window / Broadcast"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7215
      Begin VB.TextBox txtText 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2850
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   240
         Width           =   6915
      End
      Begin VB.TextBox txtChat 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   3120
         Width           =   6915
      End
   End
   Begin VB.Timer PlayerTimer 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   840
      Top             =   4800
   End
   Begin VB.Timer tmrShutdown 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   1320
      Top             =   4800
   End
   Begin VB.Timer tmrGameAI 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1800
      Top             =   4800
   End
   Begin VB.Timer tmrSpawnMapItems 
      Interval        =   1000
      Left            =   2280
      Top             =   4800
   End
   Begin VB.Timer tmrPlayerSave 
      Interval        =   60000
      Left            =   360
      Top             =   4800
   End
   Begin MSWinsockLib.Winsock Socket 
      Index           =   0
      Left            =   2760
      Top             =   4800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuShutdown 
         Caption         =   "Shutdown"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuDatabase 
      Caption         =   "&Database"
      Begin VB.Menu mnuReloadClasses 
         Caption         =   "Reload Classes"
      End
   End
   Begin VB.Menu mnuLog 
      Caption         =   "&Log"
      Begin VB.Menu mnuServerLog 
         Caption         =   "Server Log"
      End
   End
   Begin VB.Menu mnuScript 
      Caption         =   "&Script"
      Begin VB.Menu mnuReloadScript 
         Caption         =   "Reload"
      End
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim lmsg As Long
    
    lmsg = x / Screen.TwipsPerPixelX
    Select Case lmsg
        Case WM_LBUTTONDBLCLK
            frmServer.WindowState = vbNormal
            frmServer.Show
    End Select
End Sub

Private Sub Form_Resize()
    If frmServer.WindowState = vbMinimized Then
        frmServer.Hide
    End If
End Sub

Private Sub Form_Terminate()
    Call DestroyServer
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call DestroyServer
End Sub

Private Sub mnuReloadScript_Click()
If Scripting = 1 Then
    Set MyScript = Nothing
    Set clsScriptCommands = Nothing
    
    Set MyScript = New clsSadScript
    Set clsScriptCommands = New clsCommands
    MyScript.ReadInCode App.Path & "\Scripts\Main.txt", "Scripts\Main.txt", MyScript.SControl, False
    MyScript.SControl.AddObject "ScriptHardCode", clsScriptCommands, True
    Call TextAdd(frmServer.txtText, "Scripts reloaded.", True)
End If
End Sub

Private Sub mnuServerLog_Click()
    If mnuServerLog.Checked = True Then
        mnuServerLog.Checked = False
        ServerLog = False
    Else
        mnuServerLog.Checked = True
        ServerLog = True
    End If
End Sub


Private Sub PlayerTimer_Timer()
Dim i As Long

If PlayerI <= MAX_PLAYERS Then
If IsPlaying(PlayerI) Then
Call SavePlayer(PlayerI)
Call PlayerMsg(PlayerI, GetPlayerName(PlayerI) & " is now saved.", Yellow)
End If
PlayerI = PlayerI + 1
End If
If PlayerI >= MAX_PLAYERS Then
PlayerI = 1
PlayerTimer.Enabled = False
tmrPlayerSave.Enabled = True
End If


End Sub


Private Sub tmrGameAI_Timer()
    Call ServerLogic
End Sub

Private Sub tmrPlayerSave_Timer()
    Call PlayerSaveTimer
End Sub

Private Sub tmrSpawnMapItems_Timer()
    Call CheckSpawnMapItems
End Sub

Private Sub txtText_GotFocus()
    txtChat.SetFocus
End Sub

Private Sub txtChat_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And Trim(txtChat.Text) <> "" Then
        Call GlobalMsg(txtChat.Text, White)
        Call TextAdd(frmServer.txtText, "Server: " & txtChat.Text, True)
        txtChat.Text = ""
    End If
End Sub

Private Sub tmrShutdown_Timer()
Static Secs As Long

    If Secs <= 0 Then Secs = 30
    Call GlobalMsg("Server Shutdown in " & Secs & " seconds.", BrightBlue)
    Call TextAdd(frmServer.txtText, "Automated Server Shutdown in " & Secs & " seconds.", True)
    Secs = Secs - 2
    If Secs <= 0 Then
        tmrShutdown.Enabled = False
        Call DestroyServer
    End If
End Sub

Private Sub mnuShutdown_Click()
    tmrShutdown.Enabled = True
End Sub

Private Sub mnuExit_Click()
    Call DestroyServer
End Sub

Private Sub mnuReloadClasses_Click()
    Call LoadClasses
    Call TextAdd(frmServer.txtText, "All classes reloaded.", True)
End Sub

Private Sub Socket_ConnectionRequest(index As Integer, ByVal requestID As Long)
    Call AcceptConnection(index, requestID)
End Sub

Private Sub Socket_Accept(index As Integer, SocketId As Integer)
    Call AcceptConnection(index, SocketId)
End Sub

Private Sub Socket_DataArrival(index As Integer, ByVal bytesTotal As Long)
    If IsConnected(index) Then
        Call IncomingData(index, bytesTotal)
    End If
End Sub

Private Sub Socket_Close(index As Integer)
    Call CloseSocket(index)
End Sub


