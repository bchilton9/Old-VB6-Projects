VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBot 
   Appearance      =   0  'Flat
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "MailSys"
   ClientHeight    =   4350
   ClientLeft      =   150
   ClientTop       =   450
   ClientWidth     =   6840
   FillColor       =   &H00808080&
   ForeColor       =   &H00000000&
   Icon            =   "frmBot.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmBot.frx":08CA
   ScaleHeight     =   4350
   ScaleWidth      =   6840
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame6 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Bot Stats"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   855
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   4815
      Begin VB.TextBox usrname 
         Appearance      =   0  'Flat
         BackColor       =   &H00404000&
         ForeColor       =   &H008080FF&
         Height          =   285
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox txtsent 
         Appearance      =   0  'Flat
         BackColor       =   &H00404000&
         ForeColor       =   &H008080FF&
         Height          =   285
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   14
         Text            =   "0"
         ToolTipText     =   "Number of Messages Sent"
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox txterr 
         Appearance      =   0  'Flat
         BackColor       =   &H00404000&
         ForeColor       =   &H008080FF&
         Height          =   285
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "0"
         ToolTipText     =   "Number of Errors"
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox txtmem 
         Appearance      =   0  'Flat
         BackColor       =   &H00404000&
         ForeColor       =   &H008080FF&
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "0"
         ToolTipText     =   "Number of Members"
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox timon 
         Appearance      =   0  'Flat
         BackColor       =   &H00404000&
         ForeColor       =   &H008080FF&
         Height          =   285
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "0:0:0"
         ToolTipText     =   "Connection Time"
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox txtcnt 
         Appearance      =   0  'Flat
         BackColor       =   &H00404000&
         Enabled         =   0   'False
         ForeColor       =   &H008080FF&
         Height          =   285
         Left            =   120
         TabIndex        =   10
         Text            =   "False"
         ToolTipText     =   "Connection Status"
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label6 
         BackColor       =   &H00404040&
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0C0&
         Height          =   255
         Left            =   3960
         TabIndex        =   21
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label5 
         BackColor       =   &H00404040&
         Caption         =   "Sent"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0C0&
         Height          =   255
         Left            =   3120
         TabIndex        =   19
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label4 
         BackColor       =   &H00404040&
         Caption         =   "Errors"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0C0&
         Height          =   255
         Left            =   2280
         TabIndex        =   18
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label3 
         BackColor       =   &H00404040&
         Caption         =   "Users"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0C0&
         Height          =   255
         Left            =   1560
         TabIndex        =   17
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label2 
         BackColor       =   &H00404040&
         Caption         =   "Time"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0C0&
         Height          =   255
         Left            =   840
         TabIndex        =   16
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label1 
         BackColor       =   &H00404040&
         Caption         =   "Online"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0C0&
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Warps 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Map Warps"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   3015
      Left            =   5040
      TabIndex        =   8
      Top             =   120
      Width           =   1695
      Begin VB.Image cmdfurabia 
         Height          =   330
         Left            =   120
         Picture         =   "frmBot.frx":1194
         ToolTipText     =   "Furabia"
         Top             =   2520
         Width           =   1440
      End
      Begin VB.Image cmdmeo 
         Height          =   330
         Left            =   120
         Picture         =   "frmBot.frx":2A96
         ToolTipText     =   "Meovanni"
         Top             =   2160
         Width           =   1440
      End
      Begin VB.Image cmdacro 
         Height          =   330
         Left            =   120
         Picture         =   "frmBot.frx":4398
         ToolTipText     =   "Acropolis"
         Top             =   1800
         Width           =   1440
      End
      Begin VB.Image cmdhaven 
         Height          =   330
         Left            =   120
         Picture         =   "frmBot.frx":5C9A
         ToolTipText     =   "New Haven"
         Top             =   1440
         Width           =   1440
      End
      Begin VB.Image cmdimag 
         Height          =   330
         Left            =   120
         Picture         =   "frmBot.frx":759C
         ToolTipText     =   "Imaginarium"
         Top             =   1080
         Width           =   1440
      End
      Begin VB.Image cmdpuzzplace 
         Height          =   330
         Left            =   120
         Picture         =   "frmBot.frx":8E9E
         ToolTipText     =   "Puzzle Place"
         Top             =   720
         Width           =   1440
      End
      Begin VB.Image cmdGoAlleg 
         Height          =   330
         Left            =   120
         Picture         =   "frmBot.frx":A7A0
         ToolTipText     =   "Allegria"
         Top             =   360
         Width           =   1440
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "System"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   3240
      Width           =   6615
      Begin VB.CheckBox chkcnt 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "AutoConnection"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0C0&
         Height          =   255
         Left            =   2400
         TabIndex        =   7
         ToolTipText     =   "Automatic Connection"
         Top             =   360
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CheckBox prem 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Premote"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0C0&
         Height          =   255
         Left            =   5400
         TabIndex        =   6
         ToolTipText     =   "Premote"
         Top             =   360
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkServtxt 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "S-Text"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0C0&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "Show Server Text"
         Top             =   360
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CheckBox chkWhisp 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Active"
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
         ForeColor       =   &H00FFC0C0&
         Height          =   255
         Left            =   4320
         TabIndex        =   4
         ToolTipText     =   "Active"
         Top             =   360
         Value           =   2  'Grayed
         Width           =   975
      End
      Begin VB.CheckBox chkServCode 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "S-Code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0C0&
         Height          =   255
         Left            =   1200
         TabIndex        =   3
         ToolTipText     =   "Show Server Code"
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Timer StayOnline 
      Interval        =   60000
      Left            =   240
      Top             =   1200
   End
   Begin MSWinsockLib.Winsock sckFurc 
      Left            =   720
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtSend 
      Appearance      =   0  'Flat
      BackColor       =   &H00404000&
      ForeColor       =   &H008080FF&
      Height          =   375
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      ToolTipText     =   "Type here"
      Top             =   2760
      Width           =   4815
   End
   Begin VB.TextBox txtFromFurc 
      Appearance      =   0  'Flat
      BackColor       =   &H00404000&
      ForeColor       =   &H008080FF&
      Height          =   1575
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      ToolTipText     =   "Read here"
      Top             =   1080
      Width           =   4815
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   22
      Top             =   4065
      Width           =   6840
      _ExtentX        =   12065
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Text            =   "Words: 0"
            TextSave        =   "Words: 0"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9446
            Text            =   "Words Typed In Furcadia: 0"
            TextSave        =   "Words Typed In Furcadia: 0"
         EndProperty
      EndProperty
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu Connect 
         Caption         =   "Connect"
         Enabled         =   0   'False
         Shortcut        =   ^K
      End
      Begin VB.Menu Disconnect 
         Caption         =   "Disconnect"
         Shortcut        =   ^D
      End
      Begin VB.Menu Exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu Commands 
      Caption         =   "Commands"
      Begin VB.Menu Movement 
         Caption         =   "Movement"
         Shortcut        =   ^M
      End
      Begin VB.Menu ViewMember 
         Caption         =   "View Member"
         Shortcut        =   ^Y
      End
      Begin VB.Menu EditMember 
         Caption         =   "Edit Member"
         Shortcut        =   ^E
      End
      Begin VB.Menu ViewError 
         Caption         =   "View Error"
         Shortcut        =   ^P
      End
      Begin VB.Menu ClearError 
         Caption         =   "Clear Error"
         Shortcut        =   ^F
      End
      Begin VB.Menu clrfurctxt 
         Caption         =   "Clear Furcadi Text"
         Shortcut        =   ^T
      End
   End
   Begin VB.Menu Help 
      Caption         =   "Help"
      Begin VB.Menu About 
         Caption         =   "About"
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "frmBot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Minute As Integer
Public Hour As Integer
Public Day As Integer
Public urgc As Integer
Public premt As Integer
Public del As Integer
Public Desc As String
Public Connected As Boolean


Private Sub About_Click()
Load frmAbout
frmAbout.Show
End Sub

Private Sub chkServCode_Click()
If chkServCode = 1 Then
chkServtxt = 2
chkServtxt.Enabled = False
End If
If chkServCode = 0 Then
chkServtxt = 1
chkServtxt.Enabled = True
End If
End Sub
Private Sub cmdedit_Click()
box = MsgBox("Not Active Yet!", vbOKOnly, "Edit Members")
'Load Form2
'Form2.Show
End Sub



Private Sub cmdviemem_Click()
Load viewmem
viewmem.Show
End Sub

Private Sub ClearError_Click()
If txterr = "0" Then
box = MsgBox("There are no errors", vbOKOnly, "Clear Text")
Else
box = MsgBox("Are you sure you want to clear all the errors in the file.", vbOKCancel, "Clear Text")
If box = vbOK Then
    Open "C:\mailsys\errorlog.txt" For Output As #6
    Write #6, ""
    Close #6
    Open "C:\mailsys\errorq.txt" For Output As #6
    Write #6, 0
    Close #6
txterr = 0
End If
End If
End Sub

Private Sub clrfurctxt_Click()
box = MsgBox("Are you sure you want to clear all the text from Furcadia.", vbOKCancel, "Clear Text")
If box = vbOK Then txtFromFurc = ""
End Sub

Private Sub cmdacro_Click()
If Connected = True Then sckFurc.SendData "gomap *" & vbLf
End Sub

Private Sub cmdfurabia_Click()
If Connected = True Then sckFurc.SendData "gomap %" & vbLf
End Sub

Private Sub cmdhaven_Click()
If Connected = True Then sckFurc.SendData "gomap '" & vbLf
End Sub

Private Sub cmdimag_Click()
If Connected = True Then sckFurc.SendData "gomap (" & vbLf
End Sub

Private Sub cmdmeo_Click()
If Connected = True Then sckFurc.SendData "gomap &" & vbLf
End Sub

Private Sub cmdpuzzplace_Click()
If Connected = True Then sckFurc.SendData "gomap $" & vbLf
End Sub





Private Sub Connect_Click()
Load frmConnect
frmConnect.Show
frmBot.Hide
End Sub

Private Sub Disconnect_Click()
If Connected = True Then
sckFurc.Close
Connected = False
txtcnt = "False"
End If
Connect.Enabled = True
Disconnect.Enabled = False
End Sub

Private Sub Image1_Click()

End Sub

Private Sub Image2_Click()

End Sub

Private Sub Image4_Click()

End Sub

Private Sub EditMember_Click()
box = MsgBox("Not Active Yet!", vbOKOnly, "Edit Members")
'Load Form2
'Form2.Show
End Sub


Private Sub Exit_Click()
End
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub


Private Sub Movement_Click()
If Movement.Checked = False Then
Movement.Checked = True
frmMovement.Visible = True
Else
Movement.Checked = False
frmMovement.Visible = False
End If
End Sub

Private Sub prem_Click()
If prem = 1 Then
chkWhisp = 2
chkWhisp.Enabled = False
End If
If prem = 0 Then
chkWhisp = 1
chkWhisp.Enabled = True
End If
End Sub


Private Sub cmdGoAlleg_Click()
If Connected = True Then sckFurc.SendData "goalleg" & vbLf
End Sub
Private Sub cmdGoVinca_Click()
If Connected = True Then sckFurc.SendData "gostart" & vbLf
End Sub

Private Sub cmdwho_Click()
If Connected = True Then sckFurc.SendData "who" & vbLf
End Sub
Sub Form_Load()
    Open "C:\mailsys\memnum.txt" For Input As #1
    Input #1, nnum
    Close #1
    Open "C:\mailsys\sent.txt" For Input As #3
    Input #3, sent
    Close #3
    Open "C:\mailsys\errorq.txt" For Input As #3
    Input #3, errq
    Close #3
    txterr = errq
    txtsent = sent
    txtmem = nnum
Minute = 0
urgc = 0
Hour = 0
premt = 0
Desc = frmConnect.descfield & " [Uptime: 0 Minute(s)]"
End Sub


Private Sub sckFurc_DataArrival(ByVal bytesTotal As Long)
Dim s As String
sckFurc.GetData s
x = Split(s, vbLf)
For r = 0 To UBound(x) - 1
RealText x(r)
Next
End Sub

Sub RealText(Txt)
If chkServtxt.Value = Checked Or chkServtxt.Enabled = False Then
If chkServCode.Value = Checked Then txtFromFurc = txtFromFurc & Txt & vbCrLf
If chkServCode.Value = Unchecked Then
If Left(Txt, 1) = "(" Then txtFromFurc = txtFromFurc & Right(Txt, Len(Txt) - 1) & vbCrLf
End If
End If
If Txt = "END" Then
sckFurc.SendData "connect " & frmConnect.namefield & " " & frmConnect.passwordfield & vbLf & "color " & frmConnect.colorfield & vbLf & "desc " & frmConnect.descfield & vbLf
Connected = "True"
txtcnt = "True"
'sckFurc.SendData "goalleg" & vbLf
End If
If Txt = "]ccmarbled.pcx" Then
    sckFurc.SendData "vascodagama" & vbLf
End If

If Left(Txt, 15) Like "(Server going d" Then
    txterr = txterr + 1
    Open "C:\mailsys\errorq.txt" For Input As #5
    Input #5, num
    Close #5
    num = num + 1
    Open "C:\mailsys\errorq.txt" For Output As #5
    Write #5, num
    Close #5
    Open "C:\mailsys\errorlog.txt" For Append As #5
    Write #5, "Server went down [When: " & Date & " at " & Time & "]"
    Close #5
    sckFurc.Close
    Connected = False
    txtcnt = "False"
End If
If Left(Txt, 15) Like "(Someone else h" Then
    txterr = txterr + 1
    Open "C:\mailsys\errorq.txt" For Input As #5
    Input #5, num
    Close #5
    num = num + 1
    Open "C:\mailsys\errorq.txt" For Output As #5
    Write #5, num
    Close #5
    Open "C:\mailsys\errorlog.txt" For Append As #5
    Write #5, "Someone loged in as mailsys [When: " & Date & " at " & Time & "]"
    Close #5
    sckFurc.Close
    Connected = False
    txtcnt = "False"
End If
If Left(Txt, 15) Like "(Disconnected f" Then
    txterr = txterr + 1
    Open "C:\mailsys\errorq.txt" For Input As #5
    Input #5, num
    Close #5
    num = num + 1
    Open "C:\mailsys\errorq.txt" For Output As #5
    Write #5, num
    Close #5
    Open "C:\mailsys\errorlog.txt" For Append As #5
    Write #5, "Disconted for 30 mins [When: " & Date & " at " & Time & "]"
    Close #5
    sckFurc.Close
    Connected = False
    txtcnt = "False"
End If

If chkWhisp.Value = Checked Or chkWhisp.Enabled = False Then
If Left(Txt, 3) = "([ " And Right(Txt, 10) = " to you. ]" Then
    tmsg = Split(Txt, " whispers, " & Chr(34), 2)
    Furre = Right(tmsg(0), Len(tmsg(0)) - 3)
    NMsg = Left(tmsg(1), Len(tmsg(1)) - 11)
    Msg = LCase(NMsg)
    Furre = LCase(Furre)
    If Left(Msg, 5) = "read " Then
        On Error GoTo error
        snd = Right(Msg, Len(Msg) - 5)
        remsg Furre, snd, Txt
    Else
    If Left(Msg, 7) = "delete " Then
        On Error GoTo error
        snd = Right(Msg, Len(Msg) - 7)
        dodelete Furre, snd, Txt
    Else
    If Left(Msg, 5) = "send " Then
        On Error GoTo error
        aMsg = Split(Msg, " message ", 2)
        snd = Right(aMsg(0), Len(aMsg(0)) - 5)
        snd = Replace(snd, " ", "|")
        mssg = Left(aMsg(1), Len(aMsg(1)) - 0)
        sndmsg Furre, snd, mssg, Txt
    Else
    If Left(Msg, 6) = "check " Then
        On Error GoTo error
        snd = Right(Msg, Len(Msg) - 6)
        snd = Replace(snd, " ", "|")
        chkfur Furre, snd
    Else
    If Left(Msg, 8) = "suggest " Then
        On Error GoTo error
        snd = Right(Msg, Len(Msg) - 8)
        snd = Replace(snd, " ", "|")
        sugfur Furre, snd, Txt
    Else
    If Left(Msg, 5) = "card " Then
        On Error GoTo error
        aMsg = Split(Msg, " message ", 2)
        bMsg = Right(aMsg(0), Len(aMsg(0)) - 5)
        cMsg = Split(bMsg, " image ", 2)
        snd = Right(cMsg(0), Len(cMsg(0)) - 0)
        snd = Replace(snd, " ", "|")
        imag = Left(cMsg(1), Len(cMsg(1)) - 0)
        mssg = Left(aMsg(1), Len(aMsg(1)) - 0)
        mssg = Replace(mssg, " ", "%20")
        sndcard Furre, snd, mssg, imag, Txt
    Else
        DoWhisper Furre, Msg, Txt
    End If
    End If
    End If
    End If
    End If
    End If
End If
Else
    If Left(Txt, 3) = "([ " And Right(Txt, 10) = " to you. ]" Then
    tmsg = Split(Txt, " whispers, " & Chr(34), 2)
    Furre = Right(tmsg(0), Len(tmsg(0)) - 3)
    sckFurc.SendData "wh " & Furre & " Im currently offline please try again later." & vbLf
    End If
End If 'chkWhisp
Exit Sub
error:
    sckFurc.SendData "wh " & Furre & " Im Sorry,        " & Chr(34) & Msg & Chr(34) & " is not a valid command. Whisper me *help to learn how to use my service." & vbLf
    Resume stoptrying
stoptrying:
End Sub
Sub sndcard(Furre, snd, mssg, imag, Txt)

Open "C:\mailsys\members.txt" For Input As #1
Do Until (fName = Furre) Or (EOF(1))
Input #1, fName, mnum
Loop
Close #1
If fName = Furre Then
On Error GoTo snerr
Open "C:\mailsys\members.txt" For Input As #1
Do Until (sfName = snd) Or (EOF(1))
Input #1, sfName, smnum
Loop
Close #1
If sfName = snd Then
Dim qun As String
    Open "C:\mailsys\memfiles\" & smnum & "q.txt" For Input As #3
    Input #3, qun
    Close #3
    qun = qun + 1
    Open "C:\mailsys\memfiles\" & smnum & "q.txt" For Output As #3
    Write #3, qun
    Close #3
    
    Open "C:\mailsys\sent.txt" For Input As #3
    Input #3, sent
    Close #3
    sent = sent + 1
    Open "C:\mailsys\sent.txt" For Output As #3
    Write #3, sent
    Close #3
    
    Open "C:\mailsys\memfiles\" & smnum & ".txt" For Append As #1
    Write #1, qun, "SysCard", "http://www.erenetwork.com/mailsys/images/card.cgi?mode=1&from=" & Furre & "&image=" & imag & "&mess=" & mssg & "                                                                  You have a SysCard wateing for you. Press F8 now to view.  " & " [Sent: " & Date & " at " & Time & " MST]"
    Close #1
    sckFurc.SendData "wh " & Furre & " Message sent to: " & snd & vbLf
Else
    sckFurc.SendData "wh " & Furre & " " & snd & " is not registered." & vbLf
End If
Else
sckFurc.SendData "wh " & Furre & " You are not registered with me. Whisper me *help to learn how to use my service." & vbLf
End If
Exit Sub
snerr:
    Close #1, #3
    txterr = txterr + 1
    Open "C:\mailsys\errorq.txt" For Input As #5
    Input #5, num
    Close #5
    num = num + 1
    Open "C:\mailsys\errorq.txt" For Output As #5
    Write #5, num
    Close #5
    Open "C:\mailsys\errorlog.txt" For Append As #5
    Write #5, "Card error (sndcard), from " & Furre & " with " & Txt & ", [When: " & Date & " at " & Time & "]"
    Close #5
    sckFurc.SendData "wh " & Furre & " There has been an error. Whisper me *help to learn how to use my service. If you keep recieveing this error please leave a message for Mys' or e-mail him at mailsys@erenetwork.com." & vbLf
    Resume stoptrying
stoptrying:
End Sub
Sub sugfur(Furre, snd, Txt)
On Error GoTo sugfurerr
Open "C:\mailsys\members.txt" For Input As #1
Do Until (sfName = snd) Or (EOF(1))
Input #1, sfName, smnum
Loop
Close #1
If snd = "" Then
sckFurc.SendData "wh " & Furre & " I'm sorry, You must enter a Furry's name. Whisper me *help to learn how to use my service." & vbLf
Else
If sfName = snd Then
sckFurc.SendData "wh " & Furre & " " & snd & " is allready a member." & vbLf

Else
Open "C:\mailsys\suggest.txt" For Append As #1
Write #1, snd, Furre
Close #1
        Open "C:\mailsys\suggestq.txt" For Input As #3
        Input #3, qunt
        Close #3
        qunt = qunt + 1
        Open "C:\mailsys\suggestq.txt" For Output As #3
        Write #3, qunt
        Close #3

sckFurc.SendData "wh " & Furre & " I will let " & snd & " know that you suggested that he/she join's Mailsys. Thank you." & vbLf
End If
End If

Exit Sub
sugfurerr:
    Close #1, #3
    txterr = txterr + 1
    Open "C:\mailsys\errorq.txt" For Input As #5
    Input #5, num
    Close #5
    num = num + 1
    Open "C:\mailsys\errorq.txt" For Output As #5
    Write #5, num
    Close #5
    Open "C:\mailsys\errorlog.txt" For Append As #5
    Write #5, "Suggest Error (sugfur), from " & Furre & " with " & Txt & ", [When: " & Date & " at " & Time & "]"
    Close #5
    sckFurc.SendData "wh " & Furre & " There has been an error. Whisper me *help to learn how to use my service. If you keep recieveing this error please leave a message for Mys' or e-mail him at mailsys@erenetwork.com." & vbLf
    Resume stoptrying
stoptrying:
End Sub

Sub chkfur(Furre, snd)
Open "C:\mailsys\members.txt" For Input As #1
Do Until (sfName = snd) Or (EOF(1))
Input #1, sfName, smnum
Loop
Close #1
If sfName = snd Then
    sckFurc.SendData "wh " & Furre & " " & snd & " is registered." & vbLf
Else
    sckFurc.SendData "wh " & Furre & " " & snd & " is not registered." & vbLf
End If

End Sub
Sub remsg(Furre, sndr, Txt)
Open "C:\mailsys\members.txt" For Input As #1
Do Until (fName = Furre) Or (EOF(1))
Input #1, fName, mnum
Loop
Close #1
If fName = Furre Then
On Error GoTo reerr
Open "C:\mailsys\memfiles\" & mnum & ".txt" For Input As #1
Do Until (dnum = sndr) Or (EOF(1))
Input #1, dnum, ser, mssg
Loop
Close #1
If dnum = sndr Then
    sckFurc.SendData "wh " & Furre & " Message from: " & ser & " - " & mssg & vbLf
Else
    sckFurc.SendData "wh " & Furre & "  Im Sorry,        " & Chr(34) & sndr & Chr(34) & " is not a valid message number. Whisper me *help to learn how to use my service." & vbLf
End If

Else
    sckFurc.SendData "wh " & Furre & " You are not registered with me." & vbLf
End If
Exit Sub
reerr:
    Close #1
    txterr = txterr + 1
    Open "C:\mailsys\errorq.txt" For Input As #5
    Input #5, num
    Close #5
    num = num + 1
    Open "C:\mailsys\errorq.txt" For Output As #5
    Write #5, num
    Close #5
    Open "C:\mailsys\errorlog.txt" For Append As #5
    Write #5, "Read Message error (remsg), from " & Furre & " with " & Txt & ", [When: " & Date & " at " & Time & "]"
    Close #5
    sckFurc.SendData "wh " & Furre & " There has been an error. Whisper me *help to learn how to use my service. If you keep recieveing this error please leave a message for Mys' or e-mail him at mailsys@erenetwork.com." & vbLf
    Resume stoptrying
stoptrying:
End Sub

Sub doread(Furre, Txt)
Open "C:\mailsys\members.txt" For Input As #2
Do Until (fName = Furre) Or (EOF(2))
Input #2, fName, mnum
Loop
Close #2
If fName = Furre Then
On Error GoTo reerr
        Open "C:\mailsys\memfiles\" & mnum & "q.txt" For Input As #3
        Input #3, qunt
        Close #3
        If qunt = 0 Then
            sckFurc.SendData "wh " & Furre & " You dont have any messages." & vbLf
        
        Else
            On Error GoTo reerr
            sckFurc.SendData "wh " & Furre & " You have messages from:"
            Open "C:\mailsys\memfiles\" & mnum & ".txt" For Input As #1
            Do Until (EOF(1))
            Input #1, dnum, ser, mssg
            sckFurc.SendData " [" & ser & " - #" & dnum & "]"
            Loop
            Close #1
            sckFurc.SendData vbLf
        End If 'end if qount = 0
Else
    sckFurc.SendData "wh " & Furre & " You are not registered with me. Whisper me *help to learn how to use my service." & vbLf
End If

Exit Sub
reerr:
    Close #1, #2, #3
    txterr = txterr + 1
    Open "C:\mailsys\errorq.txt" For Input As #5
    Input #5, num
    Close #5
    num = num + 1
    Open "C:\mailsys\errorq.txt" For Output As #5
    Write #5, num
    Close #5
    Open "C:\mailsys\errorlog.txt" For Append As #5
    Write #5, "Read error (doread), from " & Furre & " with " & Txt & ", [When: " & Date & " at " & Time & "]"
    Close #5
    sckFurc.SendData "wh " & Furre & " There has been an error. Whisper me *help to learn how to use my service. If you keep recieveing this error please leave a message for Mys' or e-mail him at mailsys@erenetwork.com." & vbLf
    Resume stoptrying
stoptrying:
End Sub

Sub DoWhisper(Furre, Msg, Txt)
If Msg Like "*help" Then
    sckFurc.SendData "wh " & Furre & " Thank you for choosing Mailsys as your Furcadian mail service. Whisper me the following to learn how I run: *mail     *entertainment     *other" & vbLf

ElseIf Msg Like "albino" Then
albino
sckFurc.SendData "wh " & Furre & " " & usrname & vbLf

ElseIf Msg Like "alver" Then
sckFurc.SendData "wh " & Furre & " " & usrname & vbLf

ElseIf Msg Like "deverry" Then
Deverry
sckFurc.SendData "wh " & Furre & " " & usrname.Text & vbLf

ElseIf Msg Like "elf" Then
elf
sckFurc.SendData "wh " & Furre & " " & usrname.Text & vbLf

ElseIf Msg Like "felanna" Then
felana
sckFurc.SendData "wh " & Furre & " " & usrname.Text & vbLf

ElseIf Msg Like "galler" Then
galler
sckFurc.SendData "wh " & Furre & " " & usrname.Text & vbLf

ElseIf Msg Like "orc" Then
orc
sckFurc.SendData "wh " & Furre & " " & usrname.Text & vbLf

ElseIf Msg Like "join" Then
    dojoin Furre, Txt
ElseIf Msg Like "read" Then
    doread Furre, Txt
ElseIf Msg Like "stats" Then
    Open "C:\mailsys\memnum.txt" For Input As #1
    Input #1, memn
    Close #1
    Open "C:\mailsys\sent.txt" For Input As #3
    Input #3, sent
    Close #3
    sckFurc.SendData "wh " & Furre & " My manager is Mys'. The softwhere I use to run my services was made with VB6. It is curently version " & vers & ", I have been on the clock for " & Day & " Day(s) " & Hour & " Hour(s) " & Minute & " Minute(s). I have " & memn & " Furre's useing my services. I have delivered " & sent & " messages [NOT includeing the welcome message's sent at sign up.]. I am an official bot of the AICC. [http://AICC.erenetwork.com]" & vbLf
ElseIf Msg Like "news" Then
    Open "C:\mailsys\news.txt" For Input As #1
    Input #1, news
    Close #1
    sckFurc.SendData "wh " & Furre & " " & news & vbLf
ElseIf Msg Like "*other" Then
sckFurc.SendData "wh " & Furre & " Here are the commands I currently have in this section: *news     *stats" & vbLf

ElseIf Msg Like "*stats" Then
sckFurc.SendData "wh " & Furre & " Whisper me STATS and I will tell you lots about me." & vbLf

ElseIf Msg Like "*news" Then
sckFurc.SendData "wh " & Furre & " Whisper me NEWS and I will tell you latest news." & vbLf

ElseIf Msg Like "*mail" Then
sckFurc.SendData "wh " & Furre & " The following commands will help you use my mail service: #join, #read, #delete, #send, #check, #suggest, #card" & vbLf

ElseIf Msg Like "*join" Then
sckFurc.SendData "wh " & Furre & " If you wish to use my messaging service you must first sign up. It is easy. Simply type '/Mailsys JOIN' and I will create your account." & vbLf

ElseIf Msg Like "*read" Then
sckFurc.SendData "wh " & Furre & " To check your messages type '/Mailsys Read' and I will give you a list of all who have left you a message followed by a message number. Then type '/Mailsys Read #' to read the message from the Furry. Replacing the # with the message number." & vbLf

ElseIf Msg Like "*delete" Then
sckFurc.SendData "wh " & Furre & " To delete the first message on your list from some one type '/Mailsys Delete #' and this will remove the message from your list." & vbLf

ElseIf Msg Like "*send" Then
sckFurc.SendData "wh " & Furre & " To send a message type '/Mailsys SEND Furrename MESSAGE Messagebody' replacing Furrename with the name of the person you wish to mail, and messagebody with what you want the message to say." & vbLf

ElseIf Msg Like "*check" Then
sckFurc.SendData "wh " & Furre & " To see if a certain furry is registered with my serivice, type '/Mailsys CHECK furrename' replaceing furrename with the name of the person you wish to check on." & vbLf

ElseIf Msg Like "*suggest" Then
sckFurc.SendData "wh " & Furre & " Type '/Mailsys CHECK Furrename' Replacing Furrename with the Furry's name and I will check every 30 minutes for them to be on and send them a message. " & vbLf

ElseIf Msg Like "*card" Then
sckFurc.SendData "wh " & Furre & " To send a greeting card to another furry type '/Mailsys CARD Furrename IMAGE # MESSAGE MM' replacing Furrename with the furry's name, # with the number of the image, and MM with the message. For Example '/Mailsys CARD Felorin IMAGE 3 Furcadia is great!'. To view the available images please go here: http://www.erenetwork.com/mailsys/images/index.html" & vbLf

ElseIf Msg Like "insult" Then

Dim l0038 As Variant
  Dim l003C As Variant
  Dim l0040 As Variant
  Dim l0044 As Variant
  Dim l0048 As Variant
  Dim Insults As Variant
  ReDim l001A(1 To 142)
  ReDim l0020(1 To 32)
  ReDim l0026(1 To 41)
  ReDim l002C(1 To 18)
  ReDim l0032(1 To 41)
  Let l001A(1) = "stupid"
  Let l001A(2) = "annoying"
  Let l001A(3) = "numb"
  Let l001A(4) = "fat"
  Let l001A(5) = "yellow  "
  Let l001A(6) = "revolting"
  Let l001A(7) = "sickening"
  Let l001A(8) = "disgusting"
  Let l001A(9) = "perverted"
  Let l001A(10) = "stupid"
  Let l001A(11) = "illiterate"
  Let l001A(12) = "flea-bitten"
  Let l001A(13) = "depraved"
  Let l001A(14) = "uncouth"
  Let l001A(15) = "bad breathed"
  Let l001A(16) = "pitiful"
  Let l001A(17) = "dumpy"
  Let l001A(18) = "offensive"
  Let l001A(19) = "dim witted"
  Let l001A(20) = "loathsome"
  Let l001A(21) = "insignificant"
  Let l001A(22) = "blithering"
  Let l001A(23) = "repulsive"
  Let l001A(24) = "worthless"
  Let l001A(25) = "blundering"
  Let l001A(26) = "retarded"
  Let l001A(27) = "useless"
  Let l001A(28) = "obnoxious"
  Let l001A(29) = "low budget"
  Let l001A(30) = "asisine"
  Let l001A(31) = "neurotic"
  Let l001A(32) = "subhuman"
  Let l001A(33) = "crochety"
  Let l001A(34) = "indescribable"
  Let l001A(35) = "contemptible"
  Let l001A(36) = "unspeakable"
  Let l001A(37) = "sick"
  Let l001A(38) = "lazy"
  Let l001A(39) = "good for nothing"
  Let l001A(40) = "slutty"
  Let l001A(41) = "spastic"
  Let l001A(42) = "creepy"
  Let l001A(43) = "sloppy"
  Let l001A(44) = "dumb"
  Let l001A(45) = "predictable"
  Let l001A(46) = "atrocious"
  Let l001A(47) = "grotesque"
  Let l001A(48) = "ugly"
  Let l001A(49) = "ungodly"
  Let l001A(50) = "feeble-minded"
  Let l001A(51) = "clueless"
  Let l001A(52) = "demented"
  Let l001A(53) = "bewildered"
  Let l001A(54) = "outrageous"
  Let l001A(55) = "deranged"
  Let l001A(56) = "confused"
  Let l001A(57) = "miserable"
  Let l001A(58) = "detestable"
  Let l001A(59) = "annoying"
  Let l001A(60) = "shameless"
  Let l001A(61) = "ignorant"
  Let l001A(62) = "despicable"
  Let l001A(63) = "insane"
  Let l001A(64) = "sleazy"
  Let l001A(65) = "tiny brained"
  Let l001A(66) = "oblivious"
  Let l001A(67) = "hopeless"
  Let l001A(68) = "god-awful"
  Let l001A(69) = "bungling"
  Let l001A(70) = "appalling"
  Let l001A(71) = "skaggy"
  Let l001A(72) = "brainless"
  Let l001A(73) = "boring"
  Let l001A(74) = "uncultivated"
  Let l001A(75) = "inadequate"
  Let l001A(76) = "inhuman"
  Let l001A(77) = "self-exalting"
  Let l001A(78) = "testy"
  Let l001A(79) = "irresponsible"
  Let l001A(80) = "mentally deficient"
  Let l001A(81) = "disdainful"
  Let l001A(82) = "friendless"
  Let l001A(83) = "dreadfull"
  Let l001A(84) = "dorky"
  Let l001A(85) = "psychotic"
  Let l001A(86) = "opinionated"
  Let l001A(87) = "monotonous"
  Let l001A(88) = "disgraceful"
  Let l001A(89) = "preposterous"
  Let l001A(90) = "tacky"
  Let l001A(91) = "uneducated"
  Let l001A(92) = "rediculous"
  Let l001A(93) = "double ugly"
  Let l001A(94) = "irrational,cranky"
  Let l001A(95) = "goofy"
  Let l001A(96) = "crude"
  Let l001A(97) = "embarrassing"
  Let l001A(98) = "deeply disturbed"
  Let l001A(99) = "inept"
  Let l001A(100) = "undisciplined"
  Let l001A(101) = "crooked"
  Let l001A(102) = "pathetic"
  Let l001A(103) = "infantile"
  Let l001A(104) = "witless"
  Let l001A(105) = "indecent"
  Let l001A(106) = "infuriating"
  Let l001A(107) = "unimpressive"
  Let l001A(108) = "insufferable"
  Let l001A(109) = "dismal"
  Let l001A(110) = "erratic"
  Let l001A(111) = "incapable"
  Let l001A(112) = "hallucinating"
  Let l001A(113) = "pompous"
  Let l001A(114) = "pitiable"
  Let l001A(115) = "slovenly"
  Let l001A(116) = "laughable"
  Let l001A(117) = "bad tempered"
  Let l001A(118) = "decrepit"
  Let l001A(119) = "bizarre y  driveling"
  Let l001A(120) = "uncultured"
  Let l001A(121) = "cantankerous"
  Let l001A(122) = "hypocritical"
  Let l001A(123) = "foul"
  Let l001A(124) = "raunchy"
  Let l001A(125) = "putrid"
  Let l001A(126) = "filthy"
  Let l001A(127) = "idiotic"
  Let l001A(128) = "short"
  Let l001A(129) = "daft"
  Let l001A(130) = "silly"
  Let l001A(131) = "simple"
  Let l001A(132) = "hairy"
  Let l001A(133) = "overfed"
  Let l001A(134) = "worthless"
  Let l001A(135) = "childish"
  Let l001A(136) = "unwanted"
  Let l001A(137) = "stunted"
  Let l001A(138) = "antisocial"
  Let l001A(139) = "greedy"
  Let l001A(140) = "hairless"
  Let l001A(141) = "horny"
  Let l001A(142) = "pigish"
  Let l0020(1) = "toilet-full of"
  Let l0020(2) = "lump of"
  Let l0020(3) = "clump of"
  Let l0020(4) = "barrel of"
  Let l0020(5) = "box of"
  Let l0020(6) = "mound of"
  Let l0020(7) = "crock of"
  Let l0020(8) = "barrel full of"
  Let l0020(9) = "stack of"
  Let l0020(10) = "load of"
  Let l0020(11) = "eruption of"
  Let l0020(12) = "glob of"
  Let l0020(13) = "blob of"
  Let l0020(14) = "bag of"
  Let l0020(15) = "pile of"
  Let l0020(16) = "container of"
  Let l0020(17) = "cake of"
  Let l0020(18) = "bunch of"
  Let l0020(19) = "sack of"
  Let l0020(20) = "shovel-full of"
  Let l0020(21) = "bowl of"
  Let l0020(22) = "wheelbarrel full of"
  Let l0020(23) = "heap of"
  Let l0020(24) = "mountain of"
  Let l0020(25) = "ball of"
  Let l0020(26) = "mass of"
  Let l0020(27) = "truckload of"
  Let l0020(28) = "vat of"
  Let l0020(29) = "loaf of"
  Let l0020(30) = "collection of"
  Let l0020(31) = "piece of"
  Let l0020(32) = "crate of"
  Let l0026(1) = "flithy"
  Let l0026(2) = "moldy"
  Let l0026(3) = "nasty"
  Let l0026(4) = "ugly"
  Let l0026(5) = "rotting"
  Let l0026(6) = "old"
  Let l0026(7) = "crumby"
  Let l0026(8) = "musty"
  Let l0026(9) = "second-hand"
  Let l0026(10) = "fly-covered"
  Let l0026(11) = "moth-eaten"
  Let l0026(12) = "pureed"
  Let l0026(13) = "mildewed"
  Let l0026(14) = "foul"
  Let l0026(15) = "raunchy"
  Let l0026(16) = "putrid"
  Let l0026(17) = "filthy"
  Let l0026(18) = "revolting"
  Let l0026(19) = "sickening"
  Let l0026(20) = "disgusting"
  Let l0026(21) = "worthless"
  Let l0026(22) = "unwanted"
  Let l0026(23) = "putrid"
  Let l0026(24) = "synthetic"
  Let l0026(25) = "crusty"
  Let l0026(26) = "smelly"
  Let l0026(27) = "dirty"
  Let l0026(28) = "musty"
  Let l0026(29) = "septic"
  Let l0026(30) = "imitation"
  Let l0026(31) = "fertilized"
  Let l0026(32) = "steaming"
  Let l0026(33) = "sizzling"
  Let l0026(34) = "gross"
  Let l0026(35) = "recycled"
  Let l0026(36) = "reasty"
  Let l0026(37) = "spastic"
  Let l0026(38) = "creepy"
  Let l0026(39) = "sloppy"
  Let l0026(40) = "hairless"
  Let l0026(41) = "unwanted"
  Let l002C(1) = "cow"
  Let l002C(2) = "rat"
  Let l002C(3) = "pig"
  Let l002C(4) = "hog"
  Let l002C(5) = "raccoon"
  Let l002C(6) = "rabbit"
  Let l002C(7) = "llama"
  Let l002C(8) = "monkey"
  Let l002C(9) = "fox"
  Let l002C(10) = "dog"
  Let l002C(11) = "bird"
  Let l002C(12) = "squid"
  Let l002C(13) = "whale"
  Let l002C(14) = "skunk"
  Let l002C(15) = "camel"
  Let l002C(16) = "maggot"
  Let l002C(17) = "goat"
  Let l002C(18) = "coyote"
  Let l0032(1) = "goat droppings"
  Let l0032(2) = "pimple puss"
  Let l0032(3) = "pig spit"
  Let l0032(4) = "coyote crap"
  Let l0032(5) = "squid waste"
  Let l0032(6) = "monkey fleas"
  Let l0032(7) = "maggot boogers"
  Let l0032(8) = "toe jam"
  Let l0032(9) = "monkey guts"
  Let l0032(10) = "rabbit meat"
  Let l0032(11) = "ear wax"
  Let l0032(12) = "nose nuggets"
  Let l0032(13) = "swine remains"
  Let l0032(14) = "rubbish"
  Let l0032(15) = "monkey carcusses"
  Let l0032(16) = "radioactive sewage"
  Let l0032(17) = "zit cheese"
  Let l0032(18) = "vulture gizzards"
  Let l0032(19) = "hogwash"
  Let l0032(20) = "fox puke"
  Let l0032(21) = "hog vomit"
  Let l0032(22) = "cow phlegm"
  Let l0032(23) = "buffalo chips"
  Let l0032(24) = "weasel warts"
  Let l0032(25) = "dog snot"
  Let l0032(26) = "swamp mud"
  Let l0032(27) = "tripe"
  Let l0032(28) = "fish lips"
  Let l0032(29) = "cockroaches"
  Let l0032(30) = "dandruff flakes"
  Let l0032(31) = "stale carrion"
  Let l0032(32) = "horse puckies"
  Let l0032(33) = "slop"
  Let l0032(34) = "seepage"
  Let l0032(35) = "swill"
  Let l0032(36) = "lumps"
  Let l0032(37) = "cow pies"
  Let l0032(38) = "maggot fodder"
  Let l0032(39) = "remains"
  Let l0032(40) = "lizard bums"
  Let l0032(41) = "drainage"
  Randomize
  Let l0038 = Int(142 * Rnd(1) + 1)
  Let l003C = Int(32 * Rnd(1) + 1)
  Let l0040 = Int(41 * Rnd(1) + 1)
  Let l0044 = Int(18 * Rnd(1) + 1)
  Let l0048 = Int(41 * Rnd(1) + 1)
  Insults = " You " + l001A(l0038) + " " + l0020(l003C) + " " + l0032(l0048) + "!"

sckFurc.SendData "wh " & Furre & Insults & vbLf

ElseIf Msg Like "*insult" Then
sckFurc.SendData "wh " & Furre & " I can randomly generate Insults. If you would like to see one just type '/Mailsys insult'" & vbLf

ElseIf Msg Like "*namegen" Then
sckFurc.SendData "wh " & Furre & " Mystic Name Giver has now been brought to furcadia. Just whisper me the following styles and I will generate you an authentic name: Albino, Alver, Deverry, Elf, Felana, Galler, Orc" & vbLf

ElseIf Msg Like "*entertainment" Then
sckFurc.SendData "wh " & Furre & " This section is still not complete. Here are the commands I have so far: *insult     *namegen" & vbLf

Else
    sckFurc.SendData "wh " & Furre & " I dont understand. Whisper me *help to learn how to use my service." & vbLf
End If
whspnum = 0
End Sub

Sub dodelete(Furre, snd, Txt)
Dim dnum As String
Open "C:\mailsys\members.txt" For Input As #1
Do Until (fName = Furre) Or (EOF(1))
Input #1, fName, mnum
Loop
Close #1
If fName = Furre Then
On Error GoTo deerr
Open "C:\mailsys\memfiles\" & mnum & "q.txt" For Input As #1
Input #1, qun
Close #1

If qun = "0" Then
    sckFurc.SendData "wh " & Furre & " No messages to delete." & vbLf
Else
del = 0
Open "C:\mailsys\memfiles\" & mnum & ".txt" For Input As #1
Open "C:\mailsys\memfiles\" & mnum & "a.txt" For Output As #2
Do Until EOF(1)
Input #1, dnum, frm, mssg
If dnum = snd And del = 0 Then
        sckFurc.SendData "wh " & Furre & " Message #" & snd & " has been deleted." & vbLf
        nqun = qun - 1
        Open "C:\mailsys\memfiles\" & mnum & "q.txt" For Output As #3
        Write #3, nqun
        Close #3
        del = 1
Else
    If del = 1 Then dnum = dnum - 1
    Write #2, dnum, frm, mssg
End If
Loop
Close #1, #2

Open "C:\mailsys\memfiles\" & mnum & "a.txt" For Input As #1
Open "C:\mailsys\memfiles\" & mnum & ".txt" For Output As #2
Do Until EOF(1)
    Input #1, dnum, frm, mssg
    Write #2, dnum, frm, mssg
Loop
Close #1, #2
End If
Open "C:\mailsys\memfiles\" & mnum & "a.txt" For Output As #1
Write #1, ""
Close #1
If del = 0 Then sckFurc.SendData "wh " & Furre & "  Im Sorry,        " & Chr(34) & snd & Chr(34) & " is not a valid message number. Whisper me *help to learn how to use my service." & vbLf
Else
sckFurc.SendData "wh " & Furre & " You are not registered with me. Whisper me *help to learn how to use my service." & vbLf
End If

Exit Sub
deerr:
    Close #1, #2
    txterr = txterr + 1
    Open "C:\mailsys\errorq.txt" For Input As #5
    Input #5, num
    Close #5
    num = num + 1
    Open "C:\mailsys\errorq.txt" For Output As #5
    Write #5, num
    Close #5
    Open "C:\mailsys\errorlog.txt" For Append As #5
    Write #5, "Delete error (dodelete), from " & Furre & " with " & Txt & ", [When: " & Date & " at " & Time & "]"
    Close #5
    sckFurc.SendData "wh " & Furre & " There has been an error. Whisper me *help to learn how to use my service. If you keep recieveing this error please leave a message for Mys' or e-mail him at mailsys@erenetwork.com." & vbLf
    Resume stoptrying
stoptrying:
End Sub

Sub sndmsg(Furre, snd, mssg, Txt)
Open "C:\mailsys\members.txt" For Input As #1
Do Until (fName = Furre) Or (EOF(1))
Input #1, fName, mnum
Loop
Close #1
If fName = Furre Then
On Error GoTo snerr
Open "C:\mailsys\members.txt" For Input As #1
Input #1, sfName, smnum
Do Until (sfName = snd) Or (EOF(1))
Input #1, sfName, smnum
Loop
Close #1
If sfName = snd Then
Dim qun As String
    Open "C:\mailsys\memfiles\" & smnum & "q.txt" For Input As #3
    Input #3, qun
    Close #3
    qun = qun + 1
    Open "C:\mailsys\memfiles\" & smnum & "q.txt" For Output As #3
    Write #3, qun
    Close #3
    
    Open "C:\mailsys\sent.txt" For Input As #3
    Input #3, sent
    Close #3
    sent = sent + 1
    Open "C:\mailsys\sent.txt" For Output As #3
    Write #3, sent
    Close #3
    
    Open "C:\mailsys\memfiles\" & smnum & ".txt" For Append As #1
    Write #1, qun, Furre, mssg & " [Sent: " & Date & " at " & Time & " MST]"
    Close #1
    sckFurc.SendData "wh " & Furre & " Message sent to: " & snd & vbLf
Else
    sckFurc.SendData "wh " & Furre & " " & snd & " is not registered." & vbLf
End If
Else
sckFurc.SendData "wh " & Furre & " You are not registered with me. Whisper me *help to learn how to use my service." & vbLf
End If
Exit Sub
snerr:
    Close #1, #3
    txterr = txterr + 1
    Open "C:\mailsys\errorq.txt" For Input As #5
    Input #5, num
    Close #5
    num = num + 1
    Open "C:\mailsys\errorq.txt" For Output As #5
    Write #5, num
    Close #5
    Open "C:\mailsys\errorlog.txt" For Append As #5
    Write #5, "Send error (sndmsg), from " & Furre & " with " & Txt & ", [When: " & Date & " at " & Time & "]"
    Close #5
    sckFurc.SendData "wh " & Furre & " There has been an error. Whisper me *help to learn how to use my service. If you keep recieveing this error please leave a message for Mys' or e-mail him at mailsys@erenetwork.com." & vbLf
    Resume stoptrying
stoptrying:
End Sub

Sub dojoin(Furre, Txt)
Open "C:\mailsys\members.txt" For Input As #1
Do Until (fName = Furre) Or (EOF(1))
Input #1, fName, mnum
Loop
Close #1

If fName = Furre Then
On Error GoTo joerr
sckFurc.SendData "wh " & Furre & " You are already registered. Whisper me *help to learn how to use my service." & vbLf
Else
Open "C:\mailsys\memnum.txt" For Input As #1
Input #1, nnum
Close #1
num = nnum + 1
Open "C:\mailsys\memnum.txt" For Output As #1
Write #1, num
Close #1


           Open "C:\mailsys\suggestq.txt" For Input As #3
            Input #3, qun
            Close #3
            If qun >= 1 Then
                Open "C:\mailsys\suggest.txt" For Input As #1
                Open "C:\mailsys\suggesta.txt" For Output As #2
                Do Until EOF(1)
                Input #1, frm, snd
                If frm = Furre Then
                    qun = qun - 1
                    Open "C:\mailsys\suggestq.txt" For Output As #3
                    Write #3, qun
                    Close #3
                Else
                    Write #2, frm, snd
                End If 'end if frm = Furre
                Loop
                
                Close #1, #2
            End If 'end if qun >= 1
            Open "C:\mailsys\suggesta.txt" For Input As #1
            Open "C:\mailsys\suggest.txt" For Output As #2
            Do Until EOF(1)
                Input #1, frm, snd
                Write #2, frm, snd
                Loop
            Close #1, #2
            Open "C:\mailsys\suggesta.txt" For Output As #1
            Write #1, ""
            Close #1



Open "C:\mailsys\memfiles\" & num & ".txt" For Output As #1
Write #1, "1", "mailsys", "Welcome to MailSys. The all new Messageing system for the Furries. Place it in your desc and tell your friends. Let them leave you a message when your off line. [Sent: " & Date & " at " & Time & " MST]"
Close #1
Open "C:\mailsys\memfiles\" & num & "q.txt" For Output As #3
Write #3, 1
Close #3
Open "C:\mailsys\members.txt" For Append As #1
Write #1, Furre, num
Close #1
sckFurc.SendData "wh " & Furre & " You are now regiastered with Mailsys. Whisper me *help to learn how to use my service." & vbLf
End If
Exit Sub
joerr:
    Close #1, #2, #3
    txterr = txterr + 1
    Open "C:\mailsys\errorq.txt" For Input As #5
    Input #5, num
    Close #5
    num = num + 1
    Open "C:\mailsys\errorq.txt" For Output As #5
    Write #5, num
    Close #5
    Open "C:\mailsys\errorlog.txt" For Append As #5
    Write #5, "Join error (dojoin), from " & Furre & " with " & Txt & ", [When: " & Date & " at " & Time & "]"
    Close #5
    sckFurc.SendData "wh " & Furre & " There has been an error. Whisper me *help to learn how to use my service. If you keep reciveing this error please leave a message for Mys' or e-mail him at mailsys@erenetwork.com." & vbLf
    Resume stoptrying
stoptrying:
End Sub
Private Sub StayOnline_Timer()
If Connected = True Then
       
        urgc = urgc + 1
        If urgc >= 30 Then
            On Error GoTo sugerr
            urgc = 0
            Open "C:\mailsys\suggestq.txt" For Input As #1
            Input #1, qun
            Close #1
            If qun >= 1 Then
                Open "C:\mailsys\suggest.txt" For Input As #1
                Input #1, fName, sndr
                sckFurc.SendData "wh " & fName & " " & sndr & " Suggested that you join Mailsys Whisper me *help to learn how to use my service." & vbLf
                Do Until (EOF(1))
                Input #1, fName, sndr
                sckFurc.SendData "wh " & fName & " " & sndr & " Suggested that you join Mailsys Whisper me *help to learn how to use my service." & vbLf
                Loop
                Close #1
            End If
        End If
   
       
       
    If prem = 1 Then
        On Error GoTo preerr
        premt = premt + 1
        If premt >= 10 Then
            Open "C:\mailsys\prem.txt" For Input As #1
            Input #1, premote
            Close #1
            sckFurc.SendData Chr(34) & premote & vbLf
            premt = 0
        End If
    End If

    On Error GoTo descerr
    Open "C:\mailsys\memnum.txt" For Input As #1
    Input #1, nnum
    Close #1
    Open "C:\mailsys\sent.txt" For Input As #3
    Input #3, sent
    Close #3
    Open "C:\mailsys\errorq.txt" For Input As #3
    Input #3, errq
    Close #3
    txterr = errq
    txtsent = sent
    txtmem = nnum
    
    Minute = Minute + 1
    If Minute >= 60 Then
        Hour = Hour + 1
        Minute = 0
    End If
    If Hour >= 24 Then
        Day = Day + 1
        Hour = 0
    End If
    timon = Day & ":" & Hour & ":" & Minute
    sckFurc.SendData "desc " & frmConnect.descfield & " [Uptime: "
    If Day >= 1 Then sckFurc.SendData Day & " Day(s) "
    If Hour >= 1 Then sckFurc.SendData Hour & " Hour(s) "
    sckFurc.SendData Minute & " Minute(s)]" & vbLf

End If

If Connected = False And chkcnt = 1 Then
    On Error GoTo tcerr
    sckFurc.Close
    sckFurc.RemoteHost = "66.28.224.193"
    sckFurc.RemotePort = "6000"
    sckFurc.Connect
End If

Exit Sub
tcerr:
    Close #1, #3
    txterr = txterr + 1
    Open "C:\mailsys\errorq.txt" For Input As #5
    Input #5, num
    Close #5
    num = num + 1
    Open "C:\mailsys\errorq.txt" For Output As #5
    Write #5, num
    Close #5
    Open "C:\mailsys\errorlog.txt" For Append As #5
    Write #5, "Relogon error, [When: " & Date & " at " & Time & "]"
    Close #5
    Resume stoptrying
preerr:
    Close #1, #3
    txterr = txterr + 1
    Open "C:\mailsys\errorq.txt" For Input As #5
    Input #5, num
    Close #5
    num = num + 1
    Open "C:\mailsys\errorq.txt" For Output As #5
    Write #5, num
    Close #5
    Open "C:\mailsys\errorlog.txt" For Append As #5
    Write #5, "Premote error, [When: " & Date & " at " & Time & "]"
    Close #5
    sckFurc.Close
    sckFurc.RemoteHost = "66.28.224.193"
    sckFurc.RemotePort = "6000"
    sckFurc.Connect
    Resume stoptrying
descerr:
    Close #1, #3
    txterr = txterr + 1
    Open "C:\mailsys\errorq.txt" For Input As #5
    Input #5, num
    Close #5
    num = num + 1
    Open "C:\mailsys\errorq.txt" For Output As #5
    Write #5, num
    Close #5
    Open "C:\mailsys\errorlog.txt" For Append As #5
    Write #5, "Update time in Description error, [When: " & Date & " at " & Time & "]"
    Close #5
    sckFurc.Close
    sckFurc.RemoteHost = "66.28.224.193"
    sckFurc.RemotePort = "6000"
    sckFurc.Connect
    Resume stoptrying
sugerr:
    Close #1, #3
    txterr = txterr + 1
    Open "C:\mailsys\errorq.txt" For Input As #5
    Input #5, num
    Close #5
    num = num + 1
    Open "C:\mailsys\errorq.txt" For Output As #5
    Write #5, num
    Close #5
    Open "C:\mailsys\errorlog.txt" For Append As #5
    Write #5, "Suggest Timer error, [When: " & Date & " at " & Time & "]"
    Close #5
    sckFurc.Close
    sckFurc.RemoteHost = "66.28.224.193"
    sckFurc.RemotePort = "6000"
    sckFurc.Connect
    Resume stoptrying
stoptrying:
End Sub

Private Sub Styles_Change()

End Sub

Private Sub txtFromFurc_Change()
If Trim(txtFromFurc) = "" Then
    StatusBar.Panels(2).Text = "Words Typed In Furcadia: 0"
    Exit Sub
Else
    Dim temp
    temp = Trim(txtFromFurc) & " " 'add a space, in case 1 word
    temp = Split(temp, " ")
    StatusBar.Panels(2).Text = "Words Typed In Furcadia: " & UBound(temp)
End If
txtFromFurc.SelStart = Len(txtFromFurc)
If Len(txtFromFurc) > 10000 Then txtFromFurc = Right(txtFromFurc, 9000)
End Sub

Private Sub txtSend_Change()
If Trim(txtSend) = "" Then
    StatusBar.Panels(1).Text = "Words: 0"
    Exit Sub
Else
    Dim temp
    temp = Trim(txtSend) & " " 'add a space, in case 1 word
    temp = Split(temp, " ")
    StatusBar.Panels(1).Text = "Words: " & UBound(temp)
End If
End Sub

Private Sub txtSend_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    sckFurc.SendData txtSend & vbLf
    txtSend = ""
    KeyAscii = 0
End If
End Sub

Private Sub ViewError_Click()
If txterr = "0" Then
box = MsgBox("There are no errors", vbOKOnly, "Clear Text")
Else
Load viewerr
viewerr.Show
End If
End Sub

Private Sub ViewMember_Click()
Load viewmem
End Sub
