VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBot 
   BackColor       =   &H8000000B&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MailSys"
   ClientHeight    =   3975
   ClientLeft      =   150
   ClientTop       =   150
   ClientWidth     =   3615
   Icon            =   "frmBot.frx":0000
   LinkTopic       =   "Form1"
   MouseIcon       =   "frmBot.frx":0442
   ScaleHeight     =   3975
   ScaleWidth      =   3615
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSndDate 
      DataField       =   "DateSent"
      DataSource      =   "dbaMess"
      Height          =   285
      Left            =   1080
      TabIndex        =   46
      Top             =   5760
      Width           =   1095
   End
   Begin VB.TextBox txtMess 
      DataField       =   "Message"
      DataSource      =   "dbaMess"
      Height          =   285
      Left            =   1080
      TabIndex        =   44
      Top             =   5400
      Width           =   1095
   End
   Begin VB.TextBox txtFrom 
      DataField       =   "From"
      DataSource      =   "dbaMess"
      Height          =   285
      Left            =   1080
      TabIndex        =   42
      Top             =   5040
      Width           =   1095
   End
   Begin VB.TextBox txtTo 
      DataField       =   "To"
      DataSource      =   "dbaMess"
      Height          =   285
      Left            =   1080
      TabIndex        =   40
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Data dbaMess 
      Caption         =   "Messages"
      Connect         =   "Access"
      DatabaseName    =   "mailsys.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "Messages"
      Top             =   4200
      Width           =   2175
   End
   Begin VB.CommandButton cmdMem 
      Caption         =   "Member >"
      Height          =   375
      Left            =   4800
      TabIndex        =   38
      Top             =   1920
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Members"
      Height          =   1815
      Left            =   5880
      TabIndex        =   31
      Top             =   120
      Width           =   2415
      Begin VB.Data members 
         Caption         =   "members"
         Connect         =   "Access"
         DatabaseName    =   "mailsys.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   0  'Table
         RecordSource    =   "Member"
         Top             =   1320
         Width           =   2175
      End
      Begin VB.TextBox txtName 
         DataField       =   "Name"
         DataSource      =   "members"
         Height          =   285
         Left            =   720
         TabIndex        =   34
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtDate 
         DataField       =   "Joined"
         DataSource      =   "members"
         Height          =   285
         Left            =   720
         TabIndex        =   33
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtStatus 
         DataField       =   "Stats"
         DataSource      =   "members"
         Height          =   285
         Left            =   720
         TabIndex        =   32
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Name"
         Height          =   255
         Left            =   240
         TabIndex        =   37
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Date"
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Stat's"
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   960
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdPanel 
      BackColor       =   &H8000000A&
      Caption         =   "More >"
      Height          =   255
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   3360
      Width           =   735
   End
   Begin VB.TextBox txtSend 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000007&
      Height          =   285
      Left            =   120
      TabIndex        =   28
      Top             =   3000
      Width           =   3375
   End
   Begin VB.CommandButton cmdButton 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      Caption         =   "Send"
      Height          =   255
      Index           =   11
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   3360
      Width           =   735
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   26
      Top             =   3720
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   450
      Style           =   1
      SimpleText      =   "Words Typed In Furcadia: 0"
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdButton 
      BackColor       =   &H8000000A&
      Caption         =   "Movement"
      Height          =   255
      Index           =   2
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton cmdButton 
      BackColor       =   &H8000000A&
      Caption         =   "Flame"
      Height          =   255
      Index           =   9
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton cmdButton 
      BackColor       =   &H8000000A&
      Caption         =   "Phoenix"
      Height          =   255
      Index           =   8
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   2880
      Width           =   855
   End
   Begin VB.TextBox insult 
      Height          =   285
      Left            =   120
      TabIndex        =   16
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox usrname 
      Height          =   285
      Left            =   1440
      TabIndex        =   15
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton cmdButton 
      BackColor       =   &H8000000A&
      Caption         =   " Wing's"
      Height          =   255
      Index           =   10
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3360
      Width           =   855
   End
   Begin VB.CommandButton cmdButton 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      Caption         =   "Clear"
      Height          =   255
      Index           =   12
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3360
      Width           =   735
   End
   Begin VB.Frame textm 
      BackColor       =   &H8000000B&
      Height          =   1335
      Left            =   120
      TabIndex        =   10
      Top             =   0
      Width           =   3375
      Begin VB.TextBox timon 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000007&
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   21
         Text            =   "0:0:0"
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox txtcnt 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000007&
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   19
         Text            =   "False"
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtsent 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000007&
         Height          =   285
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   17
         Text            =   "0"
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox txtmem 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000007&
         Height          =   285
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "0"
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label8 
         BackColor       =   &H8000000B&
         Caption         =   "Members's:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   360
         TabIndex        =   30
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000B&
         Caption         =   "TimeOn:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   2040
         TabIndex        =   22
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000B&
         Caption         =   "Connected:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   2040
         TabIndex        =   20
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000B&
         Caption         =   "Sent:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   360
         TabIndex        =   18
         Top             =   720
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdButton 
      BackColor       =   &H8000000A&
      Caption         =   "&Vinca"
      Height          =   255
      Index           =   4
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton cmdButton 
      BackColor       =   &H8000000A&
      Caption         =   "&Allegria"
      Height          =   255
      Index           =   3
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3120
      Width           =   975
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H8000000B&
      ForeColor       =   &H8000000E&
      Height          =   1575
      Left            =   3720
      TabIndex        =   3
      Top             =   0
      Width           =   1935
      Begin VB.CheckBox chkcnt 
         BackColor       =   &H8000000B&
         Caption         =   "Auto Con"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   1200
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.CheckBox prem 
         BackColor       =   &H8000000B&
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
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   960
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.CheckBox chkServtxt 
         BackColor       =   &H8000000B&
         Caption         =   "SText"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkWhisp 
         BackColor       =   &H8000000B&
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
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Value           =   2  'Grayed
         Width           =   975
      End
      Begin VB.CheckBox chkServCode 
         BackColor       =   &H8000000B&
         Caption         =   "SCode"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdButton 
      BackColor       =   &H8000000A&
      Caption         =   "&Disconnect"
      Height          =   375
      Index           =   1
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton cmdButton 
      BackColor       =   &H8000000A&
      Caption         =   "&Connect"
      Height          =   375
      Index           =   0
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1920
      Width           =   975
   End
   Begin VB.Timer StayOnline 
      Interval        =   60000
      Left            =   720
      Top             =   1560
   End
   Begin MSWinsockLib.Winsock sckFurc 
      Left            =   240
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtFromFurc 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000007&
      Height          =   1455
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1440
      Width           =   3375
   End
   Begin VB.Label Label12 
      Caption         =   "Date"
      Height          =   255
      Left            =   120
      TabIndex        =   45
      Top             =   5760
      Width           =   735
   End
   Begin VB.Label Label11 
      Caption         =   "Message"
      Height          =   255
      Left            =   120
      TabIndex        =   43
      Top             =   5400
      Width           =   735
   End
   Begin VB.Label Label10 
      Caption         =   "From"
      Height          =   255
      Left            =   120
      TabIndex        =   41
      Top             =   5040
      Width           =   735
   End
   Begin VB.Label Label9 
      Caption         =   "To"
      Height          =   255
      Left            =   120
      TabIndex        =   39
      Top             =   4680
      Width           =   735
   End
End
Attribute VB_Name = "frmBot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Minute, Hour, Day, urgc, premt, Space, del, frcHost, frcPort As Integer
Public Connected As Boolean
Public BotName, BotPass, vers, descrip, ColorCode, Desc As String

Private Sub cmdMem_Click()
If frmBot.Width = 8505 Then
    
    cmdMem.Caption = "Member >"
    frmBot.Width = 5865
Else
    cmdMem.Caption = "Member <"
    frmBot.Width = 8505
End If
End Sub

Sub Form_Load()
dosettings
End Sub
Private Sub sckFurc_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Dim s As String
sckFurc.GetData s
x = Split(s, vbLf)
For r = 0 To UBound(x) - 1
RealText x(r)
Next
End Sub
Sub RealText(Txt)
On Error Resume Next
If chkServtxt.Value = Checked Or chkServtxt.Enabled = False Then
    If chkServCode.Value = Checked Then txtFromFurc = txtFromFurc & Txt & vbCrLf
    If chkServCode.Value = Unchecked Then
        If Left(Txt, 1) = "(" Then txtFromFurc = txtFromFurc & Right(Txt, Len(Txt) - 1) & vbCrLf
    End If
End If

If Txt = "END" Then
    sckFurc.SendData "connect " & BotName & " " & BotPass & vbLf & "color " & ColorCode & vbLf & "desc " & Desc & vbLf
    Connected = "True"
    txtcnt = "True"
    sckFurc.SendData "goalleg" & vbLf
End If
If Txt = "]ccmarbled.pcx" Then
    sckFurc.SendData "vascodagama" & vbLf
    'sckFurc.SendData "goalleg" & vbLf
End If

incomeingtxt Txt
End Sub
Private Sub txtFromFurc_Change()
On Error Resume Next
If Trim(txtFromFurc) = "" Then
    StatusBar.SimpleText = "Words Typed In Furcadia: 0"
    Exit Sub
Else
    Dim temp
    temp = Trim(txtFromFurc) & " " 'add a space, in case 1 word
    temp = Split(temp, " ")
    StatusBar.SimpleText = "Words Typed In Furcadia: " & UBound(temp)
End If
txtFromFurc.SelStart = Len(txtFromFurc)
If Len(txtFromFurc) > 10000 Then txtFromFurc = Right(txtFromFurc, 9000)
End Sub
Private Sub StayOnline_Timer()
On Error Resume Next
If Connected = True Then
    timeon
End If
If Connected = False And chkcnt = 1 Then
    sckFurc.Close
    Connected = False
    txtcnt = "False"
    sckFurc.RemoteHost = frcHost
    sckFurc.RemotePort = frcPort
    sckFurc.Connect
End If
End Sub
Private Sub chkServCode_Click()
On Error Resume Next
If chkServCode = 1 Then
    chkServtxt = 2
    chkServtxt.Enabled = False
End If
If chkServCode = 0 Then
    chkServtxt = 1
    chkServtxt.Enabled = True
End If
End Sub
Private Sub prem_Click()
On Error Resume Next
If prem = 1 Then
    chkWhisp = 2
    chkWhisp.Enabled = False
End If
If prem = 0 Then
    chkWhisp = 1
    chkWhisp.Enabled = True
End If
End Sub
Private Sub cmdPanel_Click()
On Error Resume Next
If frmBot.Width = 3675 Then
    cmdPanel.Caption = "Less <"
    frmBot.Width = 5865
Else
    cmdPanel.Caption = "More >"
    frmBot.Width = 3675
    cmdMem.Caption = "Member >"
End If
End Sub
Private Sub cmdButton_Click(Index As Integer)
On Error Resume Next
    Select Case Index
        Case 0 'Connect
            If Connected = False Then
                On Error GoTo conerr
                sckFurc.RemoteHost = frcHost
                sckFurc.RemotePort = frcPort
                sckFurc.Connect
                On Error Resume Next
            End If
        Case 1 'Disconnect
            If Connected = True Then
                sckFurc.Close
                Connected = False
                txtcnt = "False"
            End If
        Case 2 'Movement
            frmmove.Show
        Case 3 'Allegra
            If Connected = True Then sckFurc.SendData "goalleg" & vbLf
        Case 4 'Vinca
            If Connected = True Then sckFurc.SendData "gostart" & vbLf
        Case 5 'view error
            If txterr = "0" Then
                box = MsgBox("There are no errors", vbOKOnly, "Clear Text")
            Else
                Load viewerr
                viewerr.Show
            End If
        Case 6 'clear error
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
        Case 8 'phoenix
            If Connected = True Then sckFurc.SendData "phoenix" & vbLf
        Case 9 'flame
            If Connected = True Then sckFurc.SendData "flame" & vbLf
        Case 10 'wings
            If Connected = True Then sckFurc.SendData "wings" & vbLf
        Case 11 'send
            If Connected = True Then
                sckFurc.SendData txtSend & vbLf
                txtSend = ""
            End If
        Case 12 'clear
            box = MsgBox("Are you sure you want to clear all the text from Furcadia.", vbOKCancel, "Clear Text")
            If box = vbOK Then txtFromFurc = ""
        End Select
Exit Sub
conerr:
    box = MsgBox("Please Wait or press disconnect.", vbOKOnly, "Connect Error")
    Resume stoptrying
stoptrying:
End Sub
Private Sub txtSend_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 And Connected = True Then
    sckFurc.SendData txtSend & vbLf
    txtSend = ""
    KeyAscii = 0
End If
End Sub
