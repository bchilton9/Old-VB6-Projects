VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmBot 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   0  'None
   Caption         =   "Virtual Furre"
   ClientHeight    =   2310
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   6840
   Icon            =   "frmBot.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmBot.frx":0442
   ScaleHeight     =   2310
   ScaleWidth      =   6840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Height          =   1095
      Left            =   3840
      TabIndex        =   7
      Top             =   840
      Width           =   2895
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Commds"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   720
         Width           =   735
      End
      Begin VB.CommandButton cmdClear 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Clear"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   720
         Width           =   735
      End
      Begin VB.CommandButton cmdSend 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Send"
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox txtSend 
         BackColor       =   &H00FFC0C0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Caption         =   "Chat Box:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Width           =   2655
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00FFFFC0&
      Height          =   615
      Left            =   3840
      TabIndex        =   2
      Top             =   120
      Width           =   2895
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Caption         =   "TimeOn:"
         Height          =   255
         Left            =   1440
         TabIndex        =   6
         Top             =   240
         Width           =   615
      End
      Begin VB.Label timon 
         BackColor       =   &H00FFC0C0&
         Caption         =   "0:0:0"
         Height          =   255
         Left            =   2160
         TabIndex        =   5
         Top             =   240
         Width           =   495
      End
      Begin VB.Label txterr 
         BackColor       =   &H00FFC0C0&
         Caption         =   "0"
         Height          =   255
         Left            =   720
         TabIndex        =   4
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Caption         =   "Errors:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Timer StayOnline 
      Interval        =   60000
      Left            =   3960
      Top             =   1800
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFFC0&
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3615
      Begin VB.TextBox txtFromFurc 
         BackColor       =   &H00FFC0C0&
         Height          =   1695
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   240
         Width           =   3375
      End
   End
   Begin MSWinsockLib.Winsock sckFurc 
      Left            =   3720
      Top             =   1800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lblLoad 
      BackColor       =   &H00FFFFC0&
      Height          =   255
      Left            =   5280
      TabIndex        =   14
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Caption         =   "Bot Loaded:"
      Height          =   255
      Left            =   3840
      TabIndex        =   13
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Menu file 
      Caption         =   "File"
      Begin VB.Menu cmdnew 
         Caption         =   "New"
      End
      Begin VB.Menu cmdopen 
         Caption         =   "Open"
      End
      Begin VB.Menu cmdecit 
         Caption         =   "Edit"
      End
      Begin VB.Menu cmdConnect 
         Caption         =   "Connect"
      End
      Begin VB.Menu cmdDisconnect 
         Caption         =   "Disconnect"
         Enabled         =   0   'False
      End
      Begin VB.Menu cmdExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mwarp 
      Caption         =   "Map Warp's"
      Visible         =   0   'False
      Begin VB.Menu mmaps 
         Caption         =   "Main Map's"
         Begin VB.Menu cmdMaps 
            Caption         =   "Acropolis"
            Index           =   1
         End
         Begin VB.Menu cmdMaps 
            Caption         =   "Allegria"
            Index           =   2
         End
         Begin VB.Menu cmdMaps 
            Caption         =   "Challenge"
            Index           =   3
         End
         Begin VB.Menu cmdMaps 
            Caption         =   "Imaginarium"
            Index           =   4
         End
         Begin VB.Menu cmdMaps 
            Caption         =   "Meovanni"
            Index           =   5
         End
         Begin VB.Menu cmdMaps 
            Caption         =   "New Haven"
            Index           =   6
         End
         Begin VB.Menu cmdMaps 
            Caption         =   "Vinca"
            Index           =   7
         End
      End
      Begin VB.Menu omaps 
         Caption         =   "Other Map's"
         Begin VB.Menu cmdOMaps 
            Caption         =   "Bowling"
            Index           =   1
         End
         Begin VB.Menu cmdOMaps 
            Caption         =   "Challiston City Carnival"
            Index           =   2
         End
         Begin VB.Menu cmdOMaps 
            Caption         =   "Chapel"
            Index           =   3
         End
         Begin VB.Menu cmdOMaps 
            Caption         =   "Chess Island"
            Index           =   4
         End
         Begin VB.Menu cmdOMaps 
            Caption         =   "Connect Four"
            Index           =   5
         End
         Begin VB.Menu cmdOMaps 
            Caption         =   "CraZee PillowZ"
            Index           =   6
         End
         Begin VB.Menu cmdOMaps 
            Caption         =   "Furabbian Nights"
            Index           =   7
         End
         Begin VB.Menu cmdOMaps 
            Caption         =   "Furcadia Con"
            Index           =   8
         End
         Begin VB.Menu cmdOMaps 
            Caption         =   "Goldwyn"
            Index           =   9
         End
         Begin VB.Menu cmdOMaps 
            Caption         =   "Hotel Califurnia"
            Index           =   10
         End
         Begin VB.Menu cmdOMaps 
            Caption         =   "Persona & Roleplay"
            Index           =   11
         End
         Begin VB.Menu cmdOMaps 
            Caption         =   "Pillow Game"
            Index           =   12
         End
         Begin VB.Menu cmdOMaps 
            Caption         =   "Pookie War"
            Index           =   13
         End
         Begin VB.Menu cmdOMaps 
            Caption         =   "Rainbow Beach"
            Index           =   14
         End
         Begin VB.Menu cmdOMaps 
            Caption         =   "Team Furtress"
            Index           =   15
         End
         Begin VB.Menu cmdOMaps 
            Caption         =   "Theriopolis"
            Index           =   16
         End
      End
   End
   Begin VB.Menu command 
      Caption         =   "Movement"
      Visible         =   0   'False
      Begin VB.Menu opncons 
         Caption         =   "Movement Console"
      End
      Begin VB.Menu cmdCommand 
         Caption         =   "Lie"
         Index           =   1
      End
      Begin VB.Menu cmdCommand 
         Caption         =   "Who"
         Index           =   2
      End
      Begin VB.Menu cmdCommand 
         Caption         =   "Sit"
         Index           =   3
      End
      Begin VB.Menu cmdCommand 
         Caption         =   "Stand"
         Index           =   4
      End
      Begin VB.Menu cmdCommand 
         Caption         =   "Get"
         Index           =   5
      End
      Begin VB.Menu cmdCommand 
         Caption         =   "Use"
         Index           =   6
      End
      Begin VB.Menu cmdCommand 
         Caption         =   "Wing's"
         Index           =   7
      End
      Begin VB.Menu cmdCommand 
         Caption         =   "Dragon"
         Index           =   8
      End
      Begin VB.Menu cmdCommand 
         Caption         =   "Dragon Breath"
         Index           =   9
      End
      Begin VB.Menu cmdCommand 
         Caption         =   "Phoenix"
         Index           =   10
      End
      Begin VB.Menu cmdCommand 
         Caption         =   "Phoenix Flame"
         Index           =   11
      End
   End
   Begin VB.Menu cmmd 
      Caption         =   "Command's"
      Begin VB.Menu cmdAdd 
         Caption         =   "Add/Edit/Delete"
      End
   End
   Begin VB.Menu togg 
      Caption         =   "Toggle"
      Visible         =   0   'False
      Begin VB.Menu chkServCode 
         Caption         =   "ServerCode"
      End
      Begin VB.Menu chkWhis 
         Caption         =   "Whisper"
         Checked         =   -1  'True
      End
      Begin VB.Menu chkSay 
         Caption         =   "Say"
         Checked         =   -1  'True
      End
      Begin VB.Menu chkSign 
         Caption         =   "Sign"
         Checked         =   -1  'True
      End
      Begin VB.Menu chkTime 
         Caption         =   "Timer"
         Enabled         =   0   'False
      End
      Begin VB.Menu chkcnt 
         Caption         =   "Auto Connect"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu help 
      Caption         =   "Help"
      Begin VB.Menu about 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmBot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sloc, mess, Desc, howSend, botPath, BotName, BotPass, ColorCode, cDesc, ipath, bload As String
Public frcHost, frcPort, Minute, Hour, day As Integer
Public Connected As Boolean

Private Sub about_Click()
frmabout.Show
End Sub

Private Sub chkcnt_Click()
If chkcnt.Checked = True Then
chkcnt.Checked = False
Else
chkcnt.Checked = True
End If
End Sub

Private Sub chkSay_Click()
If chkSay.Checked = True Then
chkSay.Checked = False
Else
chkSay.Checked = True
End If
End Sub

Private Sub chkServCode_Click()
If chkServCode.Checked = True Then
chkServCode.Checked = False
Else
chkServCode.Checked = True
End If
End Sub

Private Sub chkSign_Click()
If chkSign.Checked = True Then
chkSign.Checked = False
Else
chkSign.Checked = True
End If
End Sub

Private Sub chkWhis_Click()
If chkWhis.Checked = True Then
chkWhis.Checked = False
Else
chkWhis.Checked = True
End If
End Sub

Private Sub cmdAdd_Click()
frmAddFunc.Show
End Sub

Private Sub cmdCommand_Click(Index As Integer)
    Dim action          As String
    Select Case Index
        Case 1 'Lie Down
            action = "lie"
        Case 2 'who
            action = "who"
        Case 3 'Sit down
            action = "sit"
        Case 4 'Stand Up
            action = "stand"
        Case 5 'Pick up object
            action = "get"
        Case 6 'use object
            action = "use"
        Case 7 'Toggle wings
            action = "wings"
        Case 8 'Toggle Dragon
            action = "dragon"
        Case 9 'Dragon breath
            action = "breath"
        Case 10 'Toggle Phoenix
            action = "phoenix"
        Case 11 'Phoenix flame
            action = "flame"
    End Select
    sckFurc.SendData action & vbLf
End Sub

Private Sub cmdecit_Click()
frmEdit.Show
End Sub

Private Sub cmdMaps_Click(Index As Integer)
    Dim action          As String
    Select Case Index
        Case 1 'Go to Acropolis
            action = "gomap *"
        Case 2 'Go to AI
            action = "goalleg"
        Case 3 'Go to Challenge
            action = "gomap $"
        Case 4 'Go to Imaginarium
            action = "gomap ("
        Case 5 'Go to Meovanni
            action = "gomap &"
        Case 6 'Go to New Haven
            action = "gomap '"
        Case 7 'Go to Vinca
            action = "gostart"
    End Select
    sckFurc.SendData action & vbLf
End Sub

Private Sub cmdnew_Click()
frmnew.Show
End Sub

Private Sub cmdOMaps_Click(Index As Integer)
    Dim action          As String
    Select Case Index
        Case 1 'Go to Bowling
            action = "gomap <"
        Case 2 'Go to Challiston City Carnival
            action = "gomap 4"
        Case 3 'Go to Chapel
            action = "gomap 6"
        Case 4 'Go to Chess Island
            action = "gomap 9"
        Case 5 'Go to Connect Four
            action = "gomap ,"
        Case 6 'Go to CraZee PillowZ
            action = "gomap /"
        Case 7 'Go to Furabbian Nights
            action = "gomap %"
        Case 8 'Go to Furcadia Con
            action = "gomap 1"
        Case 9 'Go to Goldwyn
            action = "gomap !"
        Case 10 'Go to Hotel Califurnia
            action = "gomap -"
        Case 11 'Go to Persona & Roleplay
            action = "gomap +"
        Case 12 'Go to Pillow Game
            action = "gomap 2"
        Case 13 'Go to Pookie War
            action = "gomap 5"
        Case 14 'Go to Rainbow Beach
            action = "gomap 8"
        Case 15 'Go to Team Furtress
            action = "gomap 3"
        Case 16 'Go to Theriopolis
            action = "gomap )"
    End Select
    sckFurc.SendData action & vbLf
End Sub

Private Sub cmdopen_Click()
frmopen.Show
End Sub

Private Sub Command1_Click()
comds.Show
End Sub

Sub Form_Load()
Minute = 0
Hour = 0
day = 0
Open "settings.ini" For Input As #1
Input #1, bload
Input #1, frcHost
Input #1, frcPort
Close #1
lblLoad.Caption = bload
If lblLoad.Caption = "Defalt" Or lblLoad.Caption = "" Then
cmdConnect.Enabled = False
cmdecit.Enabled = False
cmmd.Enabled = False
End If
End Sub

Private Sub cmdConnect_Click()
Open bload & ".bot" For Input As #1
Input #1, BotName
Input #1, BotPass
Input #1, cDesc
Input #1, ColorCode
Close #1
Desc = cDesc & " [Uptime: 0 Minute(s)]"
cmdConnect.Enabled = False
cmdopen.Enabled = False
cmdnew.Enabled = False
cmdecit.Enabled = False
sckFurc.RemoteHost = frcHost
sckFurc.RemotePort = frcPort
sckFurc.Connect
End Sub

Private Sub cmdDisconnect_Click()
If Connected = True Then
sckFurc.Close
Connected = False
togg.Visible = False
mwarp.Visible = False
Command.Visible = False
cmdDisconnect.Enabled = False
cmdConnect.Enabled = True
txtSend.Enabled = False
Command1.Enabled = False
cmdClear.Enabled = False
cmdSend.Enabled = False
cmdopen.Enabled = True
cmdnew.Enabled = True
cmdecit.Enabled = True
txtFromFurc = ""
End If
End Sub

Private Sub opncons_Click()
    frmmove.Show
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
If chkServCode.Checked = True Then txtFromFurc = txtFromFurc & Txt & vbCrLf
If chkServCode.Checked = False Then
If Left(Txt, 1) = "(" Then txtFromFurc = txtFromFurc & Right(Txt, Len(Txt) - 1) & vbCrLf
End If
If Txt = "END" Then
sckFurc.SendData "connect " & BotName & " " & BotPass & vbLf & "color " & ColorCode & vbLf & "desc " & Desc & vbLf
Connected = True
togg.Visible = True
mwarp.Visible = True
Command.Visible = True
cmdDisconnect.Enabled = True
txtSend.Enabled = True
Command1.Enabled = True
cmdClear.Enabled = True
cmdSend.Enabled = True
End If


If Txt = "]ccmarbled.pcx" Then sckFurc.SendData "vascodagama" & vbLf
If Left(Txt, 15) Like "(Server going d" Then
    sckFurc.Close
    Connected = False
End If
If Left(Txt, 15) Like "(Someone else h" Then
    sckFurc.Close
    Connected = False
End If
If Left(Txt, 15) Like "(Disconnected f" Then
    sckFurc.Close
    Connected = False
End If

'gets the furres name with whispering to the bot
If chkWhis.Checked = True Then
If Left(Txt, 3) = "([ " And Right(Txt, 10) = " to you. ]" Then
    Tmsg = Split(Txt, " whispers, " & Chr(34), 2)
    Furre = Right(Tmsg(0), Len(Tmsg(0)) - 3)
    Msg = Left(Tmsg(1), Len(Tmsg(1)) - 11)
    styp = "wh"
    docommd Furre, Msg, styp
End If
End If

'Looks at the furre when theay bump a sign
If chkSign.Checked = True Then
If Left(Txt, 1) = "7" Then
    sloc = Left(Txt, Len(Txt) - 6)
    sloc = Right(sloc, Len(sloc) - 5)
    floc = Right(Txt, Len(Txt) - 12)
    sckFurc.SendData "l  " & floc & vbLf
End If
End If

'Gets the furres name when the bot looks at them.
If Left(Txt, 10) = "((You see " Then
    Furre = Mid(Txt, 11, Len(Txt) - 12)
    doSign Furre
End If

'Reads what the furry says
If chkSay.Checked = True Then
If Left(Txt, 1) = "(" Then
num = InStr(1, Txt, ":")
On Error GoTo sayerr
    If Mid(Txt, num, 2) = ": " Then
        Tmsg = Split(Txt, ": ", 2)
        Furre = Right(Tmsg(0), Len(Tmsg(0)) - 1)
        Msg = Left(Tmsg(1), Len(Tmsg(1)))
        styp = "sa"
        docommd Furre, Msg, styp
    End If
End If
End If
Exit Sub
sayerr:
    Resume stoptrying
stoptrying:
End Sub
Sub docommd(Furre, Msg, styp)
On Error GoTo cmderr
Open BotName & ".lst" For Input As #1
Input #1, typ, x, y, rec, mess
Do Until rec = Msg Or EOF(1)
Input #1, typ, x, y, rec, mess
Loop
Close #1
        mess = Replace(mess, "%1", Furre)
        mess = Replace(mess, ";", vbLf)
        mess = Replace(mess, "movesw", "m 1")
        mess = Replace(mess, "movese", "m 3")
        mess = Replace(mess, "movenw", "m 7")
        mess = Replace(mess, "movene", "m 9")
        mess = Replace(mess, "/", "wh ")
        mess = Replace(mess, "say ", Chr(34))
        mess = Replace(mess, "emit ", Chr(34) & "emit")
        mess = Replace(mess, "emitloud ", Chr(34) & "emitloud")
        mess = Replace(mess, "eject ", Chr(34) & "eject")
        mess = Replace(mess, "share ", Chr(34) & "share")
        mess = Replace(mess, "entrymusic ", Chr(34) & "entrymusic")
        mess = Replace(mess, "entrytext ", Chr(34) & "entrytext")
        mess = Replace(mess, "emote ", ":")
        
    If typ = styp And rec Like Msg Then
            sckFurc.SendData mess
    End If
Exit Sub
cmderr:
    Close #1
    Resume stoptrying
stoptrying:
End Sub
Sub doSign(Furre)
    xl = Mid(sloc, 2, 1)
    xl = Asc(xl) - 32
    x = xl * 2
    xl = Mid(sloc, 1, 1)
    xl = Asc(xl) - 32
    x = x + (xl * 2)
    yl = Mid(sloc, 4, 1)
    y = Asc(yl) - 32
    yl = Mid(sloc, 3, 1)
    yl = Asc(yl) - 32
    y = y + (yl * 2)

On Error GoTo sinerr
Open BotName & ".lst" For Input As #1
Input #1, typ, sx, sy, rec, mess
Do Until sx = x And sy = y Or EOF(1)
Input #1, typ, sx, sy, rec, mess
Loop
Close #1
If sx = x And sy = y Then
    mess = Replace(mess, "%1", Furre)
    mess = Replace(mess, ";", vbLf)
    mess = Replace(mess, "movesw", "m 1")
    mess = Replace(mess, "movese", "m 3")
    mess = Replace(mess, "movenw", "m 7")
    mess = Replace(mess, "movene", "m 9")
    mess = Replace(mess, "/", "wh ")
    mess = Replace(mess, "say ", Chr(34))
    mess = Replace(mess, "emit ", Chr(34) & "emit")
    mess = Replace(mess, "emitloud ", Chr(34) & "emitloud")
    mess = Replace(mess, "eject ", Chr(34) & "eject")
    mess = Replace(mess, "share ", Chr(34) & "share")
    mess = Replace(mess, "entrymusic ", Chr(34) & "entrymusic")
    mess = Replace(mess, "entrytext ", Chr(34) & "entrytext")
    mess = Replace(mess, "emote ", ":")
    sckFurc.SendData mess
End If

x = 0
y = 0
Exit Sub
sinerr:
    Close #1
    Resume stoptrying
stoptrying:
End Sub

Private Sub cmdExit_Click()
End
End Sub

Private Sub StayOnline_Timer()
If Connected = True Then
  
    On Error GoTo descerr
    'Open "errorq.txt" For Input As #3
    'Input #3, errq
    'Close #3
    'txterr.Caption = errq
    
    Minute = Minute + 1
    If Minute >= 60 Then
        Hour = Hour + 1
        Minute = 0
    End If
    If Hour >= 24 Then
        day = day + 1
        Hour = 0
    End If
    timon.Caption = day & ":" & Hour & ":" & Minute
    sckFurc.SendData "desc " & cDesc & " [Uptime: "
    If day >= 1 Then sckFurc.SendData day & " Day(s) "
    If Hour >= 1 Then sckFurc.SendData Hour & " Hour(s) "
    sckFurc.SendData Minute & " Minute(s)]" & vbLf

End If

If Connected = False And chkcnt.Enabled = True Then
    On Error GoTo tcerr
    sckFurc.Close
    sckFurc.RemoteHost = "66.28.224.193"
    sckFurc.RemotePort = "6000"
    sckFurc.Connect
End If

Exit Sub
tcerr:
    Resume stoptrying
descerr:
    sckFurc.Close
    sckFurc.RemoteHost = "66.28.224.193"
    sckFurc.RemotePort = "6000"
    sckFurc.Connect
    Resume stoptrying
stoptrying:
End Sub

Private Sub txtFromFurc_Change()
txtFromFurc.SelStart = Len(txtFromFurc)
If Len(txtFromFurc) > 10000 Then txtFromFurc = Right(txtFromFurc, 9000)
End Sub
Private Sub cmdSend_Click()
    txtSend = Replace(txtSend, ";", vbLf)
    txtSend = Replace(txtSend, "movesw", "m 1")
    txtSend = Replace(txtSend, "movese", "m 3")
    txtSend = Replace(txtSend, "movenw", "m 7")
    txtSend = Replace(txtSend, "movene", "m 9")
    txtSend = Replace(txtSend, "/", "wh ")
    txtSend = Replace(txtSend, "say ", Chr(34))
    txtSend = Replace(txtSend, "emit ", Chr(34) & "emit")
    txtSend = Replace(txtSend, "emitloud ", Chr(34) & "emitloud")
    txtSend = Replace(txtSend, "eject ", Chr(34) & "eject")
    txtSend = Replace(txtSend, "share ", Chr(34) & "share")
    txtSend = Replace(txtSend, "entrymusic ", Chr(34) & "entrymusic")
    txtSend = Replace(txtSend, "entrytext ", Chr(34) & "entrytext")
    txtSend = Replace(txtSend, "emote ", ":")
        sckFurc.SendData txtSend & vbLf
        txtSendFur = ""
        txtSend = ""
End Sub
Private Sub txtSend_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtSend = Replace(txtSend, ";", vbLf)
    txtSend = Replace(txtSend, "movesw", "m 1")
    txtSend = Replace(txtSend, "movese", "m 3")
    txtSend = Replace(txtSend, "movenw", "m 7")
    txtSend = Replace(txtSend, "movene", "m 9")
    txtSend = Replace(txtSend, "/", "wh ")
    txtSend = Replace(txtSend, "say ", Chr(34))
    txtSend = Replace(txtSend, "emit ", Chr(34) & "emit")
    txtSend = Replace(txtSend, "emitloud ", Chr(34) & "emitloud")
    txtSend = Replace(txtSend, "eject ", Chr(34) & "eject")
    txtSend = Replace(txtSend, "share ", Chr(34) & "share")
    txtSend = Replace(txtSend, "entrymusic ", Chr(34) & "entrymusic")
    txtSend = Replace(txtSend, "entrytext ", Chr(34) & "entrytext")
    txtSend = Replace(txtSend, "emote ", ":")
        sckFurc.SendData txtSend & vbLf
        txtSendFur = ""
        txtSend = ""
    KeyAscii = 0
End If
End Sub
Private Sub cmdClear_Click()
        txtSendFur = ""
        txtSend = ""
End Sub
