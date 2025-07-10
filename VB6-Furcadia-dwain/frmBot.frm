VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmBot 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dwain"
   ClientHeight    =   4890
   ClientLeft      =   150
   ClientTop       =   150
   ClientWidth     =   5520
   Icon            =   "frmBot.frx":0000
   LinkTopic       =   "Form1"
   MouseIcon       =   "frmBot.frx":0442
   ScaleHeight     =   4890
   ScaleWidth      =   5520
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox seek 
      Caption         =   "seek"
      Height          =   255
      Left            =   4200
      TabIndex        =   18
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Frame System 
      Caption         =   "System"
      Height          =   1095
      Left            =   4080
      TabIndex        =   14
      Top             =   120
      Width           =   1335
      Begin VB.CheckBox chkBank 
         Caption         =   "Bank"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   720
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.CheckBox chkServtxt 
         Caption         =   "SText"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.CheckBox chkServCode 
         Caption         =   "SCode"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.CommandButton CmdUse 
      Caption         =   "Use"
      Height          =   495
      Left            =   2040
      TabIndex        =   13
      Top             =   4320
      Width           =   615
   End
   Begin VB.CommandButton CmdGet 
      Caption         =   "Get"
      Height          =   495
      Left            =   1440
      TabIndex        =   12
      Top             =   4320
      Width           =   615
   End
   Begin VB.CommandButton cmdGoVinca 
      Caption         =   "Vinca"
      Height          =   255
      Left            =   2760
      TabIndex        =   11
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton cmdGoAlleg 
      Caption         =   "Allegria"
      Height          =   255
      Left            =   2760
      TabIndex        =   10
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton cmdWho 
      Caption         =   "Who"
      Height          =   495
      Left            =   2040
      TabIndex        =   9
      Top             =   3840
      Width           =   615
   End
   Begin VB.CommandButton cmdLay 
      Caption         =   "Lay"
      Height          =   495
      Index           =   0
      Left            =   1440
      TabIndex        =   8
      Top             =   3840
      Width           =   615
   End
   Begin VB.CommandButton cmdSE 
      Caption         =   "SE"
      Height          =   495
      Left            =   720
      TabIndex        =   7
      Top             =   4320
      Width           =   615
   End
   Begin VB.CommandButton cmdSW 
      Caption         =   "SW"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   4320
      Width           =   615
   End
   Begin VB.CommandButton cmdNE 
      Caption         =   "NE"
      Height          =   495
      Left            =   720
      TabIndex        =   5
      Top             =   3840
      Width           =   615
   End
   Begin VB.CommandButton cmdNW 
      Caption         =   "NW"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   3840
      Width           =   615
   End
   Begin VB.CommandButton cmdDisconnect 
      Caption         =   "Disconnect"
      Height          =   255
      Left            =   2760
      TabIndex        =   3
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      Height          =   255
      Left            =   2760
      TabIndex        =   2
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Timer StayOnline 
      Interval        =   60000
      Left            =   120
      Top             =   600
   End
   Begin MSWinsockLib.Winsock sckFurc 
      Left            =   120
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtSend 
      Height          =   285
      Left            =   360
      TabIndex        =   1
      Top             =   3360
      Width           =   3255
   End
   Begin VB.TextBox txtFromFurc 
      Height          =   3135
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "frmBot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Put your bots name and description in the corresponding places
Const BotName = "Dwain"
Const BotPass = "0519aa"
Const descript = "Dwain."
Const ColorCode = "8!!(+*****! ! !!"
Dim Sign As String
Dim saying As String
Dim value As Integer
Dim rank As Integer
Dim lvl As Integer
Dim Valadate As String
Dim newValue As Integer
Dim memver As String
Dim memName As String
'Minute is used for the timer
Public Minute As Integer
'Set your bots ColorCode and Desc in Form_Load
Public Desc As String
'Connected is set to True when connected to Furc and False when
'disconnected from Furc
Public Connected As Boolean


Private Sub chkBank_Click()
If chkBank = 0 Then sckFurc.SendData "m 7" & vbLf
If chkBank = 1 Then sckFurc.SendData "m 1" & vbLf & "m 3" & vbLf
End Sub

Private Sub CmdGet_Click()
sckFurc.SendData "get" & vbLf
End Sub

Private Sub cmdGoAlleg_Click()
'When you click the Allegria button, "goalleg" is sent to Furcadia.
'Which sends your bot to Allegria. vbLf is like hitting the Enter key
sckFurc.SendData "goalleg" & vbLf
End Sub

Private Sub cmdGoVinca_Click()
'When you click the Vinca button, "gostart" is sent to Furcadia.
'Which sends your bot to the Vinca.
sckFurc.SendData "gostart" & vbLf
End Sub

Private Sub cmdlie_Click()
'When you click the Lay button, "lie' is sent to Furcadia.
'Which tells your bot to lie down.
sckFurc.SendData "lie" & vbLf
End Sub

Private Sub cmdNE_Click()
'When you click the NE button, "m 9" is sent to Furcadia.
'Which makes your bot move one space to the northeast.
sckFurc.SendData "m 9" & vbLf
End Sub

Private Sub cmdNW_Click()
'When you click the NW button, "m 7" is sent to Furcadia.
'Which makes your bot move one space to the northwest.
sckFurc.SendData "m 7" & vbLf
End Sub

Private Sub cmdSE_Click()
'When you click the SE button, "m 3" is sent to Furcadia.
'Which makes your bot move one space to the southeast.
sckFurc.SendData "m 3" & vbLf
End Sub

Private Sub cmdSW_Click()
'When you click the SW button, "m 1" is sent to Furcadia.
'Which makes your bot move one space to the southwest.
sckFurc.SendData "m 1" & vbLf
End Sub

Private Sub CmdUse_Click()
sckFurc.SendData "use" & vbLf
End Sub

Private Sub cmdWho_Click()
'When you click the NE button, "who" is sent to Furcadia.
'Which checks to see who is on the current map.
sckFurc.SendData "who" & vbLf
End Sub


Sub Form_Load()
'When the bot program is loading it sets the integer variable to 0.
'The ColorCode variable is set to what your bots color and species code is.
'The Desc variable is set to what your bots description is.
Minute = 0
Desc = descript & " [Uptime: 0 Minute(s)]"
End Sub

Private Sub cmdConnect_Click()
'When you click the connect button and the value of Connected is False,
'the IP and port to Furcadia are set, the bot connects to Furcadia, and
'the value of Connected is changed to True.
If Connected = False Then
sckFurc.RemoteHost = "66.28.224.193"
sckFurc.RemotePort = "6000"
sckFurc.Connect
Connected = True
End If
End Sub

Private Sub cmdDisconnect_Click()
'When you connect the Disconnect button and the value of Connected is True,
'then the connection to Furcadia is closed and the value of Connected is changed
'to False.
If Connected = True Then
sckFurc.Close
Connected = False
End If
End Sub

Private Sub sckFurc_DataArrival(ByVal bytesTotal As Long)
'Declare s As a string. s is the variable that holds the information that
'Furcadia sends to your bot.
Dim s As String
'sckFurc gets the data from Furc and puts it into s
sckFurc.GetData s
'The information that Furcadia sends is split up each time there is a
'vbLf (End of a line) and puts the line into the next array element of x
x = Split(s, vbLf)
'For every line in x, Sub RealText is called.
For r = 0 To UBound(x) - 1
RealText x(r)
Next
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

Sub RealText(Txt)
If chkServtxt.value = Checked Or chkServtxt.Enabled = False Then
'If the checkbox with the Server Code is checked then you see all of the server
'code
If chkServCode.value = Checked Then txtFromFurc = txtFromFurc & Txt & vbCrLf
'If the checkbox with the Server Code label is not checked you do not see any of
'the server code. You'll only see what you would see in the Furcadia client.
If chkServCode.value = Unchecked Then
If Left(Txt, 1) = "(" Then txtFromFurc = txtFromFurc & Right(Txt, Len(Txt) - 1) & vbCrLf
End If
End If
'When the text "END" is sent to the bot, the bot sends the information to login
'to Furcadia
If Txt = "END" Then sckFurc.SendData "connect " & BotName & " " & BotPass & vbLf & "color " & ColorCode & vbLf & "desc " & Desc & vbLf
'When your bot enters a dream, it sends "vascodagama" to Furcadia to let it into
'the dream.
If Txt = "]ccmarbled.pcx" Then
sckFurc.SendData "vascodagama" & vbLf
'sckFurc.SendData "m 7" & vbLf & "use" & vbLf & "m 1" & vbLf & "m 3" & vbLf
Sign = 0
End If
'When someone whispers the bot, it gets there name and message and calls the
'DoWhisper(Furre, Msg) sub which is used to respond to whispers.

If Left(Txt, 11) = "/8!!(+*****" Then
    newm = Right(Txt, Len(Txt) - 17)
    oldm = Right(Txt, Len(Txt) - 11)
    oldm = Left(oldm, Len(oldm) - 6)
doseek oldm, newm
End If

If Left(Txt, 3) = "([ " And Right(Txt, 10) = " to you. ]" Then
    Tmsg = Split(Txt, " whispers, " & Chr(34), 2)
    Furre = Right(Tmsg(0), Len(Tmsg(0)) - 3)
    Msg = Left(Tmsg(1), Len(Tmsg(1)) - 11)
    Furre = LCase(Furre)
    DoWhisper Furre, Msg
End If

End Sub

Sub DoWhisper(Furre, Msg)
doseek oldm, newm
sckFurc.SendData "wh " & Furre & " I dont understand. Try whispering HELP to Penna." & vbLf

End Sub


Private Sub cmdExit_Click()
'When you click the Exit button, the bot program is closed.
End
End Sub

Private Sub seek_Click()
If seek.value = 1 Then
sckFurc.SendData "use" & vbLf & "m 9" & vbLf
End If
End Sub

Private Sub StayOnline_Timer()
'Each minute the timer is set off. The Minute variable is increased by one. Your
'bot changes its desc to add the Minute which is an Uptimer.
Minute = Minute + 1
sckFurc.SendData "desc " & descript & " [Uptime: " & Minute & " Minute(s)]" & vbLf
End Sub

Private Sub txtFromFurc_Change()
'A textbox can only hold a certain amount of text so what is in it is reduced
'when it reaches its max.
txtFromFurc.SelStart = Len(txtFromFurc)
If Len(txtFromFurc) > 10000 Then txtFromFurc = Right(txtFromFurc, 9000)
End Sub

Private Sub txtSend_KeyPress(KeyAscii As Integer)
'When the enter key is pressed, whatever is in the txtSend textbox is sent to
'Furcadia, txtSend is made blank, and KeyAscii = 0 so that you dont hear the beep
'that the enter keys causes.
If KeyAscii = 13 Then
    sckFurc.SendData txtSend & vbLf
    txtSend = ""
    KeyAscii = 0
End If
End Sub
