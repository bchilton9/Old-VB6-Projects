VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmBot 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Jovati"
   ClientHeight    =   4785
   ClientLeft      =   150
   ClientTop       =   150
   ClientWidth     =   5385
   Icon            =   "frmBot.frx":0000
   LinkTopic       =   "Form1"
   MouseIcon       =   "frmBot.frx":0442
   ScaleHeight     =   4785
   ScaleWidth      =   5385
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tfight 
      Interval        =   2000
      Left            =   120
      Top             =   1080
   End
   Begin VB.CommandButton cmdturnl 
      Caption         =   "Turn >"
      Height          =   495
      Left            =   2040
      TabIndex        =   28
      Top             =   4200
      Width           =   615
   End
   Begin VB.CommandButton cmdturnr 
      Caption         =   "< Turn"
      Height          =   495
      Left            =   1440
      TabIndex        =   27
      Top             =   4200
      Width           =   615
   End
   Begin VB.CommandButton cmdGoVinca 
      Caption         =   "&Vinca"
      Height          =   375
      Left            =   2760
      TabIndex        =   22
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton cmdGoAlleg 
      Caption         =   "&Allegria"
      Height          =   375
      Left            =   2760
      TabIndex        =   21
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      Caption         =   "System"
      Height          =   1335
      Left            =   4080
      TabIndex        =   15
      Top             =   1920
      Width           =   1215
      Begin VB.CheckBox chkServtxt 
         Caption         =   "SText"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkWhisp 
         Caption         =   "Whispers"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   960
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CheckBox chkFollow 
         Caption         =   "Follow"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   720
         Width           =   975
      End
      Begin VB.CheckBox chkServCode 
         Caption         =   "SCode"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Fun"
      Height          =   615
      Left            =   4080
      TabIndex        =   14
      Top             =   3360
      Width           =   1215
      Begin VB.CheckBox chkbar 
         Caption         =   "BarTend"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Fighting"
      Height          =   1695
      Left            =   4080
      TabIndex        =   12
      Top             =   120
      Width           =   1215
      Begin VB.CheckBox chkDream 
         Caption         =   "Dream"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   960
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CommandButton cmdrest 
         Caption         =   "&Reset"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   1320
         Width           =   975
      End
      Begin VB.CheckBox chkTurn 
         Caption         =   "Turns"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   720
         Width           =   975
      End
      Begin VB.CheckBox chkTim 
         Caption         =   "Timer"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   480
         Width           =   975
      End
      Begin VB.CheckBox chkFight 
         Caption         =   "On/Off"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.CommandButton cmduse 
      Caption         =   "&Use"
      Height          =   495
      Left            =   2040
      TabIndex        =   11
      Top             =   3720
      Width           =   615
   End
   Begin VB.CommandButton cmdGet 
      Caption         =   "&Get"
      Height          =   495
      Left            =   1440
      TabIndex        =   10
      Top             =   3720
      Width           =   615
   End
   Begin VB.CommandButton cmdWho 
      Caption         =   "&Who"
      Height          =   495
      Left            =   2040
      TabIndex        =   9
      Top             =   3240
      Width           =   615
   End
   Begin VB.CommandButton cmdLay 
      Caption         =   "&Lay"
      Height          =   495
      Index           =   0
      Left            =   1440
      TabIndex        =   8
      Top             =   3240
      Width           =   615
   End
   Begin VB.CommandButton cmdSE 
      Caption         =   "SE"
      Height          =   495
      Left            =   720
      TabIndex        =   7
      Top             =   3960
      Width           =   615
   End
   Begin VB.CommandButton cmdSW 
      Caption         =   "SW"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   3960
      Width           =   615
   End
   Begin VB.CommandButton cmdNE 
      Caption         =   "NE"
      Height          =   495
      Left            =   720
      TabIndex        =   5
      Top             =   3480
      Width           =   615
   End
   Begin VB.CommandButton cmdNW 
      Caption         =   "NW"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   3480
      Width           =   615
   End
   Begin VB.CommandButton cmdDisconnect 
      Caption         =   "&Disconnect"
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "&Connect"
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   3240
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
      Left            =   120
      TabIndex        =   1
      Top             =   2880
      Width           =   3855
   End
   Begin VB.TextBox txtFromFurc 
      Height          =   2655
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
Dim Sign As String
Dim lastwalk As String
Dim whatwalk As String
Dim hit
Public Minute As Integer
Public onet
Public twot
Public Desc As String
Public Connected As Boolean
'Bot Settings
Const BotName = "Valka"
Const BotPass = "0519aa"
Const descrip = "This mighty Fighter is a brave adventurer and good friend. He defends his companions against monsters and outher enemies with fierce devotion."
Const ColorCode = "! G2+88888!#!!#!"
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
Private Sub chkTurn_Click()
If chkTurn = 1 Then
chkDream.Enabled = False
chkTim.Enabled = False
End If
If chkTurn = 0 Then
chkDream.Enabled = True
chkTim.Enabled = True
End If
End Sub
Private Sub chkDream_Click()
If chkDream = 1 Then
chkTurn.Enabled = False
chkTim.Enabled = False
End If
If chkDream = 0 Then
chkTurn.Enabled = True
chkTim.Enabled = True
End If
End Sub
Private Sub chkTim_Click()
If chkTim = 1 Then
chkTurn.Enabled = False
chkDream.Enabled = False
End If
If chkTim = 0 Then
chkTurn.Enabled = True
chkDream.Enabled = True
End If
End Sub
Private Sub cmdGet_Click()
sckFurc.SendData "get" & vbLf
End Sub
Private Sub cmdGoAlleg_Click()
sckFurc.SendData "goalleg" & vbLf
End Sub
Private Sub cmdGoVinca_Click()
sckFurc.SendData "gostart" & vbLf
End Sub
Private Sub cmdlie_Click()
sckFurc.SendData "lie" & vbLf
End Sub
Private Sub cmdNE_Click()
sckFurc.SendData "m 9" & vbLf
End Sub
Private Sub cmdNW_Click()
sckFurc.SendData "m 7" & vbLf
End Sub
Private Sub cmdrest_Click()
sckFurc.SendData Chr(34) & "emit Fight system has been reset." & vbLf
sckFurc.SendData "m 1" & vbLf
Open "fnum.txt" For Output As #1
Write #1, 0
Close #1
Open "fighter1.txt" For Output As #1
Write #1, "na", 0, 0, 0, 0, 0, "na"
Close #1
Open "fighter2.txt" For Output As #1
Write #1, "na", 0, 0, 0, 0, 0, "na"
Close #1
End Sub
Private Sub cmdSE_Click()
sckFurc.SendData "m 3" & vbLf
End Sub
Private Sub cmdSW_Click()
sckFurc.SendData "m 1" & vbLf
End Sub

Private Sub cmdturnl_Click()
sckFurc.SendData ">" & vbLf
End Sub

Private Sub cmdturnr_Click()
sckFurc.SendData "<" & vbLf
End Sub

Private Sub cmduse_Click()
sckFurc.SendData "use" & vbLf
End Sub
Private Sub cmdWho_Click()
sckFurc.SendData "who" & vbLf
End Sub
Sub Form_Load()
Minute = 0
Desc = descrip & " [Uptime: 0 Minute(s)]"
End Sub
Private Sub cmdConnect_Click()
If Connected = False Then
sckFurc.RemoteHost = "66.28.224.193"
sckFurc.RemotePort = "6000"
sckFurc.Connect
Connected = True
lastwalk = "none"
End If
End Sub
Private Sub cmdDisconnect_Click()
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
X = Split(s, vbLf)
'For every line in x, Sub RealText is called.
For r = 0 To UBound(X) - 1
RealText X(r)
Next
End Sub
Sub RealText(Txt)
If chkServtxt.Value = Checked Or chkServtxt.Enabled = False Then
'If the checkbox with the Server Code is checked then you see all of the server
'code
If chkServCode.Value = Checked Then txtFromFurc = txtFromFurc & Txt & vbCrLf
'If Left(Txt, 1) = "/" Then txtFromFurc = txtFromFurc & Left(Txt, Len(Txt) - 6) & vbCrLf
'If the checkbox with the Server Code label is not checked you do not see any of
'the server code. You'll only see what you would see in the Furcadia client.
If chkServCode.Value = Unchecked Then
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
sckFurc.SendData "m 1" & vbLf
End If
'When someone whispers the bot, it gets there name and message and calls the
'DoWhisper(Furre, Msg) sub which is used to respond to whispers.
If chkWhisp.Value = Checked Then
If Left(Txt, 3) = "([ " And Right(Txt, 10) = " to you. ]" Then
    tmsg = Split(Txt, " whispers, " & Chr(34), 2)
    Furre = Right(tmsg(0), Len(tmsg(0)) - 3)
    NMsg = Left(tmsg(1), Len(tmsg(1)) - 11)
    Msg = LCase(NMsg)
    DoWhisper Furre, Msg
End If
End If 'chkWhisp
'make the bot follow it owner
If chkFollow.Value = Checked Then
If Left(Txt, 11) = "/8!8)+=====" Then
        frl = Mid(Txt, 17, Len(Txt) - 0)
        whatwalk = Mid(frl, 1, Len(frl) - 4)
        'whatwalk = LCase(wwalk)
    dowalk whatwalk, lastwalk
End If
End If 'chkFollow


'make the bot act like a bartender
If chkbar.Value = Checked Then
If Left(Txt, 1) = "(" And Right(Txt, 1) = "#" Then
    tmsg = Split(Txt, ": #", Len(Txt), 2)
    Furre = Right(tmsg(0), Len(tmsg(0)) - 1)
    ord = Left(tmsg(1), Len(tmsg(1)) - 1)
    Order = LCase(ord)
    doserve Furre, Order
End If
End If 'chkBar


'watch for fight data
If chkFight.Value = Checked Then
Open "fnum.txt" For Input As #1
Input #1, fnum
Close #1
If fnum = 2 Then
        If chkTurn = "1" Then
        If Left(Txt, 1) = "(" And Right(Txt, 1) = "@" Then
            tmsg = Split(Txt, ": ", 2)
            Furre = Right(tmsg(0), Len(tmsg(0)) - 1)
            'atk = Left(Tmsg(1), Len(Tmsg(1)) - 1)
            Open "turn.txt" For Input As #1
            Input #1, turn
            Close #1
            
            
            If turn = 1 Then
                Open "fighter1.txt" For Input As #1
                Input #1, fName, a, s, d, f, g, h
                Close #1
                End If
            If turn = 2 Then
                Open "fighter2.txt" For Input As #1
                Input #1, fName, a, s, d, f, g, h
                Close #1
            End If
            
            
            If fName = Furre Then
                doattk Furre, turn
            Else
                sckFurc.SendData "wh " & Furre & " Please wate intel your turn" & vbLf
            End If
            End If
        End If
        
        If chkTim = "1" Then
        If Left(Txt, 1) = "(" And Right(Txt, 1) = "@" Then
            tmsg = Split(Txt, ": ", 2)
            Furre = Right(tmsg(0), Len(tmsg(0)) - 1)
            attk = Left(tmsg(1), Len(tmsg(1)) - 1)

        End If
        End If
        
End If
        If chkDream = Checked Then
        'watch for data 7 = 3 = 4 J = 4     7 ? 7 > 6 F > 6
            If Left(Txt, 2) = "7 " And Right(Txt, 3) = "= 4" Then
            sckFurc.SendData "l  = 4" & vbLf
            End If
            If Left(Txt, 2) = "7 " And Right(Txt, 3) = "> 6" Then
            sckFurc.SendData "l  > 6" & vbLf
            End If
            
            
            'Gets the furres name when the bot looks at them.
            If Left(Txt, 10) = "((You see " Then
            Furre = Mid(Txt, 11, Len(Txt) - 12)
            Open "members.txt" For Input As #1
            Input #1, fName, mnum
            Do Until (fName = Furre) Or (EOF(1))
            Input #1, fName, mnum
            Loop
            Close #1
            If fName = Furre Then
                Open "\memfiles\" & mnum & ".txt" For Input As #1
                Input #1, fName, mnum, lvl, clas, gold, xp, weap, armo, sone, stwo, sthe, sfor, sfiv
                Close #1
                    Open "fnum.txt" For Input As #1
                    Input #1, fnum
                    Close #1
                        If fnum = 0 Then
                            Open "fnum.txt" For Output As #2
                            Write #2, 1
                            Close #2
                            
    If clas = "Wizard" Then
    hp = lvl * 15
    man = lvl * 20
    End If
    If clas = "Fighter" Then
    hp = lvl * 20
    man = 0
    End If
    If clas = "Thief" Then
    hp = lvl * 15
    man = lvl * 10
    End If
    If clas = "Paladin" Then
    hp = lvl * 20
    man = lvl * 10
    End If
    If clas = "Priest" Then
    hp = lvl * 10
    man = lvl * 10
    End If
                            Open "fighter1.txt" For Output As #3
                            Write #3, Furre, lvl, hp, man, weap, armo, Class
                            Close #3
                        End If
                        If fnum = 1 Then
                            Open "fnum.txt" For Output As #2
                            Write #2, 2
                            Close #2
    If clas = "Wizard" Then
    hp = lvl * 15
    man = lvl * 20
    End If
    If clas = "Fighter" Then
    hp = lvl * 20
    man = 0
    End If
    If clas = "Thief" Then
    hp = lvl * 15
    man = lvl * 10
    End If
    If clas = "Paladin" Then
    hp = lvl * 20
    man = lvl * 10
    End If
    If clas = "Priest" Then
    hp = lvl * 10
    man = lvl * 10
    End If
                            Open "fighter2.txt" For Output As #3
                            Write #3, Furre, lvl, hp, man, weap, armo, Class
                            Close #3
                        End If
                        
                        
If weap = 0 Then weap = "Paws"
If weap = 1 Then weap = "Dagger"
If weap = 2 Then weap = "Knife"
If weap = 3 Then weap = "Hand ax"
If weap = 4 Then weap = "Quarterstaff"
If weap = 5 Then weap = "Spear"
If weap = 6 Then weap = "Warhammer"
If weap = 7 Then weap = "Battle ax"
If weap = 8 Then weap = "Morneing Star"
If weap = 9 Then weap = "Flail"
If weap = 10 Then weap = "Mace"
If weap = 11 Then weap = "Broad Sword"
If weap = 12 Then weap = "Short Bow"
If weap = 13 Then weap = "Crossbow"
If weap = 14 Then weap = "Shord Sword"
If weap = 15 Then weap = "Long Sword"
If weap = 16 Then weap = "TwoHand Sword"

If armo = 0 Then armo = "Fir"
If armo = 1 Then armo = "Padded"
If armo = 2 Then armo = "Leather"
If armo = 3 Then armo = "Chain Mail"
If armo = 4 Then armo = "Splint Mail"
If armo = 5 Then armo = "Ring Mail"
If armo = 6 Then armo = "Scale Mail"
If armo = 7 Then armo = "Banded Mail"
If armo = 8 Then armo = "Plate Mail"
        sckFurc.SendData Chr(34) & "emit " & Furre & " The " & clas & " Welding " & weap & " and " & armo & " has entered the fighting area." & vbLf
            Open "fnum.txt" For Input As #1
            Input #1, fnum
            Close #1
            If fnum = 2 Then
                sckFurc.SendData Chr(34) & "emit let the fight begine" & vbLf
                turn = 1
                dofdream turn
            End If
        Else
            sckFurc.SendData Chr(34) & "emit " & Furre & " is not a registered member." & vbLf
        End If
        
            End If
        End If
End If 'fight
End Sub
Sub DoWhisper(Furre, Msg)
'When anyone whispers the bot anything it whispers back
'The Like operator is used for string comparisons
If Msg Like "help" Then
whspnum = 1
Else
If Msg Like "stats" Then
whspnum = 2
Else
If Msg Like "join fight" Then
whspnum = 3
Else
If Msg Like "join" Then
whspnum = 4
Else
If Msg Like "join fighter" Then
whspnum = 5
clas = "Fighter"
Else
If Msg Like "join wizard" Then
whspnum = 5
clas = "Wizard"
Else
If Msg Like "join thief" Then
whspnum = 5
clas = "Thief"
Else
If Msg Like "join paladin" Then
whspnum = 5
clas = "Paladin"
Else
If Msg Like "join priest" Then
whspnum = 5
clas = "Preiest"
Else
If Msg Like "fighter" Then
whspnum = 6
Else
If Msg Like "wizard" Then
whspnum = 7
Else
If Msg Like "thief" Then
whspnum = 8
Else
If Msg Like "paladin" Then
whspnum = 9
Else
If Msg Like "priest" Then
whspnum = 10
Else
If Msg Like "fighting" Then
whspnum = 11
Else
If Msg Like "how to fight" Then
whspnum = 12
Else
If Msg Like "bar" Then
whspnum = 13
Else
If Msg Like "menu" Then
whspnum = 14
Else
If Msg Like "buy weapon 1" Then
wea = 1
whspnum = 15
Else
If Msg Like "buy weapon 2" Then
wea = 2
whspnum = 15
Else
If Msg Like "buy weapon 3" Then
wea = 3
whspnum = 15
Else
If Msg Like "buy weapon 4" Then
wea = 4
whspnum = 15
Else
If Msg Like "buy weapon 5" Then
wea = 5
whspnum = 15
Else
If Msg Like "buy weapon 6" Then
wea = 6
whspnum = 15
Else
If Msg Like "buy weapon 7" Then
wea = 7
whspnum = 15
Else
If Msg Like "buy weapon 8" Then
wea = 8
whspnum = 15
Else
If Msg Like "buy weapon 9" Then
wea = 9
whspnum = 15
Else
If Msg Like "buy weapon 10" Then
wea = 10
whspnum = 15
Else
If Msg Like "buy weapon 11" Then
wea = 11
whspnum = 15
Else
If Msg Like "buy weapon 12" Then
wea = 12
whspnum = 15
Else
If Msg Like "buy weapon 13" Then
wea = 13
whspnum = 15
Else
If Msg Like "buy weapon 14" Then
wea = 14
whspnum = 15
Else
If Msg Like "buy weapon 15" Then
wea = 15
whspnum = 15
Else
If Msg Like "buy weapon 16" Then
wea = 16
whspnum = 15


Else
If Msg Like "buy armor 1" Then
wea = 1
whspnum = 16
Else
If Msg Like "buy armor 2" Then
wea = 2
whspnum = 16
Else
If Msg Like "buy armor 3" Then
wea = 3
whspnum = 16
Else
If Msg Like "buy armor 4" Then
wea = 4
whspnum = 16
Else
If Msg Like "buy armor 5" Then
wea = 5
whspnum = 16
Else
If Msg Like "buy armor 6" Then
wea = 6
whspnum = 16
Else
If Msg Like "buy armor 7" Then
wea = 7
whspnum = 16
Else
If Msg Like "buy armor 8" Then
wea = 8
whspnum = 16
Else
If Msg Like "buy" Then
whspnum = 17
Else
If Msg Like "weapons" Then
whspnum = 18
Else
If Msg Like "armor" Then
whspnum = 19
Else
whspnum = 0
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
If whspnum = 0 Then sckFurc.SendData "wh " & Furre & " I dont understand. Try Whispering me help." & vbLf
If whspnum = 1 Then sckFurc.SendData "wh " & Furre & " Commands, JOIN, FIGHTING, STATS, BAR, BUY." & vbLf
If whspnum = 2 Then dostats Furre
If whspnum = 3 Then joinfight Furre
If whspnum = 4 Then sckFurc.SendData "wh " & Furre & " To join you must first chose a class. Classes are Fighter, Wizard, Thief, Paladin, and Priest. Whisper me a class name to lern more. After you have chosen a class whisper me JOIN CLASS" & vbLf
If whspnum = 5 Then doregister Furre, clas
If whspnum = 6 Then sckFurc.SendData "wh " & Furre & " This mighty Fighter is a brave adventurer and good friend. He defends his companions against monsters and outher enemies with fierce devotion." & vbLf
If whspnum = 7 Then sckFurc.SendData "wh " & Furre & " This powerful Wizard controls vast magical energies, shaping them and casting them as mighty spells. He studies strange tongues and obscure facts and devotes much of his time to magical research." & vbLf
If whspnum = 8 Then sckFurc.SendData "wh " & Furre & " This cunning Thief makes his way through the world using his wits, stealth, and roguish talents. His companions depend on his skills to aid them in avoiding locks, traps, and outher hidden dangers." & vbLf
If whspnum = 9 Then sckFurc.SendData "wh " & Furre & " The Paladin. This holy warrior stands pure and true against the evils of the world. He upholds all that is good, living for the ideals of righteousness, justice, honesty, and chivalry. He strives to be a liveing example of these virtues so that outhers might learn from him as wall as gain by his actions." & vbLf
If whspnum = 10 Then sckFurc.SendData "wh " & Furre & " The Priest serves as a protector and healer for his companions. When evil threatens, he woun't hesitate to hunt it down and destroy it. He calls upon the power of his faith to cast powerful spells to aid his allies and distory his enemies." & vbLf
    If chkTim = Checked Then fighting = "Timed, Attacks may be made eaven if outher Furre hasent attacked. 7 sec delay between attacks."
    If chkTurn = Checked Then fighting = "Take Turns, Furre's must take turns between attacks."
    If chkDream = Checked Then fighting = "In Dream, Jovati anamates all fighting."
If whspnum = 11 Then sckFurc.SendData "wh " & Furre & " Fighting is curenty set as " & fighting & vbLf
If whspnum = 12 Then sckFurc.SendData "wh " & Furre & " If fighting is set to times or turns. attacks may be made by sending a @ to Furcadia. Furres must join the fight by whispering me JOIN FIGHT. If fight is set to dream Furres must move into the fighting bot for the fight to begine." & vbLf
    If chkbar.Value = Checked Then bara = "Enabeled"
    If chkbar.Value = Unchecked Then bara = "Disabeled"
If whspnum = 13 Then sckFurc.SendData "wh " & Furre & " To order something type #ITEM# replaceing ITEM with anything you want Whisper me MENU for a list of what I can make. Bartender is curenty " & bara & vbLf
If whspnum = 14 Then sckFurc.SendData "wh " & Furre & " Food: HOTDOG HAMBURGER. Drinks: BEER ROOTBEER" & vbLf
If whspnum = 15 Then buyweapon Furre, wea
If whspnum = 16 Then buyarmo Furre, wea
If whspnum = 17 Then sckFurc.SendData "wh " & Furre & " To buy weapons and armor whisper me BUY WEAPON # or BUY ARMOR # replaceing # with weapon or armor number. For lists whisper WEAPON or ARMOR to me. Items are 10X the item number in Gold." & vbLf
If whspnum = 18 Then sckFurc.SendData "wh " & Furre & " Weapons: [Dagger - 1] [Knife - 2] [Hand ax - 3] [Quarterstaff - 4] [Spear - 5] [Warhammer - 6] [Battle ax - 7] [Morneing Star - 8] [Flail - 9] [Mace - 10] [Broad Sword - 11] [Short Bow - 12] [Crossbow - 13] [Shord Sword - 14] [Long Sword - 15] [TwoHand Sword - 16]." & vbLf
If whspnum = 19 Then sckFurc.SendData "wh " & Furre & " Armor: [Padded - 1] [Leather - 2] [Chain Mail - 3] [Splint Mail - 4] [Ring Mail - 5] [Scale Mail - 6] [Banded Mail - 7] [Plate Mail - 8]." & vbLf
End Sub
Sub buyweapon(Furre, wea)
Open "members.txt" For Input As #1
Input #1, fName, mnum
Do Until (fName = Furre) Or (EOF(1))
Input #1, fName, mnum
Loop
Close #1
If fName = Furre Then
    Open "memfiles\" & mnum & ".txt" For Input As #1
    Input #1, fName, mnum, lvl, clas, gold, xp, weap, armo, sone, stwo, sthe, sfor, sfiv
    Close #1
    price = wea * 10
    If wea < weap Then
        sckFurc.SendData "wh " & Furre & " you would be dumb to buy a less of a weapon." & vbLf
    Else
    If price < gold Then
        ngold = gold - price
        Open "memfiles\" & mnum & ".txt" For Output As #1
        Write #1, fName, mnum, lvl, clas, ngold, xp, wea, armo, sone, stwo, sthe, sfor, sfiv
        Close #1
If wea = 1 Then wea = "Dagger"
If wea = 2 Then wea = "Knife"
If wea = 3 Then wea = "Hand ax"
If wea = 4 Then wea = "Quarterstaff"
If wea = 5 Then wea = "Spear"
If wea = 6 Then wea = "Warhammer"
If wea = 7 Then wea = "Battle ax"
If wea = 8 Then wea = "Morneing Star"
If wea = 9 Then wea = "Flail"
If wea = 10 Then wea = "Mace"
If wea = 11 Then wea = "Broad Sword"
If wea = 12 Then wea = "Short Bow"
If wea = 13 Then wea = "Crossbow"
If wea = 14 Then wea = "Shord Sword"
If wea = 15 Then wea = "Long Sword"
If wea = 16 Then wea = "TwoHand Sword"
        sckFurc.SendData "wh " & Furre & " you now have a " & wea & vbLf
    Else
        sckFurc.SendData "wh " & Furre & " you dont have enuff gold to buy that." & vbLf
    End If
    End If
Else
    sckFurc.SendData "wh " & Furre & " You are not a member." & vbLf
End If
End Sub

Sub buyarmo(Furre, wea)
Open "members.txt" For Input As #1
Input #1, fName, mnum
Do Until (fName = Furre) Or (EOF(1))
Input #1, fName, mnum
Loop
Close #1
If fName = Furre Then
    Open "memfiles\" & mnum & ".txt" For Input As #1
    Input #1, fName, mnum, lvl, clas, gold, xp, weap, armo, sone, stwo, sthe, sfor, sfiv
    Close #1
    price = wea * 10
    If wea < weap Then
        sckFurc.SendData "wh " & Furre & " you would be dumb to buy a less of armor." & vbLf
    Else
    If price < gold Then
        ngold = gold - price
        Open "memfiles\" & mnum & ".txt" For Output As #1
        Write #1, fName, mnum, lvl, clas, ngold, xp, weap, wea, sone, stwo, sthe, sfor, sfiv
        Close #1
If wea = 0 Then wea = "Fir"
If wea = 1 Then wea = "Padded"
If wea = 2 Then wea = "Leather"
If wea = 3 Then wea = "Chain Mail"
If wea = 4 Then wea = "Splint Mail"
If wea = 5 Then wea = "Ring Mail"
If wea = 6 Then wea = "Scale Mail"
If wea = 7 Then wea = "Banded Mail"
If wea = 8 Then wea = "Plate Mail"
        sckFurc.SendData "wh " & Furre & " you now have " & wea & vbLf
    Else
        sckFurc.SendData "wh " & Furre & " you dont have enuff gold to buy that." & vbLf
    End If
    End If
Else
    sckFurc.SendData "wh " & Furre & " You are not a member." & vbLf
End If
End Sub
Sub doserve(Furre, Order)
If Order = "beer" Then sckFurc.SendData ":Pops the top on an ice cold Beer and sends it down the bar to " & Furre & "." & vbLf
If Order = "rootbeer" Then sckFurc.SendData ":opens a bottle of A" & Chr(38) & "W RootBeer and hands it to " & Furre & "." & vbLf
If Order = "hamburger" Then sckFurc.SendData ":Frys up a patty on the grill. Slaps lots of veggies and sauses on it to make it look biger. Puts it in a basket full of soggy frys and hands it to " & Furre & "." & vbLf
If Order = "hotdog" Then sckFurc.SendData ":Pulls the oldest Hotdog off the turning hotdog cooker thing slips it into a dryed out bun and hands it to " & Furre & "." & vbLf
End Sub
Sub dofdream(turn)
anum = Int((5 * Rnd) + 1)
If anum = 1 Then atk = "Punched"
If anum = 2 Then atk = "Kicked"
If anum = 3 Then atk = "Bashed"
If anum = 4 Then atk = "Stabed"
If anum = 5 Then atk = "Slashed"
            Open "turn.txt" For Input As #1
            Input #1, turn
            Close #1

If turn = 1 Then
    Open "fighter1.txt" For Input As #1
    Input #1, fnam, lvl, hp, man, weap, armo, clas
    Close #1
    Open "fighter2.txt" For Input As #1
    Input #1, tfnam, tlvl, thp, tman, tweap, tarmo, tClas
    Close #1
    Open "turn.txt" For Output As #1
    Write #1, 2
    Close #1
    rnum = lvl * 2
    num = Int((rnum * Rnd) + 1)
    hit = num + weap - tarmo
    If hit < 1 Then hit = 1
    miss = Int((rnum * Rnd) + 1)
    
    If miss = hit Then
        sckFurc.SendData Chr(34) & "emit " & fnam & "'s Attack Missed." & vbLf
        Open "turn.txt" For Output As #1
        Write #1, 2
        Close #1
    Else
       sckFurc.SendData Chr(34) & "emit " & fnam & " " & atk & " " & tfnam & " For " & hit & " Points of damage." & vbLf
       nhp = thp - hit
       Open "fighter2.txt" For Output As #1
       Write #1, tfnam, tlvl, nhp, tman, tweap, tarmo, tClass
       Close #1
       If nhp < 1 Then
       dowin fnam, tfnam
       sckFurc.SendData "m 1" & vbLf
       Else
        Open "turn.txt" For Output As #1
        Write #1, 2
        Close #1
       End If
    End If
End If


If turn = 2 Then
    Open "fighter2.txt" For Input As #1
    Input #1, fnam, lvl, hp, man, weap, armo, clas
    Close #1
    Open "fighter1.txt" For Input As #1
    Input #1, tfnam, tlvl, thp, tman, tweap, tarmo, tClas
    Close #1
    Open "turn.txt" For Output As #1
    Write #1, 2
    Close #1
    rnum = lvl * 2
    num = Int((rnum * Rnd) + 1)
    hit = num + weap - tarmo
    If hit < 1 Then hit = 1
    miss = Int((rnum * Rnd) + 1)
    
    If miss = hit Then
        sckFurc.SendData Chr(34) & "emit " & fnam & "'s Attack Missed." & vbLf
        Open "turn.txt" For Output As #1
        Write #1, 1
        Close #1
    Else
       sckFurc.SendData Chr(34) & "emit " & fnam & " " & atk & " " & tfnam & " For " & hit & " Points of damage." & vbLf
       nhp = thp - hit
       Open "fighter1.txt" For Output As #1
       Write #1, tfnam, tlvl, nhp, tman, tweap, tarmo, tClass
       Close #1
       If nhp < 1 Then
       dowin fnam, tfnam
       sckFurc.SendData "m 1" & vbLf
       Else
       Open "turn.txt" For Output As #1
        Write #1, 1
        Close #1
       End If
    End If
End If

End Sub
Private Sub tfight_Timer()
            Open "fnum.txt" For Input As #1
            Input #1, fnum
            Close #1
            Open "turn.txt" For Input As #1
            Input #1, turn
            Close #1
        If chkDream = Checked And fnum = 2 Then
        dofdream turn
        End If
End Sub


Sub doattk(Furre, turn)
anum = Int((5 * Rnd) + 1)
If anum = 1 Then atk = "Punched"
If anum = 2 Then atk = "Kicked"
If anum = 3 Then atk = "Bashed"
If anum = 4 Then atk = "Stabed"
If anum = 5 Then atk = "Slashed"
If turn = 1 Then
    Open "fighter1.txt" For Input As #1
    Input #1, fnam, lvl, hp, man, weap, armo, clas
    Close #1
    Open "fighter2.txt" For Input As #1
    Input #1, tfnam, tlvl, thp, tman, tweap, tarmo, tClas
    Close #1
    Open "turn.txt" For Output As #1
    Write #1, 2
    Close #1
    rnum = lvl * 2
    num = Int((rnum * Rnd) + 1)
    hit = num + weap - tarmo
    If hit < 1 Then hit = 1
    miss = Int((rnum * Rnd) + 1)
    
    If miss = hit Then
        sckFurc.SendData Chr(34) & "emit " & Furre & "'s Attack Missed." & vbLf
    Else
       sckFurc.SendData Chr(34) & "emit " & Furre & " " & atk & " " & tfnam & " For " & hit & " Points of damage." & vbLf
       nhp = thp - hit
       Open "fighter2.txt" For Output As #1
       Write #1, tfnam, tlvl, nhp, tman, tweap, tarmo, tClass
       Close #1
       If nhp < 1 Then dowin fnam, tfnam
    End If
End If
If turn = 2 Then
    Open "fighter2.txt" For Input As #1
    Input #1, fnam, lvl, hp, man, weap, armo, Class
    Close #1
    Open "fighter1.txt" For Input As #1
    Input #1, tfnam, tlvl, thp, tman, tweap, tarmo, tClass
    Close #1
    Open "turn.txt" For Output As #1
    Write #1, 1
    Close #1
    rnum = lvl * 2
    num = Int((rnum * Rnd) + 1)
    hit = num + weap - tarmo
    If hit < 1 Then hit = 1
    miss = Int((rnum * Rnd) + 1)
    If miss = hit Then
        sckFurc.SendData Chr(34) & "emit " & Furre & "'s Attack Missed." & vbLf
    Else
       sckFurc.SendData Chr(34) & "emit " & Furre & " " & atk & " " & tfnam & " For " & hit & " Points of damage." & vbLf
       nhp = thp - hit
       Open "fighter1.txt" For Output As #1
       Write #1, tfnam, tlvl, nhp, tman, tweap, tarmo, tClass
       Close #1
       If nhp < 1 Then dowin fnam, tfnam
    End If
End If

End Sub
Sub dowin(win, lose)
Open "members.txt" For Input As #1
Input #1, fName, mnum
Do Until (fName = win) Or (EOF(1))
Input #1, fName, mnum
Loop
Close #1
If fName = win Then
Open "memfiles\" & mnum & ".txt" For Input As #1
Input #1, fName, mnum, lvl, Class, gold, xp, weap, armo, sone, stwo, sthe, sfor, sfiv
Close #1
If lvl < 10 Then
nxp = xp + 10
End If
If lvl > 10 Then
nxp = xp + 5
End If
ngold = gold + 5

If nxp >= 100 Then
    nxp = 0
    lvl = lvl + 1
    sckFurc.SendData Chr(34) & "emit " & win & " has ganed a lvl." & vbLf
End If
Open "memfiles\" & mnum & ".txt" For Output As #1
Write #1, fName, mnum, lvl, Class, ngold, nxp, weap, armo, sone, stwo, sthe, sfor, sfiv
Close #1
End If

Open "members.txt" For Input As #1
Input #1, fName, mnum
Do Until (fName = lose) Or (EOF(1))
Input #1, fName, mnum
Loop
Close #1
If fName = lose Then
Open "memfiles\" & mnum & ".txt" For Input As #1
Input #1, fName, mnum, lvl, Class, gold, xp, weap, armo, sone, stwo, sthe, sfor, sfiv
Close #1
ngold = gold + 3
If lvl < 10 Then
nxp = xp + 7
End If
If lvl > 10 Then
nxp = xp + 3
End If
If nxp >= 100 Then
    nxp = 0
    lvl = lvl + 1
    sckFurc.SendData Chr(34) & "emit " & lose & " has ganed a lvl." & vbLf
End If
Open "memfiles\" & mnum & ".txt" For Output As #1
Write #1, fName, mnum, lvl, Class, ngold, nxp, weap, armo, sone, stwo, sthe, sfor, sfiv
Close #1
End If

sckFurc.SendData Chr(34) & "emit " & win & " has defeted " & lose & vbLf
Open "fnum.txt" For Output As #1
Write #1, 0
Close #1
Open "fighter1.txt" For Output As #1
Write #1, "na", 0, 0, 0, 0, 0, "na"
Close #1
Open "fighter2.txt" For Output As #1
Write #1, "na", 0, 0, 0, 0, 0, "na"
Close #1
End Sub
Sub joinfight(Furre)
Open "members.txt" For Input As #1
Input #1, fName, mnum
Do Until (fName = Furre) Or (EOF(1))
Input #1, fName, mnum
Loop
Close #1
If fName = Furre Then
Open "memfiles\" & mnum & ".txt" For Input As #1
Input #1, fName, mnum, lvl, clas, gold, xp, weap, armo, sone, stwo, sthe, sfor, sfiv
Close #1
    Open "fnum.txt" For Input As #1
    Input #1, fnum
    Close #1
    
    
    If fnum = 0 Then
        Open "fnum.txt" For Output As #2
        Write #2, 1
        Close #2
    If clas = "Wizard" Then
    hp = lvl * 15
    man = lvl * 20
    End If
    If clas = "Fighter" Then
    hp = lvl * 20
    man = 0
    End If
    If clas = "Thief" Then
    hp = lvl * 15
    man = lvl * 10
    End If
    If clas = "Paladin" Then
    hp = lvl * 20
    man = lvl * 10
    End If
    If clas = "Priest" Then
    hp = lvl * 10
    man = lvl * 10
    End If
        Open "fighter1.txt" For Output As #3
        Write #3, Furre, lvl, hp, man, weap, armo, Class
        Close #3
        timeO = Timer
            If weap = 0 Then weap = "Paws"
If weap = 1 Then weap = "Dagger"
If weap = 2 Then weap = "Knife"
If weap = 3 Then weap = "Hand ax"
If weap = 4 Then weap = "Quarterstaff"
If weap = 5 Then weap = "Spear"
If weap = 6 Then weap = "Warhammer"
If weap = 7 Then weap = "Battle ax"
If weap = 8 Then weap = "Morneing Star"
If weap = 9 Then weap = "Flail"
If weap = 10 Then weap = "Mace"
If weap = 11 Then weap = "Broad Sword"
If weap = 12 Then weap = "Short Bow"
If weap = 13 Then weap = "Crossbow"
If weap = 14 Then weap = "Shord Sword"
If weap = 15 Then weap = "Long Sword"
If weap = 16 Then weap = "TwoHand Sword"

If armo = 0 Then armo = "Fir"
If armo = 1 Then armo = "Padded"
If armo = 2 Then armo = "Leather"
If armo = 3 Then armo = "Chain Mail"
If armo = 4 Then armo = "Splint Mail"
If armo = 5 Then armo = "Ring Mail"
If armo = 6 Then armo = "Scale Mail"
If armo = 7 Then armo = "Banded Mail"
If armo = 8 Then armo = "Plate Mail"
        sckFurc.SendData Chr(34) & "emit " & Furre & " The " & clas & " Welding " & weap & " and " & armo & " has entered the fighting area." & vbLf
    End If

    If fnum = 1 Then
        Open "fnum.txt" For Output As #2
        Write #2, 2
        Close #2
    If clas = "Wizard" Then
    hp = lvl * 15
    man = lvl * 20
    End If
    If clas = "Fighter" Then
    hp = lvl * 20
    man = 0
    End If
    If clas = "Thief" Then
    hp = lvl * 15
    man = lvl * 10
    End If
    If clas = "Paladin" Then
    hp = lvl * 20
    man = lvl * 10
    End If
    If clas = "Priest" Then
    hp = lvl * 10
    man = lvl * 10
    End If
        Open "fighter2.txt" For Output As #3
        Write #3, Furre, lvl, hp, man, weap, armo, Class
        Close #3
        timeT = Timer
            If weap = 0 Then weap = "Paws"
If weap = 1 Then weap = "Dagger"
If weap = 2 Then weap = "Knife"
If weap = 3 Then weap = "Hand ax"
If weap = 4 Then weap = "Quarterstaff"
If weap = 5 Then weap = "Spear"
If weap = 6 Then weap = "Warhammer"
If weap = 7 Then weap = "Battle ax"
If weap = 8 Then weap = "Morneing Star"
If weap = 9 Then weap = "Flail"
If weap = 10 Then weap = "Mace"
If weap = 11 Then weap = "Broad Sword"
If weap = 12 Then weap = "Short Bow"
If weap = 13 Then weap = "Crossbow"
If weap = 14 Then weap = "Shord Sword"
If weap = 15 Then weap = "Long Sword"
If weap = 16 Then weap = "TwoHand Sword"

If armo = 0 Then armo = "Fir"
If armo = 1 Then armo = "Padded"
If armo = 2 Then armo = "Leather"
If armo = 3 Then armo = "Chain Mail"
If armo = 4 Then armo = "Splint Mail"
If armo = 5 Then armo = "Ring Mail"
If armo = 6 Then armo = "Scale Mail"
If armo = 7 Then armo = "Banded Mail"
If armo = 8 Then armo = "Plate Mail"
        sckFurc.SendData Chr(34) & "emit " & Furre & " The " & Class & " Welding " & weap & " and " & armo & " has entered the fighting area." & vbLf
        sckFurc.SendData Chr(34) & "emit Let the fight begine" & vbLf
    End If
    If fnum = 2 Then
        sckFurc.SendData "wh " & Furre & " Please wate intell the fight ends befor joining." & vbLf
    End If
Else
    sckFurc.SendData "wh " & Furre & " You are not a member. Whisper JOIN to me to join." & vbLf
End If
End Sub
Sub doregister(Furre, clas)
Open "members.txt" For Input As #1
Input #1, fName, mnum
Do Until (fName = Furre) Or (EOF(1))
Input #1, fName, mnum
Loop
Close #1

Open "memnum.txt" For Input As #1
Input #1, nnum
Close #1
num = nnum + 1

Open "memnum.txt" For Output As #1
Write #1, num
Close #1

If fName = Furre Then
sckFurc.SendData "wh " & Furre & " You are allready a member." & vbLf
Else
Open "memfiles\" & num & ".txt" For Output As #1
Write #1, Furre, num, 1, clas, 0, 0, 0, 0, 0, 0, 0, 0, 0
Close #1
Open "members.txt" For Append As #1
Write #1, Furre, num
Close #1
sckFurc.SendData "wh " & Furre & " You are now a member." & vbLf
End If


End Sub
Sub dostats(Furre)
Open "members.txt" For Input As #1
Input #1, fName, mnum
Do Until (fName = Furre) Or (EOF(1))
Input #1, fName, mnum
Loop
Close #1
If fName = Furre Then
Open "memfiles\" & mnum & ".txt" For Input As #1
Input #1, fName, mnum, lvl, clas, gold, xp, weap, armo, sone, stwo, sthe, sfor, sfiv
Close #1

    If clas = "Wizard" Then
    stren = 9
    dex = 12
    intel = 18
    wis = 16
    cha = 13
    hp = lvl * 15
    man = lvl * 20
    End If
    If clas = "Fighter" Then
    stren = 15
    dex = 13
    intel = 10
    wis = 9
    cha = 10
    hp = lvl * 20
    man = 0
    End If
    If clas = "Thief" Then
    stren = 11
    dex = 18
    intel = 12
    wis = 10
    cha = 12
    hp = lvl * 15
    man = lvl * 10
    End If
    If clas = "Paladin" Then
    stren = 14
    dex = 12
    intel = 9
    wis = 14
    cha = 17
    hp = lvl * 20
    man = lvl * 10
    End If
    If clas = "Priest" Then
    stren = 14
    dex = 13
    intel = 11
    wis = 17
    cha = 10
    hp = lvl * 10
    man = lvl * 10
    End If
If weap = 0 Then weap = "Paws"
If weap = 1 Then weap = "Dagger"
If weap = 2 Then weap = "Knife"
If weap = 3 Then weap = "Hand ax"
If weap = 4 Then weap = "Quarterstaff"
If weap = 5 Then weap = "Spear"
If weap = 6 Then weap = "Warhammer"
If weap = 7 Then weap = "Battle ax"
If weap = 8 Then weap = "Morneing Star"
If weap = 9 Then weap = "Flail"
If weap = 10 Then weap = "Mace"
If weap = 11 Then weap = "Broad Sword"
If weap = 12 Then weap = "Short Bow"
If weap = 13 Then weap = "Crossbow"
If weap = 14 Then weap = "Shord Sword"
If weap = 15 Then weap = "Long Sword"
If weap = 16 Then weap = "TwoHand Sword"

If armo = 0 Then armo = "Fur"
If armo = 1 Then armo = "Padded"
If armo = 2 Then armo = "Leather"
If armo = 3 Then armo = "Chain Mail"
If armo = 4 Then armo = "Splint Mail"
If armo = 5 Then armo = "Ring Mail"
If armo = 6 Then armo = "Scale Mail"
If armo = 7 Then armo = "Banded Mail"
If armo = 8 Then armo = "Plate Mail"

sckFurc.SendData "wh " & Furre & " Stats: [Member# - " & mnum & "] [Class - " & clas & "] [LvL - " & lvl & "] [Strength - " & stren & "] [Dexterity - " & dex & "] [Intelligence - " & intel & "] [Wisdom - " & wis & "] [Exp - " & xp & "%] [Health - " & hp & "] [Mana - " & man & "] [Gold - " & gold & "] [Charisma - " & cha & "] [Weapon - " & weap & "] [Armor - " & armo & "]" & vbLf
Else
sckFurc.SendData "wh " & Furre & " You are not a member. Whisper me JOIN to learn how to become a member." & vbLf
End If
End Sub


Sub dowalk(whatwalk, lastwalk)
If (lastwalk = "j") Or (lastwalk = "k") Or (lastwalk = "l") Then ' moveing up
    If (whatwalk = "j") Or (whatwalk = "k") Or (whatwalk = "l") Then
        sckFurc.SendData "m 9" & vbLf
        Else ' move left
            If (whatwalk = "f") Or (whatwalk = "g") Or (whatwalk = "h") Then
            sckFurc.SendData "m 9" & vbLf
            sckFurc.SendData "m 7" & vbLf
        Else 'move right
            If (whatwalk = "b") Or (whatwalk = "c") Or (whatwalk = "d") Then
            sckFurc.SendData "m 9" & vbLf
            sckFurc.SendData "m 3" & vbLf
        Else ' move down
            If (whatwalk = "`") Or (whatwalk = "_") Or (whatwalk = "^") Then
            sckFurc.SendData "m 7" & vbLf
            sckFurc.SendData "m 9" & vbLf
            sckFurc.SendData "m 9" & vbLf
            sckFurc.SendData "m 3" & vbLf
        End If
        End If
        End If
    End If
lastwalk = whatwalk
End If
If (lastwalk = "f") Or (lastwalk = "g") Or (lastwalk = "h") Then 'moveing left
        If (whatwalk = "f") Or (whatwalk = "g") Or (whatwalk = "h") Then
        sckFurc.SendData "m 7" & vbLf
                Else 'move up
                    If (whatwalk = "j") Or (whatwalk = "k") Or (whatwalk = "l") Then
                    sckFurc.SendData "m 7" & vbLf
                    sckFurc.SendData "m 9" & vbLf
                Else 'move down
                    If (whatwalk = "`") Or (whatwalk = "_") Or (whatwalk = "^") Then
                    sckFurc.SendData "m 7" & vbLf
                    sckFurc.SendData "m 1" & vbLf
                Else 'move right
                If (whatwalk = "b") Or (whatwalk = "c") Or (whatwalk = "d") Then
                    sckFurc.SendData "m 9" & vbLf
                    sckFurc.SendData "m 7" & vbLf
                    sckFurc.SendData "m 7" & vbLf
                    sckFurc.SendData "m 1" & vbLf
                End If
                End If
                End If
        End If
lastwalk = whatwalk
End If
If (lastwalk = "b") Or (lastwalk = "c") Or (lastwalk = "d") Then 'moveing right
        If (whatwalk = "b") Or (whatwalk = "c") Or (whatwalk = "d") Then
        sckFurc.SendData "m 3" & vbLf
                Else 'move up
                    If (whatwalk = "j") Or (whatwalk = "k") Or (whatwalk = "l") Then
                    sckFurc.SendData "m 3" & vbLf
                    sckFurc.SendData "m 9" & vbLf
                Else 'move down
                    If (whatwalk = "`") Or (whatwalk = "_") Or (whatwalk = "^") Then
                    sckFurc.SendData "m 3" & vbLf
                    sckFurc.SendData "m 1" & vbLf
                Else 'move left
                    If (whatwalk = "f") Or (whatwalk = "g") Or (whatwalk = "h") Then
                    sckFurc.SendData "m 9" & vbLf
                    sckFurc.SendData "m 3" & vbLf
                    sckFurc.SendData "m 3" & vbLf
                    sckFurc.SendData "m 1" & vbLf
                End If
                End If
                End If
        End If
lastwalk = whatwalk
End If
If (lastwalk = "`") Or (lastwalk = "_") Or (lastwalk = "^") Then 'moveing down
        If (whatwalk = "`") Or (whatwalk = "_") Or (whatwalk = "^") Then
        sckFurc.SendData "m 1" & vbLf
                Else 'move left
                    If (whatwalk = "f") Or (whatwalk = "g") Or (whatwalk = "h") Then
                    sckFurc.SendData "m 1" & vbLf
                    sckFurc.SendData "m 7" & vbLf
                Else 'move right
                    If (whatwalk = "b") Or (whatwalk = "c") Or (whatwalk = "d") Then
                    sckFurc.SendData "m 1" & vbLf
                    sckFurc.SendData "m 3" & vbLf
                Else 'move up
                    If (whatwalk = "j") Or (whatwalk = "k") Or (whatwalk = "l") Then
                    sckFurc.SendData "m 7" & vbLf
                    sckFurc.SendData "m 1" & vbLf
                    sckFurc.SendData "m 1" & vbLf
                    sckFurc.SendData "m 3" & vbLf
                End If
                End If
                End If
        End If
lastwalk = whatwalk
End If
If (lastwalk = "none") Then
        If (whatwalk = "`") Or (whatwalk = "_") Or (whatwalk = "^") Then
            sckFurc.SendData "m 1" & vbLf
        Else 'move left
            If (whatwalk = "f") Or (whatwalk = "g") Or (whatwalk = "h") Then
            sckFurc.SendData "m 7" & vbLf
        Else 'move right
            If (whatwalk = "b") Or (whatwalk = "c") Or (whatwalk = "d") Then
            sckFurc.SendData "m 3" & vbLf
        Else 'move left
            If (whatwalk = "j") Or (whatwalk = "k") Or (whatwalk = "l") Then
            sckFurc.SendData "m 9" & vbLf
        End If
        End If
        End If
        End If
lastwalk = whatwalk
End If
End Sub


Private Sub cmdExit_Click()
'When you click the Exit button, the bot program is closed.
End
End Sub

Private Sub StayOnline_Timer()
'Each minute the timer is set off. The Minute variable is increased by one. Your
'bot changes its desc to add the Minute which is an Uptimer.
Minute = Minute + 1
sckFurc.SendData "desc " & descrip & " [Uptime: " & Minute & " Minute(s)]" & vbLf
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
