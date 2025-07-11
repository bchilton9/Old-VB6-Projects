VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmBot 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Senn"
   ClientHeight    =   4770
   ClientLeft      =   150
   ClientTop       =   150
   ClientWidth     =   5355
   Icon            =   "frmBot.frx":0000
   LinkTopic       =   "Form1"
   MouseIcon       =   "frmBot.frx":0442
   ScaleHeight     =   4770
   ScaleWidth      =   5355
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdturnl 
      Caption         =   "Turn >"
      Height          =   495
      Left            =   2040
      TabIndex        =   19
      Top             =   4200
      Width           =   615
   End
   Begin VB.CommandButton cmdturnr 
      Caption         =   "< Turn"
      Height          =   495
      Left            =   1440
      TabIndex        =   18
      Top             =   4200
      Width           =   615
   End
   Begin VB.CommandButton cmdGoVinca 
      Caption         =   "&Vinca"
      Height          =   375
      Left            =   2760
      TabIndex        =   15
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton cmdGoAlleg 
      Caption         =   "&Allegria"
      Height          =   375
      Left            =   2760
      TabIndex        =   14
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      Caption         =   "System"
      Height          =   1095
      Left            =   4080
      TabIndex        =   12
      Top             =   240
      Width           =   1215
      Begin VB.CheckBox chkServtxt 
         Caption         =   "SText"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkWhisp 
         Caption         =   "Whispers"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CheckBox chkServCode 
         Caption         =   "SCode"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   975
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
Const BotName = "xxxxxx"
Const BotPass = "xxxxxx"
Const descrip = "This Priest serves as a protector and healer for her companions. When evil threatens, she woun't hesitate to hunt it down and destroy it. She calls upon the power of his faith to cast powerful spells to aid her allies and distory her enemies."
Const ColorCode = "!8J'+999=9 #! #!"
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
sckFurc.SendData "m 9" & vbLf & "use" & vbLf & ">" & vbLf
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


'###########SIGNS############
'Sign 1

If Left(Txt, 2) = "7 " Then
ine = Right(Txt, 10 - 4)
inet = Left(ine, 10 - 8)
If inet = " !" Then
sckFurc.SendData "l  9 J" & vbLf
Sign = 3
End If
    SiLo = Right(Txt, 3)
    sckFurc.SendData "l  " & SiLo & vbLf
'main ; J: J: H
If (SiLo = ": H") Then Sign = 1
If (SiLo = ": J") Then Sign = 1
If (SiLo = "; J") Then Sign = 1
'1t10 fight < >; >; <
If (SiLo = "; <") Then Sign = 2
If (SiLo = "; >") Then Sign = 2
If (SiLo = "< >") Then Sign = 2
End If
'If (SiLo = "") Then Sign = 0
'Gets the furres name when the bot looks at them.
If Left(Txt, 10) = "((You see " Then
    Furre = Mid(Txt, 11, Len(Txt) - 12)
    DoSign Furre
End If



End Sub
Sub DoSign(Furre)
Open "C:\Jovati\memnum.txt" For Input As #1
Input #1, mnum
Close #1

    If (Sign = 0) Then sckFurc.SendData "<" & vbLf & ">" & vbLf
    If (Sign = 1) Then sckFurc.SendData "wh " & Furre & " Krabice City Limit. Populashion " & mnum & vbLf
    If (Sign = 2) Then sckFurc.SendData "wh " & Furre & " Fighting arena. To learn how to fight whisper FIGHT to Senn." & vbLf
    If (Sign = 3) Then sckFurc.SendData "wh " & Furre & " Welcome to Krabice. If you need help whisper HELP to Senn." & vbLf
    Sign = 0
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
If Msg Like "na" Then
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
If Msg Like "naa" Then
whspnum = 11
Else
If Msg Like "fight" Then
whspnum = 12
Else
If Msg Like "naaa" Then
whspnum = 13
Else
If Msg Like "naaaa" Then
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
If Msg Like "weapon" Then
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
If whspnum = 1 Then sckFurc.SendData "wh " & Furre & " Commands, JOIN, FIGHT, STATS, BUY." & vbLf
If whspnum = 2 Then dostats Furre
If whspnum = 3 Then ckFurc.SendData "wh " & Furre & " I dont understand. Try Whispering me help." & vbLf
If whspnum = 4 Then sckFurc.SendData "wh " & Furre & " To join you must first chose a class. Classes are Fighter, Wizard, Thief, Paladin, and Priest. Whisper me a class name to lern more. After you have chosen a class whisper me JOIN CLASS" & vbLf
If whspnum = 5 Then doregister Furre, clas
If whspnum = 6 Then sckFurc.SendData "wh " & Furre & " The mighty Fighter is a brave adventurer and good friend. He defends his companions against monsters and outher enemies with fierce devotion." & vbLf
If whspnum = 7 Then sckFurc.SendData "wh " & Furre & " The powerful Wizard controls vast magical energies, shaping them and casting them as mighty spells. He studies strange tongues and obscure facts and devotes much of his time to magical research." & vbLf
If whspnum = 8 Then sckFurc.SendData "wh " & Furre & " The cunning Thief makes his way through the world using his wits, stealth, and roguish talents. His companions depend on his skills to aid them in avoiding locks, traps, and outher hidden dangers." & vbLf
If whspnum = 9 Then sckFurc.SendData "wh " & Furre & " The Paladin. This holy warrior stands pure and true against the evils of the world. He upholds all that is good, living for the ideals of righteousness, justice, honesty, and chivalry. He strives to be a liveing example of these virtues so that outhers might learn from him as wall as gain by his actions." & vbLf
If whspnum = 10 Then sckFurc.SendData "wh " & Furre & " The Priest serves as a protector and healer for his companions. When evil threatens, he woun't hesitate to hunt it down and destroy it. He calls upon the power of his faith to cast powerful spells to aid his allies and distory his enemies." & vbLf
If whspnum = 11 Then sckFurc.SendData "wh " & Furre & " I dont understand. Try Whispering me help." & vbLf
If whspnum = 12 Then sckFurc.SendData "wh " & Furre & " Fighting as all anomated. Move into an open sparing ring and let the fight begine." & vbLf
If whspnum = 13 Then sckFurc.SendData "wh " & Furre & " I dont understand. Try Whispering me help." & vbLf
If whspnum = 14 Then sckFurc.SendData "wh " & Furre & " I dont understand. Try Whispering me help." & vbLf
If whspnum = 15 Then buyweapon Furre, wea
If whspnum = 16 Then buyarmo Furre, wea
If whspnum = 17 Then sckFurc.SendData "wh " & Furre & " To buy weapons and armor whisper me BUY WEAPON # or BUY ARMOR # replaceing # with weapon or armor number. For lists whisper WEAPON or ARMOR to me. Items are 10X the item number in Gold." & vbLf
If whspnum = 18 Then sckFurc.SendData "wh " & Furre & " Weapons: [Dagger - 1] [Knife - 2] [Hand ax - 3] [Quarterstaff - 4] [Spear - 5] [Warhammer - 6] [Battle ax - 7] [Morneing Star - 8] [Flail - 9] [Mace - 10] [Broad Sword - 11] [Short Bow - 12] [Crossbow - 13] [Shord Sword - 14] [Long Sword - 15] [TwoHand Sword - 16]." & vbLf
If whspnum = 19 Then sckFurc.SendData "wh " & Furre & " Armor: [Padded - 1] [Leather - 2] [Chain Mail - 3] [Splint Mail - 4] [Ring Mail - 5] [Scale Mail - 6] [Banded Mail - 7] [Plate Mail - 8]." & vbLf
End Sub
Sub buyweapon(Furre, wea)
Open "C:\Jovati\members.txt" For Input As #1
Input #1, fName, mnum
Do Until (fName = Furre) Or (EOF(1))
Input #1, fName, mnum
Loop
Close #1
If fName = Furre Then
    Open "C:\Jovati\memfiles\" & mnum & ".txt" For Input As #1
    Input #1, fName, mnum, lvl, clas, gold, xp, weap, armo, sone, stwo, sthe, sfor, sfiv
    Close #1
    price = wea * 10
    If wea < weap Then
        sckFurc.SendData "wh " & Furre & " you would be dumb to buy a less of a weapon." & vbLf
    Else
    If price < gold Then
        ngold = gold - price
        Open "C:\Jovati\memfiles\" & mnum & ".txt" For Output As #1
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
Open "C:\Jovati\members.txt" For Input As #1
Input #1, fName, mnum
Do Until (fName = Furre) Or (EOF(1))
Input #1, fName, mnum
Loop
Close #1
If fName = Furre Then
    Open "C:\Jovati\memfiles\" & mnum & ".txt" For Input As #1
    Input #1, fName, mnum, lvl, clas, gold, xp, weap, armo, sone, stwo, sthe, sfor, sfiv
    Close #1
    price = wea * 10
    If wea < weap Then
        sckFurc.SendData "wh " & Furre & " you would be dumb to buy a less of armor." & vbLf
    Else
    If price < gold Then
        ngold = gold - price
        Open "C:\Jovati\memfiles\" & mnum & ".txt" For Output As #1
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

Sub doregister(Furre, clas)
Open "C:\Jovati\members.txt" For Input As #1
Input #1, fName, mnum
Do Until (fName = Furre) Or (EOF(1))
Input #1, fName, mnum
Loop
Close #1

Open "C:\Jovati\memnum.txt" For Input As #1
Input #1, nnum
Close #1
num = nnum + 1

Open "C:\Jovati\memnum.txt" For Output As #1
Write #1, num
Close #1

If fName = Furre Then
sckFurc.SendData "wh " & Furre & " You are allready a member." & vbLf
Else
Open "C:\Jovati\memfiles\" & num & ".txt" For Output As #1
Write #1, Furre, num, 1, clas, 0, 0, 0, 0, 0, 0, 0, 0, 0
Close #1
Open "C:\Jovati\members.txt" For Append As #1
Write #1, Furre, num
Close #1
sckFurc.SendData "wh " & Furre & " You are now a member." & vbLf
End If


End Sub
Sub dostats(Furre)
Open "C:\Jovati\members.txt" For Input As #1
Input #1, fName, mnum
Do Until (fName = Furre) Or (EOF(1))
Input #1, fName, mnum
Loop
Close #1
If fName = Furre Then
Open "C:\Jovati\memfiles\" & mnum & ".txt" For Input As #1
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

