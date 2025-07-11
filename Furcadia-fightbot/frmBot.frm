VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmBot 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Valka (Fight Bot)"
   ClientHeight    =   2190
   ClientLeft      =   150
   ClientTop       =   450
   ClientWidth     =   4095
   Icon            =   "frmBot.frx":0000
   LinkTopic       =   "Form1"
   MouseIcon       =   "frmBot.frx":0442
   ScaleHeight     =   2190
   ScaleWidth      =   4095
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmsSend 
      Caption         =   "Send"
      Height          =   255
      Left            =   3360
      TabIndex        =   56
      Top             =   1800
      Width           =   615
   End
   Begin VB.Frame Frame3 
      Caption         =   "member stuff"
      Height          =   3735
      Left            =   120
      TabIndex        =   7
      Top             =   2400
      Width           =   4095
      Begin VB.TextBox txtCoin 
         DataField       =   "Coin"
         DataSource      =   "members"
         Height          =   285
         Left            =   3120
         TabIndex        =   58
         Top             =   2880
         Width           =   855
      End
      Begin VB.Data members 
         Caption         =   "Members"
         Connect         =   "Access"
         DatabaseName    =   "C:\Documents and Settings\Byron Chilton\My Documents\vbprojects\fightbot\members.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   360
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "members"
         Top             =   240
         Width           =   3615
      End
      Begin VB.TextBox txtNum 
         DataField       =   "ID"
         DataSource      =   "members"
         Height          =   285
         Left            =   1200
         TabIndex        =   19
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox txtName 
         DataField       =   "Name"
         DataSource      =   "members"
         Height          =   285
         Left            =   1200
         TabIndex        =   18
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox txtLvl 
         DataField       =   "Level"
         DataSource      =   "members"
         Height          =   285
         Left            =   1200
         TabIndex        =   17
         Top             =   1440
         Width           =   495
      End
      Begin VB.TextBox txtExp 
         DataField       =   "Exp"
         DataSource      =   "members"
         Height          =   285
         Left            =   1200
         TabIndex        =   16
         Top             =   1800
         Width           =   495
      End
      Begin VB.TextBox txtArmor 
         DataField       =   "Armor"
         DataSource      =   "members"
         Height          =   285
         Left            =   1200
         TabIndex        =   15
         Top             =   2160
         Width           =   495
      End
      Begin VB.TextBox txtWeapon 
         DataField       =   "Weapon"
         DataSource      =   "members"
         Height          =   285
         Left            =   1200
         TabIndex        =   14
         Top             =   2520
         Width           =   495
      End
      Begin VB.TextBox txtMana 
         DataField       =   "Mana"
         DataSource      =   "members"
         Height          =   285
         Left            =   3120
         TabIndex        =   13
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox txtItem1 
         DataField       =   "Item1"
         DataSource      =   "members"
         Height          =   285
         Left            =   3120
         TabIndex        =   12
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox txtItem2 
         DataField       =   "Item2"
         DataSource      =   "members"
         Height          =   285
         Left            =   3120
         TabIndex        =   11
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox txtItem3 
         DataField       =   "Item3"
         DataSource      =   "members"
         Height          =   285
         Left            =   3120
         TabIndex        =   10
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox txtItem4 
         DataField       =   "Item4"
         DataSource      =   "members"
         Height          =   285
         Left            =   3120
         TabIndex        =   9
         Top             =   2160
         Width           =   855
      End
      Begin VB.TextBox txtItem5 
         DataField       =   "Item5"
         DataSource      =   "members"
         Height          =   285
         Left            =   3120
         TabIndex        =   8
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label Label22 
         Caption         =   "Coin"
         Height          =   255
         Left            =   2160
         TabIndex        =   57
         Top             =   2880
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Level"
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Number"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Name"
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Exp"
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Armor"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Weapon"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Mana"
         Height          =   255
         Left            =   2160
         TabIndex        =   25
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "Item 1"
         Height          =   255
         Left            =   2160
         TabIndex        =   24
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label10 
         Caption         =   "Item 2"
         Height          =   255
         Left            =   2160
         TabIndex        =   23
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label11 
         Caption         =   "Item 3"
         Height          =   255
         Left            =   2160
         TabIndex        =   22
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label12 
         Caption         =   "Item 4"
         Height          =   255
         Left            =   2160
         TabIndex        =   21
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Label13 
         Caption         =   "Item 5"
         Height          =   255
         Left            =   2160
         TabIndex        =   20
         Top             =   2520
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Fight Stuff"
      Height          =   4815
      Left            =   4920
      TabIndex        =   2
      Top             =   240
      Width           =   4095
      Begin VB.TextBox txtTurn 
         Height          =   375
         Left            =   1920
         TabIndex        =   54
         Text            =   "a"
         Top             =   4200
         Width           =   1215
      End
      Begin VB.CheckBox chkFight 
         Caption         =   "Check1"
         Height          =   255
         Left            =   1560
         TabIndex        =   53
         Top             =   840
         Width           =   255
      End
      Begin VB.Timer timFight 
         Interval        =   3000
         Left            =   120
         Top             =   4080
      End
      Begin VB.TextBox fb 
         Height          =   375
         Index           =   7
         Left            =   2520
         TabIndex        =   51
         Top             =   3600
         Width           =   1335
      End
      Begin VB.TextBox fa 
         Height          =   375
         Index           =   7
         Left            =   120
         TabIndex        =   50
         Top             =   3600
         Width           =   1335
      End
      Begin VB.TextBox fa 
         Height          =   375
         Index           =   6
         Left            =   120
         TabIndex        =   48
         Top             =   3120
         Width           =   1335
      End
      Begin VB.TextBox fb 
         Height          =   375
         Index           =   6
         Left            =   2520
         TabIndex        =   47
         Top             =   3120
         Width           =   1335
      End
      Begin VB.TextBox fb 
         Height          =   375
         Index           =   5
         Left            =   2520
         TabIndex        =   45
         Top             =   2640
         Width           =   1335
      End
      Begin VB.TextBox fa 
         Height          =   375
         Index           =   5
         Left            =   120
         TabIndex        =   44
         Top             =   2640
         Width           =   1335
      End
      Begin VB.TextBox fb 
         Height          =   375
         Index           =   4
         Left            =   2520
         TabIndex        =   43
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox fb 
         Height          =   375
         Index           =   3
         Left            =   2520
         TabIndex        =   42
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox fb 
         Height          =   375
         Index           =   2
         Left            =   2520
         TabIndex        =   41
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox fb 
         Height          =   375
         Index           =   1
         Left            =   2520
         TabIndex        =   40
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox fa 
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   39
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox fa 
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   38
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox fa 
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   37
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox fa 
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   36
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox fa 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox fb 
         Height          =   375
         Index           =   0
         Left            =   2520
         TabIndex        =   4
         Top             =   240
         Width           =   1335
      End
      Begin VB.CheckBox cfa 
         Caption         =   "Check1"
         Height          =   255
         Left            =   1560
         TabIndex        =   3
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label21 
         Caption         =   "turn"
         Height          =   255
         Left            =   1320
         TabIndex        =   55
         Top             =   4320
         Width           =   855
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         Caption         =   "Attack"
         Height          =   255
         Left            =   1680
         TabIndex        =   52
         Top             =   3720
         Width           =   735
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         Caption         =   "HP"
         Height          =   255
         Left            =   1680
         TabIndex        =   49
         Top             =   3240
         Width           =   735
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         Caption         =   "Side"
         Height          =   255
         Left            =   1680
         TabIndex        =   46
         Top             =   2760
         Width           =   735
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         Caption         =   "Mana"
         Height          =   255
         Left            =   1680
         TabIndex        =   35
         Top             =   2280
         Width           =   735
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Caption         =   "Weapon"
         Height          =   255
         Left            =   1680
         TabIndex        =   34
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         Caption         =   "Armor"
         Height          =   255
         Left            =   1680
         TabIndex        =   33
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Caption         =   "Level"
         Height          =   255
         Left            =   1680
         TabIndex        =   32
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Vs."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   6
         Top             =   360
         Width           =   375
      End
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
      Top             =   1800
      Width           =   3255
   End
   Begin VB.TextBox txtFromFurc 
      Height          =   1575
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   3855
   End
   Begin VB.Menu file 
      Caption         =   "File"
      Begin VB.Menu cnt 
         Caption         =   "Connect"
      End
      Begin VB.Menu discon 
         Caption         =   "Disconnect"
      End
      Begin VB.Menu exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu cont 
      Caption         =   "Controls"
   End
   Begin VB.Menu system 
      Caption         =   "System"
      Begin VB.Menu stext 
         Caption         =   "Server Text"
         Checked         =   -1  'True
      End
      Begin VB.Menu scode 
         Caption         =   "Server Code"
         Checked         =   -1  'True
      End
      Begin VB.Menu whsp 
         Caption         =   "Whispers"
         Checked         =   -1  'True
      End
   End
End
Attribute VB_Name = "frmBot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const BotName = "Valka"
Const BotPass = "0519aa"
Const descript = "This mighty Fighter is a brave adventurer and good friend. He defends his companions against monsters and outher enemies with fierce devotion."
Const ColorCode = "5583+88888!#!"
Public Minute As Integer
Public Desc As String
Public Connected As Boolean
Dim side As String

Private Sub cmsSend_Click()
    sckFurc.SendData txtSend & vbLf
    txtSend = ""
End Sub

Private Sub cnt_Click()
If Connected = False Then
sckFurc.RemoteHost = "66.28.224.193"
sckFurc.RemotePort = "6000"
sckFurc.Connect
Sign = 0
Connected = True
End If
End Sub

Private Sub cont_Click()
frmMove.Show
End Sub

Private Sub discon_Click()
If Connected = True Then
sckFurc.Close
Connected = False
End If
End Sub

Private Sub exit_Click()
msg = MsgBox("Are you sure you wish to exit?", vbOKCancel)
If msg = vbOK Then
End
End If
End Sub

Sub Form_Load()
scode.Checked = False
Minute = 0
Desc = descript & " [Uptime: 0 Minute(s)]"
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
On Error Resume Next
'If the checkbox with the Server Code is checked then you see all of the server
'code
If stext.Checked = True Or stext.Enabled = False Then
If scode.Checked = True Then txtFromFurc = txtFromFurc & Txt & vbCrLf
'If the checkbox with the Server Code label is not checked you do not see any of
'the server code. You'll only see what you would see in the Furcadia client.
If scode.Checked = False Then
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
sckFurc.SendData "m 9" & vbLf & "m 7" & vbLf & "m 9" & vbLf & "m 9" & vbLf & "use" & vbLf & "m 1" & vbLf
End If

If Left(Txt, 2) = "7 " And Right(Txt, 3) = "+ :" Then
sckFurc.SendData "l  , 8" & vbLf
side = "right"
End If
If Left(Txt, 2) = "7 " And Right(Txt, 3) = "* 8" Then
sckFurc.SendData "l  + 6" & vbLf
side = "left"
End If

'(Ske: reset arena
If Right(Txt, 13) = ": RESET ARENA" Or Right(Txt, 13) = ": reset arena" Or Right(Txt, 13) = ": Reset Arena" Then
nam = Right(Txt, Len(Txt) - 1)
nam = Left(nam, Len(nam) - 13)
If nam = fa(0).Text Or nam = fb(0).Text Then
cfa = 0
chkFight = 0
txtTurn.Text = "a"

fa(0) = ""
fa(1) = ""
fa(2) = ""
fa(3) = ""
fa(4) = ""
fa(5) = ""
fa(6) = ""
fa(7) = ""

fb(0) = ""
fb(1) = ""
fb(2) = ""
fb(3) = ""
fb(4) = ""
fb(5) = ""
fb(6) = ""
fb(7) = ""

sckFurc.SendData Chr(34) & "Arena has been reset by " & nam & vbLf & "m 1" & vbLf
Else
sckFurc.SendData "wh " & nam & " You are not in the fight! Please do not reset someone's fight." & vbLf
End If
End If


If Left(Txt, 9) = "((You see" Then
Txt = Right(Txt, Len(Txt) - 10)
Txt = Left(Txt, Len(Txt) - 2)


members.Recordset.MoveFirst
Do Until txtName.Text = Txt Or members.Recordset.EOF
members.Recordset.MoveNext
Loop

If txtName.Text <> Txt Then
sckFurc.SendData Chr(34) & Txt & " is not registered." & vbLf & "m 1" & vbLf
cfa = 0
chkFight = 0
txtTurn.Text = "a"

fa(0) = ""
fa(1) = ""
fa(2) = ""
fa(3) = ""
fa(4) = ""
fa(5) = ""
fa(6) = ""
fa(7) = ""

fb(0) = ""
fb(1) = ""
fb(2) = ""
fb(3) = ""
fb(4) = ""
fb(5) = ""
fb(6) = ""
fb(7) = ""

Else
If cfa = 1 Then
fb(0).Text = txtName.Text
fb(1).Text = txtLvl.Text
fb(2).Text = txtArmor.Text
fb(3).Text = txtWeapon.Text
fb(4).Text = txtMana.Text
fb(5).Text = side
fb(6).Text = (fb(1) * 10) + fb(2)
fb(7).Text = (fb(1) * 5) + fb(3)

sckFurc.SendData Chr(34) & "emit " & Txt & " at Level " & fb(1) & " has entered the Arena!" & vbLf
sckFurc.SendData Chr(34) & "emit " & "Let the fight begin!" & vbLf
chkFight = 1
Else


fa(0).Text = txtName.Text
fa(1) = txtLvl.Text
fa(2) = txtArmor.Text
fa(3).Text = txtWeapon.Text
fa(4).Text = txtMana.Text
fa(5).Text = side
fa(6).Text = (fa(1) * 10) + fa(2)
fa(7).Text = (fa(1) * 5) + fa(3)

cfa = 1
sckFurc.SendData Chr(34) & "emit " & Txt & " at Level " & fa(1) & " has entered the Arena!" & vbLf
End If
End If


End If

If whsp.Checked = True Then
'When someone whispers the bot, it gets there name and message and calls the
'DoWhisper(Furre, Msg) sub which is used to respond to whispers.
If Left(Txt, 3) = "([ " And Right(Txt, 10) = " to you. ]" Then
    Tmsg = Split(Txt, " whispers, " & Chr(34), 2)
    Furre = Right(Tmsg(0), Len(Tmsg(0)) - 3)
    msg = Left(Tmsg(1), Len(Tmsg(1)) - 11)
    DoWhisper Furre, msg
End If
End If
End Sub

Private Sub scode_Click()
If scode.Checked = False Then
scode.Checked = True
stext.Enabled = False

ElseIf scode.Checked = True Then
scode.Checked = False
stext.Enabled = True
End If
End Sub

Private Sub stext_Click()
If stext.Checked = False Then
stext.Checked = True
ElseIf stext.Checked = True Then
stext.Checked = False
End If
End Sub

Private Sub timFight_Timer()
On Error Resume Next
If chkFight = 1 Then
If fa(6) <= 0 Or fb(6) <= 0 Then
doWin

Else
If txtTurn.Text = "a" Then
Randomize Timer
farnd = Int((fa(7) * Rnd) + 1)
fb(6) = fb(6) - farnd
sckFurc.SendData Chr(34) & "emit " & fa(0) & " hits " & fb(0) & " for " & farnd & " Points of Damage!" & vbLf
txtTurn.Text = "b"
Else
Randomize Timer
fbrnd = Int((fb(7) * Rnd) + 1)
fa(6) = fa(6) - fbrnd
sckFurc.SendData Chr(34) & "emit " & fb(0) & " hits " & fa(0) & " for " & fbrnd & " Points of Damage!" & vbLf
txtTurn.Text = "a"
End If
End If

End If
End Sub

Sub doWin()
On Error Resume Next
If fa(6) <= 0 Then
    sckFurc.SendData Chr(34) & "emit " & fb(0) & " Has defeted " & fa(0) & " in a battle of wit's!" & vbLf
    If fa(5) = "left" Then
        sckFurc.SendData "m 7" & vbLf & "m 1" & vbLf
    Else
        sckFurc.SendData "m 3" & vbLf & "m 1" & vbLf
    End If


    members.Recordset.MoveFirst
    Do Until txtName.Text = fa(0).Text Or members.Recordset.EOF
        members.Recordset.MoveNext
    Loop

    If txtName.Text = fa(0) Then
        members.Recordset.Edit
        txtExp = txtExp + 3
        If txtExp >= 100 Then
            txtExp = 0
            txtLvl = txtLvl + 1
            sckFurc.SendData "wh " & fa(0) & " Welcome to level " & txtLvl & "." & vbLf
            sckFurc.SendData Chr(34) & "emit " & fa(0) & " Has gained a Level." & vbLf
        End If
        members.Recordset.Update
    End If
    
    members.Recordset.MoveFirst
    Do Until txtName.Text = fb(0).Text Or members.Recordset.EOF
        members.Recordset.MoveNext
    Loop

    If txtName.Text = fb(0) Then
        members.Recordset.Edit
        txtExp = txtExp + 5
        If txtExp >= 100 Then
            txtExp = 0
            txtLvl = txtLvl + 1
            sckFurc.SendData "wh " & fb(0) & " Welcome to level " & txtLvl & "." & vbLf
            sckFurc.SendData Chr(34) & "emit " & fb(0) & " Has gained a Level." & vbLf
        End If
        members.Recordset.Update
    End If
    

ElseIf fb(6) <= 0 Then
    sckFurc.SendData Chr(34) & "emitloud " & fa(0) & " Has defeted " & fb(0) & " in a battle of wit's!" & vbLf
    If fb(5) = "left" Then
        sckFurc.SendData "m 7" & vbLf & "m 1" & vbLf
    Else
        sckFurc.SendData "m 3" & vbLf & "m 1" & vbLf
    End If

    members.Recordset.MoveFirst
    Do Until txtName.Text = fa(0).Text Or members.Recordset.EOF
        members.Recordset.MoveNext
    Loop

    If txtName.Text = fa(0) Then
        members.Recordset.Edit
        txtExp = txtExp + 5
        If txtExp >= 100 Then
            txtExp = 0
            txtLvl = txtLvl + 1
            sckFurc.SendData "wh " & fa(0) & " Welcome to level " & txtLvl & "." & vbLf
            sckFurc.SendData Chr(34) & "emit " & fa(0) & " Has gained a Level." & vbLf
        End If
        members.Recordset.Update
    End If
    
    members.Recordset.MoveFirst
    Do Until txtName.Text = fb(0).Text Or members.Recordset.EOF
        members.Recordset.MoveNext
    Loop

    If txtName.Text = fb(0) Then
        members.Recordset.Edit
        txtExp = txtExp + 3
        If txtExp >= 100 Then
            txtExp = 0
            txtLvl = txtLvl + 1
            sckFurc.SendData "wh " & fb(0) & " Welcome to level " & txtLvl & "." & vbLf
            sckFurc.SendData Chr(34) & "emit " & fb(0) & " Has gained a Level." & vbLf
        End If
        members.Recordset.Update
    End If

End If

cfa = 0
chkFight = 0
txtTurn.Text = "a"

fa(0) = ""
fa(1) = ""
fa(2) = ""
fa(3) = ""
fa(4) = ""
fa(5) = ""
fa(6) = ""
fa(7) = ""

fb(0) = ""
fb(1) = ""
fb(2) = ""
fb(3) = ""
fb(4) = ""
fb(5) = ""
fb(6) = ""
fb(7) = ""

End Sub

Sub DoWhisper(Furre, msg)
done = False
msg = LCase(msg)
'When anyone whispers the bot anything it whispers back
'The Like operator is used for string comparisons


If msg Like "*help*" Then
sckFurc.SendData "wh " & Furre & " Help Commands: JOIN, STATS, FIGHTING, BUYING, SELLING." & vbLf
done = True

ElseIf msg Like "*fighting*" Then
sckFurc.SendData "wh " & Furre & " To fight you must be registered. Then just move into the arena with a chalanger. If you get stuck in the arena say outloud RESET ARENA and i will let you out." & vbLf
done = True

ElseIf msg Like "*stats*" Then
doStats Furre
done = True

ElseIf msg Like "*lock*" And Furre = "Ske" Then
sckFurc.SendData "wh " & Furre & " Arena is Locked." & vbLf & "m 9" & vbLf
done = True

ElseIf msg Like "*unlock*" And Furre = "Ske" Then
sckFurc.SendData "wh " & Furre & " Arena is Unlocked." & vbLf & "m 1" & vbLf
done = True

ElseIf msg Like "*join*" Then
dojoin Furre
done = True

ElseIf msg Like "*buy *" Then
doBuy Furre, msg
done = True

ElseIf msg Like "*sell weapon*" Then
dosellweapon Furre
done = True

ElseIf msg Like "*sell armor*" Then
dosellarmor Furre
done = True

ElseIf msg Like "*armor*" Then
sckFurc.SendData "wh " & Furre & " Armor Type (Price): Rahide (40), Tathered Leather (100), Leather (200), Studed Leather (250), Rusty Chain Mail (300), Chain mail (400), Banded (400), Rusty Steel (550), Steel (600), Fine Steel (1000)." & vbLf
done = True

ElseIf msg Like "*weapon*" Then
sckFurc.SendData "wh " & Furre & " Weapon Type (Price): Rusty Dagger (40), Rusty Axe (100), Heavy Axe (200), Dagger (250), Axe (300), Rusty Sword (400), Fine Steel Dagger (450), Fine Steel Axe (550), Sword (600), Fine Steel Sword (1000)." & vbLf
done = True

ElseIf msg Like "*buying*" Then
sckFurc.SendData "wh " & Furre & " To buy armor and weapons whisper me BUY NAME, replace NAME with the item you want to buy. For a list of Armor whisper me ARMOR, for a list of Weapons whisper me WEAPONS." & vbLf
done = True

ElseIf msg Like "*selling*" Then
sckFurc.SendData "wh " & Furre & " To sell your weapons or armor whisper me SELL WEAPON or SELL ARMOR." & vbLf
done = True

End If

If done = False Then
sckFurc.SendData "wh " & Furre & " I don't understand! Whisper me HELP for more info!" & vbLf
End If
End Sub

Sub dosellweapon(Furre)

Dim price As Integer

members.Recordset.MoveFirst
Do Until txtName.Text = Furre Or members.Recordset.EOF
members.Recordset.MoveNext
Loop



If txtName.Text <> Furre Then
sckFurc.SendData "wh " & Furre & " You are not registered." & vbLf

Else

If txtWeapon.Text <> 0 Then
sckFurc.SendData "wh " & Furre & " Weapon sold" & vbLf

    Select Case txtWeapon.Text
        Case 1
            price = 40
        Case 2
            price = 100
        Case 3
            price = 200
        Case 4
            price = 250
        Case 5
            price = 300
        Case 6
            price = 400
        Case 7
            price = 450
        Case 8
            price = 550
        Case 9
            price = 600
        Case 10
            price = 1000
     End Select

        txtWeapon.Text = "0"
        txtCoin = price + txtCoin
        members.Recordset.Update
        
Else
sckFurc.SendData "wh " & Furre & " You dont have a weapon to sell." & vbLf
End If
End If
End Sub

Sub dosellarmor(Furre)

Dim price As Integer

members.Recordset.MoveFirst
Do Until txtName.Text = Furre Or members.Recordset.EOF
members.Recordset.MoveNext
Loop
If txtName.Text <> Furre Then
sckFurc.SendData "wh " & Furre & " You are not registered." & vbLf

Else

If txtArmor.Text <> 0 Then

sckFurc.SendData "wh " & Furre & " Armor sold" & vbLf

    Select Case txtArmor.Text
        Case 1
            price = 40
        Case 2
            price = 100
        Case 3
            price = 200
        Case 4
            price = 250
        Case 5
            price = 300
        Case 6
            price = 400
        Case 7
            price = 450
        Case 8
            price = 550
        Case 9
            price = 600
        Case 10
            price = 1000
     End Select

        txtArmor.Text = "0"
        txtCoin = txtCoin + price
        members.Recordset.Update
        
Else
sckFurc.SendData "wh " & Furre & " You dont have armor to sell." & vbLf
End If
End If
End Sub

Sub doBuy(Furre, msg)

'On Error Resume Next

Dim price As Integer

done = False
members.Recordset.MoveFirst
Do Until txtName.Text = Furre Or members.Recordset.EOF
members.Recordset.MoveNext
Loop

If txtName.Text <> Furre Then
sckFurc.SendData "wh " & Furre & " You are not registered." & vbLf

Else

    msg = Right(msg, Len(msg) - 4)
    Select Case msg
        Case "rusty dagger"
            num = 1
            price = 40
            typ = "weapon"
        Case "rusty axe"
            num = 2
            price = 100
            typ = "weapon"
        Case "heavy axe"
            num = 3
            price = 200
            typ = "weapon"
        Case "dagger"
            num = 4
            price = 250
            typ = "weapon"
        Case "axe"
            num = 5
            price = 300
            typ = "weapon"
        Case "rusty sword"
            num = 6
            price = 400
            typ = "weapon"
        Case "fine steel dagger"
            num = 7
            price = 450
            typ = "weapon"
        Case "fine steel axe"
            num = 8
            price = 550
            typ = "weapon"
        Case "sword"
            num = 9
            price = 600
            typ = "weapon"
        Case "fine steel sword"
            num = 10
            price = 1000
            typ = "weapon"
            
        Case "rawhide"
            num = 1
            price = 40
            typ = "armor"
        Case "tathered leather"
            num = 2
            price = 100
            typ = "armor"
        Case "leather"
            num = 3
            price = 200
            typ = "armor"
        Case "studed leather"
            num = 4
            price = 250
            typ = "armor"
        Case "rusty chain mail"
            num = 5
            price = 300
            typ = "armor"
        Case "chain mail"
            num = 6
            price = 400
            typ = "armor"
        Case "banded"
            num = 7
            price = 450
            typ = "armor"
        Case "rusty steel"
            num = 8
            price = 550
            typ = "armor"
        Case "steel"
            num = 9
            price = 600
            typ = "armor"
        Case "fine steel"
            num = 10
            price = 1000
            typ = "armor"
            
            
    End Select
        
If typ = "weapon" Then
If txtWeapon <> 0 Then
sckFurc.SendData "wh " & Furre & " You must first sell your old weapon." & vbLf
done = True
Else
    If txtCoin >= price Then
        sckFurc.SendData "wh " & Furre & " You now have a " & msg & "." & vbLf
        txtWeapon.Text = num
        txtCoin = txtCoin - price
        members.Recordset.Update
        done = True
    Else
        sckFurc.SendData "wh " & Furre & " You dont have enuff coin to buy that." & vbLf
        done = True
    End If
End If
ElseIf typ = "armor" Then
If txtArmor <> 0 Then
sckFurc.SendData "wh " & Furre & " You must first sell your old armor." & vbLf
done = True

Else
    If txtCoin >= price Then
    
        sckFurc.SendData "wh " & Furre & " You now have " & msg & " Armor." & vbLf
        txtArmor.Text = num
        txtCoin = txtCoin - price
        members.Recordset.Update
        done = True
    Else
        sckFurc.SendData "wh " & Furre & " You dont have enuff coin to buy that." & vbLf
        done = True
    End If
End If
End If

If done = False Then
sckFurc.SendData "wh " & Furre & " I don't understand! Whisper me HELP for more info!" & vbLf
End If

End If
End Sub

Sub dojoin(Furre)
On Error Resume Next
members.Recordset.MoveFirst
Do Until txtName.Text = Furre Or members.Recordset.EOF
members.Recordset.MoveNext
Loop

If txtName.Text <> Furre Then
members.Recordset.AddNew
txtName.Text = Furre
txtLvl.Text = 1
txtExp.Text = 0
txtArmor.Text = 0
txtWeapon.Text = 0
txtMana.Text = 0
sckFurc.SendData "wh " & Furre & " You are now registered." & vbLf



Else
sckFurc.SendData "wh " & Furre & " You are allready registered." & vbLf
End If
End Sub


Public Sub doStats(Furre)

members.Recordset.MoveFirst
Do Until txtName.Text = Furre Or members.Recordset.EOF
members.Recordset.MoveNext
Loop

If txtName.Text <> Furre Then
sckFurc.SendData "wh " & Furre & " You are not registered." & vbLf

Else
Select Case txtArmor.Text
    Case 0
        Armor = "Cloth"
    Case 1
        Armor = "Rahide"
    Case 2
        Armor = "Tathered Leather"
    Case 3
        Armor = "Leather"
    Case 4
        Armor = "Studed Leather"
    Case 5
        Armor = "Rusty Chain Mail"
    Case 6
        Armor = "Chain mail"
    Case 7
        Armor = "Banded"
    Case 8
        Armor = "Rusty Steel"
    Case 9
        Armor = "Steel"
    Case 10
        Armor = "Fine Steel"
End Select

Select Case txtWeapon.Text
    Case 0
        Weapon = "Paws"
    Case 1
        Weapon = "Rusty Dagger"
    Case 2
        Weapon = "Rusty Axe"
    Case 3
        Weapon = "Heavy Axe"
    Case 4
        Weapon = "Dagger"
    Case 5
        Weapon = "Axe"
    Case 6
        Weapon = "Rusty Sword"
    Case 7
        Weapon = "Fine Steel Dagger"
    Case 8
        Weapon = "Fine Steel Axe"
    Case 9
        Weapon = "Sword"
    Case 10
        Weapon = "Fine Steel Sword"
End Select

sckFurc.SendData "wh " & Furre & " Youre Stats are,||||||Member#: " & txtNum.Text & ", Level: " & txtLvl.Text & ", Exp: " & txtExp.Text & "%, Armor: " & Armor & ", Weapon: " & Weapon & ", Mana: " & txtMana.Text & ", Coin's: " & txtCoin.Text & "." & vbLf

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

Private Sub whsp_Click()
If whsp.Checked = False Then
whsp.Checked = True
ElseIf stext.Checked = True Then
whsp.Checked = False
End If
End Sub
