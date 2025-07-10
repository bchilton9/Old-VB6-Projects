VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmBot 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Valka"
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
   Begin VB.Frame Frame2 
      Caption         =   "System"
      Height          =   855
      Left            =   4080
      TabIndex        =   20
      Top             =   1320
      Width           =   1215
      Begin VB.CheckBox chkServCode 
         Caption         =   "SCode"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   480
         Width           =   975
      End
      Begin VB.CheckBox chkServtxt 
         Caption         =   "SText"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Value           =   1  'Checked
         Width           =   975
      End
   End
   Begin VB.Timer tfight 
      Interval        =   2000
      Left            =   120
      Top             =   1080
   End
   Begin VB.CommandButton cmdturnl 
      Caption         =   "Turn >"
      Height          =   495
      Left            =   2040
      TabIndex        =   18
      Top             =   4200
      Width           =   615
   End
   Begin VB.CommandButton cmdturnr 
      Caption         =   "< Turn"
      Height          =   495
      Left            =   1440
      TabIndex        =   17
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
   Begin VB.Frame Frame1 
      Caption         =   "Fighting"
      Height          =   975
      Left            =   4080
      TabIndex        =   12
      Top             =   120
      Width           =   1215
      Begin VB.CheckBox chkDream 
         Caption         =   "Dream"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1080
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdrest 
         Caption         =   "&Reset"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   975
      End
      Begin VB.CheckBox chkFight 
         Caption         =   "On/Off"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Value           =   1  'Checked
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
Dim fighter
Public Minute As Integer
Public onet
Public twot
Public Desc As String
Public Connected As Boolean
'Bot Settings
Const BotName = "Valka"
Const BotPass = "0519aa"
Const descrip = "This mighty Fighter is a brave adventurer and good friend. He defends his companions against monsters and outher enemies with fierce devotion."
Const ColorCode = "  H2+88888!#!!#!"

Private Sub chkFight_Click()
If chkFight = 0 Then
sckFurc.SendData "m 9" & vbLf
Open "C:\Jovati\fnum.txt" For Output As #1
Write #1, 0
Close #1
Open "C:\Jovati\fighter1.txt" For Output As #1
Write #1, "na", 0, 0, 0, 0, 0, "na"
Close #1
Open "C:\Jovati\fighter2.txt" For Output As #1
Write #1, "na", 0, 0, 0, 0, 0, "na"
Close #1
End If
If chkFight = 1 Then
sckFurc.SendData "m 1" & vbLf
End If
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
sckFurc.SendData Chr(34) & "emitloud Fight system has been reset." & vbLf & "m 1" & vbLf
Open "C:\Jovati\fnum.txt" For Output As #1
Write #1, 0
Close #1
Open "C:\Jovati\fighter1.txt" For Output As #1
Write #1, "na", 0, 0, 0, 0, 0, "na"
Close #1
Open "C:\Jovati\fighter2.txt" For Output As #1
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
End If
End Sub
Private Sub cmdDisconnect_Click()
If Connected = True Then
sckFurc.Close
Connected = False
End If
End Sub


Private Sub sckFurc_DataArrival(ByVal bytesTotal As Long)
Dim s As String
sckFurc.GetData s
X = Split(s, vbLf)
For r = 0 To UBound(X) - 1
RealText X(r)
Next
End Sub
Sub RealText(Txt)
If chkServtxt.Value = Checked Or chkServtxt.Enabled = False Then
If chkServCode.Value = Checked Then txtFromFurc = txtFromFurc & Txt & vbCrLf
If chkServCode.Value = Unchecked Then
If Left(Txt, 1) = "(" Then txtFromFurc = txtFromFurc & Right(Txt, Len(Txt) - 1) & vbCrLf
End If
End If
If Txt = "END" Then sckFurc.SendData "connect " & BotName & " " & BotPass & vbLf & "color " & ColorCode & vbLf & "desc " & Desc & vbLf
If Txt = "]ccmarbled.pcx" Then
sckFurc.SendData "vascodagama" & vbLf
sckFurc.SendData "get" & vbLf & "get" & vbLf & "use" & vbLf & "m 1" & vbLf
lastwalk = "none"
End If
If Left(Txt, 3) = "([ " And Right(Txt, 10) = " to you. ]" Then
    tmsg = Split(Txt, " whispers, " & Chr(34), 2)
    Furre = Right(tmsg(0), Len(tmsg(0)) - 3)
    NMsg = Left(tmsg(1), Len(tmsg(1)) - 11)
    Msg = LCase(NMsg)
    DoWhisper Furre, Msg
End If 'chkWhisp

'watch for fight data
If chkFight.Value = Checked Then
Open "C:\Jovati\fnum.txt" For Input As #1
Input #1, fnum
Close #1
        'watch for data
            If Left(Txt, 1) = "<" Then
            pace = Right(Txt, Len(Txt) - 11)
            pace = Left(pace, Len(pace) - 2)
            If pace = " 7 Y" Then
            sckFurc.SendData "l  7 Y" & vbLf
            fighter = 1
            ElseIf pace = " 9 ]" Then
            sckFurc.SendData "l  9 ]" & vbLf
            fighter = 2
            End If
            End If
            'Gets the furres name when the bot looks at them.
            If Left(Txt, 10) = "((You see " Then
                Furre = Mid(Txt, 11, Len(Txt) - 12)
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
    
    
                If fighter = 1 Then
                    Open "C:\Jovati\fighter1.txt" For Output As #3
                    Write #3, Furre, lvl, hp, man, weap, armo, Class
                    Close #3
                    Open "C:\Jovati\fnum.txt" For Input As #2
                    Input #2, fnum
                    Close #2
                    fnum = fnum + 1
                    Open "C:\Jovati\fnum.txt" For Output As #2
                    Write #2, fnum
                    Close #2
                End If
    
                If fighter = 2 Then
                    Open "C:\Jovati\fighter2.txt" For Output As #3
                    Write #3, Furre, lvl, hp, man, weap, armo, Class
                    Close #3
                    Open "C:\Jovati\fnum.txt" For Input As #2
                    Input #2, fnum
                    Close #2
                    fnum = fnum + 1
                    Open "C:\Jovati\fnum.txt" For Output As #2
                    Write #2, fnum
                    Close #2
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
            
            Open "C:\Jovati\fnum.txt" For Input As #1
            Input #1, fnum
            Close #1
            
            If fnum = 2 Then
                sckFurc.SendData Chr(34) & "emit let the fight begine" & vbLf
                turn = 1
                dofdream turn
            End If
            Else
                If fighter = "1" Then sckFurc.SendData "m 1" & vbLf
                If fighter = "2" Then sckFurc.SendData "m 1" & vbLf
                sckFurc.SendData Chr(34) & "emit " & Furre & " is not a registered member." & vbLf
        End If ' fur = fur
End If ' you see
End If ' CHK FIGHT
End Sub

Sub dofdream(turn)
anum = Int((5 * Rnd) + 1)
If anum = 1 Then atk = "Punched"
If anum = 2 Then atk = "Kicked"
If anum = 3 Then atk = "Bashed"
If anum = 4 Then atk = "Stabed"
If anum = 5 Then atk = "Slashed"
            Open "c:\Jovati\turn.txt" For Input As #1
            Input #1, turn
            Close #1

If turn = 1 Then
    Open "C:\Jovati\fighter1.txt" For Input As #1
    Input #1, fnam, lvl, hp, man, weap, armo, clas
    Close #1
    Open "C:\Jovati\fighter2.txt" For Input As #1
    Input #1, tfnam, tlvl, thp, tman, tweap, tarmo, tClas
    Close #1
    Open "c:\Jovati\turn.txt" For Output As #1
    Write #1, 2
    Close #1
    rnum = lvl * 2
    num = Int((rnum * Rnd) + 1)
    hit = num + weap - tarmo
    If hit < 1 Then hit = 1
    miss = Int((rnum * Rnd) + 1)
    
    If miss = hit Then
        sckFurc.SendData Chr(34) & "emit " & fnam & "'s Attack Missed." & vbLf
    Else
       sckFurc.SendData Chr(34) & "emit " & fnam & " " & atk & " " & tfnam & " For " & hit & " Points of damage." & vbLf
       nhp = thp - hit
       Open "C:\Jovati\fighter2.txt" For Output As #1
       Write #1, tfnam, tlvl, nhp, tman, tweap, tarmo, tClass
       Close #1
       If nhp < 1 Then
       dowin fnam, tfnam
       sckFurc.SendData "m 1" & vbLf
       Else
        Open "c:\Jovati\turn.txt" For Output As #1
        Write #1, 2
        Close #1
       End If
    End If
End If


If turn = 2 Then
    Open "C:\Jovati\fighter2.txt" For Input As #1
    Input #1, fnam, lvl, hp, man, weap, armo, clas
    Close #1
    Open "C:\Jovati\fighter1.txt" For Input As #1
    Input #1, tfnam, tlvl, thp, tman, tweap, tarmo, tClas
    Close #1
    Open "c:\Jovati\turn.txt" For Output As #1
    Write #1, 2
    Close #1
    rnum = lvl * 2
    num = Int((rnum * Rnd) + 1)
    hit = num + weap - tarmo
    If hit < 1 Then hit = 1
    miss = Int((rnum * Rnd) + 1)
    
    If miss = hit Then
        sckFurc.SendData Chr(34) & "emit " & fnam & "'s Attack Missed." & vbLf
    Else
       sckFurc.SendData Chr(34) & "emit " & fnam & " " & atk & " " & tfnam & " For " & hit & " Points of damage." & vbLf
       nhp = thp - hit
       Open "C:\Jovati\fighter1.txt" For Output As #1
       Write #1, tfnam, tlvl, nhp, tman, tweap, tarmo, tClass
       Close #1
       If nhp < 1 Then
       dowin fnam, tfnam
       sckFurc.SendData "m 1" & vbLf
       Else
       Open "c:\Jovati\turn.txt" For Output As #1
        Write #1, 1
        Close #1
       End If
    End If
End If

End Sub
Private Sub tfight_Timer()
            Open "C:\Jovati\fnum.txt" For Input As #1
            Input #1, fnum
            Close #1
            Open "c:\Jovati\turn.txt" For Input As #1
            Input #1, turn
            Close #1
        If fnum = 2 Then
        dofdream turn
        End If
End Sub
Sub dowin(win, lose)
Open "C:\Jovati\members.txt" For Input As #1
Input #1, fName, mnum
Do Until (fName = win) Or (EOF(1))
Input #1, fName, mnum
Loop
Close #1
    Open "C:\Jovati\fighter1.txt" For Input As #1
    Input #1, mfnam, mlvl, mhp, mman, mweap, marmo, mclas
    Close #1
    If mfnam = win Then sckFurc.SendData "m 7" & vbLf
    If mfnam = lose Then sckFurc.SendData "m 3" & vbLf
If fName = win Then
Open "C:\Jovati\memfiles\" & mnum & ".txt" For Input As #1
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
    sckFurc.SendData Chr(34) & "emitloud " & win & " has ganed a lvl." & vbLf
End If
Open "C:\Jovati\memfiles\" & mnum & ".txt" For Output As #1
Write #1, fName, mnum, lvl, Class, ngold, nxp, weap, armo, sone, stwo, sthe, sfor, sfiv
Close #1
End If

Open "C:\Jovati\members.txt" For Input As #1
Input #1, fName, mnum
Do Until (fName = lose) Or (EOF(1))
Input #1, fName, mnum
Loop
Close #1
If fName = lose Then
Open "C:\Jovati\memfiles\" & mnum & ".txt" For Input As #1
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
    sckFurc.SendData Chr(34) & "emitloud " & lose & " has ganed a lvl." & vbLf
End If
Open "C:\Jovati\memfiles\" & mnum & ".txt" For Output As #1
Write #1, fName, mnum, lvl, Class, ngold, nxp, weap, armo, sone, stwo, sthe, sfor, sfiv
Close #1
End If

sckFurc.SendData Chr(34) & "emitloud " & win & " has defeted " & lose & " in a duel to the death." & vbLf
Open "C:\Jovati\fnum.txt" For Output As #1
Write #1, 0
Close #1
Open "C:\Jovati\fighter1.txt" For Output As #1
Write #1, "na", 0, 0, 0, 0, 0, "na"
Close #1
Open "C:\Jovati\fighter2.txt" For Output As #1
Write #1, "na", 0, 0, 0, 0, 0, "na"
Close #1
End Sub


Private Sub StayOnline_Timer()
Minute = Minute + 1
sckFurc.SendData "desc " & descrip & " [Uptime: " & Minute & " Minute(s)]" & vbLf
End Sub


Private Sub txtFromFurc_Change()
txtFromFurc.SelStart = Len(txtFromFurc)
If Len(txtFromFurc) > 10000 Then txtFromFurc = Right(txtFromFurc, 9000)
End Sub

Private Sub txtSend_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    sckFurc.SendData txtSend & vbLf
    txtSend = ""
    KeyAscii = 0
End If
End Sub
