VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmBot 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EZbot"
   ClientHeight    =   5280
   ClientLeft      =   150
   ClientTop       =   150
   ClientWidth     =   6930
   Icon            =   "frmBot.frx":0000
   LinkTopic       =   "Form1"
   MouseIcon       =   "frmBot.frx":0442
   ScaleHeight     =   5280
   ScaleWidth      =   6930
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Time On"
      Height          =   615
      Left            =   4080
      TabIndex        =   23
      Top             =   3000
      Width           =   1455
      Begin VB.TextBox txtTime 
         Height          =   285
         Left            =   120
         TabIndex        =   24
         Text            =   "0:0:0"
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdWhis 
      Caption         =   "Whispers"
      Height          =   495
      Left            =   5640
      TabIndex        =   22
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit EZbot"
      Height          =   495
      Left            =   4080
      TabIndex        =   21
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Connected"
      Height          =   615
      Left            =   4080
      TabIndex        =   19
      Top             =   2280
      Width           =   1455
      Begin VB.TextBox txtCont 
         Height          =   285
         Left            =   120
         TabIndex        =   20
         Text            =   "False"
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "System"
      Height          =   2055
      Left            =   4080
      TabIndex        =   17
      Top             =   120
      Width           =   1455
      Begin VB.CheckBox chkServCode 
         Caption         =   "ServerCode"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdWings 
      Caption         =   "Toggle Wings"
      Height          =   495
      Left            =   2760
      TabIndex        =   16
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton cmdUse 
      Caption         =   "Use"
      Height          =   495
      Left            =   720
      TabIndex        =   15
      Top             =   4200
      Width           =   615
   End
   Begin VB.CommandButton cmdGet 
      Caption         =   "Get"
      Height          =   495
      Left            =   120
      TabIndex        =   14
      Top             =   4200
      Width           =   615
   End
   Begin VB.CommandButton cmdTl 
      Caption         =   "<"
      Height          =   495
      Left            =   1440
      TabIndex        =   13
      Top             =   4680
      Width           =   615
   End
   Begin VB.CommandButton cmgTr 
      Caption         =   ">"
      Height          =   495
      Left            =   2040
      TabIndex        =   12
      Top             =   4680
      Width           =   615
   End
   Begin VB.CommandButton cmdGoVinca 
      Caption         =   "Vinca"
      Height          =   495
      Left            =   2760
      TabIndex        =   11
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton cmdGoAlleg 
      Caption         =   "Allegria"
      Height          =   495
      Left            =   2760
      TabIndex        =   10
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdWho 
      Caption         =   "Who"
      Height          =   495
      Left            =   720
      TabIndex        =   9
      Top             =   3720
      Width           =   615
   End
   Begin VB.CommandButton cmdLay 
      Caption         =   "Lay"
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   3720
      Width           =   615
   End
   Begin VB.CommandButton cmdSE 
      Caption         =   "SE"
      Height          =   495
      Left            =   2040
      TabIndex        =   7
      Top             =   4200
      Width           =   615
   End
   Begin VB.CommandButton cmdSW 
      Caption         =   "SW"
      Height          =   495
      Left            =   1440
      TabIndex        =   6
      Top             =   4200
      Width           =   615
   End
   Begin VB.CommandButton cmdNE 
      Caption         =   "NE"
      Height          =   495
      Left            =   2040
      TabIndex        =   5
      Top             =   3720
      Width           =   615
   End
   Begin VB.CommandButton cmdNW 
      Caption         =   "NW"
      Height          =   495
      Left            =   1440
      TabIndex        =   4
      Top             =   3720
      Width           =   615
   End
   Begin VB.CommandButton cmdDisconnect 
      Caption         =   "Disconnect"
      Height          =   495
      Left            =   4080
      TabIndex        =   3
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      Height          =   495
      Left            =   4080
      TabIndex        =   2
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Timer StayOnline 
      Interval        =   60000
      Left            =   240
      Top             =   840
   End
   Begin MSWinsockLib.Winsock sckFurc 
      Left            =   240
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtSend 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   3360
      Width           =   3855
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
Public Day As Integer
Public Hour As Integer
Public Minute As Integer
Public BotName, BotPass, ColorCode, Desca, Desc As String
Public furAddr, furPort As Variant
Public Connected As Boolean
Private Sub cmdGet_Click()
sckFurc.SendData "get" & vbLf
End Sub
Private Sub cmdGoAlleg_Click()
sckFurc.SendData "goalleg" & vbLf
End Sub
Private Sub cmdGoVinca_Click()
sckFurc.SendData "gostart" & vbLf
End Sub
Private Sub cmdLay_Click()
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
Private Sub cmdTl_Click()
sckFurc.SendData "<" & vbLf
End Sub
Private Sub cmdUse_Click()
sckFurc.SendData "use" & vbLf
End Sub
Private Sub cmdWhis_Click()
Load whisper
whisper.Show
End Sub
Private Sub cmdWho_Click()
sckFurc.SendData "who" & vbLf
End Sub
Private Sub cmdWings_Click()
sckFurc.SendData "wings" & vbLf
End Sub
Private Sub cmgTr_Click()
sckFurc.SendData ">" & vbLf
End Sub
Sub Form_Load()
Open "C:/EZbot/setting.ini" For Input As #1
Input #1, BotName
Input #1, BotPass
Input #1, botCod
Input #1, botSex
Input #1, botRace
Input #1, Desca
Input #1, furAddr
Input #1, furPort
Close #1
If botSex = "Female" Then cSex = " "
If botSex = "Male" Then cSex = "!"
If botSex = "Unspecified" Then cSex = Chr(34)
If botRace = "Rodent" Then cRace = " "
If botRace = "Equine" Then cRace = "!"
If botRace = "Feline" Then cRace = Chr(34)
If botRace = "Canine" Then cRace = "#"
If botRace = "Musteline" Then cRace = "$"
If botRace = "Lapine" Then cRace = 5
ColorCode = botCod & cSex & cRace & "!"
Minute = 0
Desc = Desca & " [Uptime: 0 Minute(s)]"
sckFurc.RemoteHost = furAddr
sckFurc.RemotePort = furPort
sckFurc.Connect
Connected = True
End Sub
Private Sub cmdConnect_Click()
If Connected = False Then
sckFurc.RemoteHost = furAddr
sckFurc.RemotePort = furPort
sckFurc.Connect
Connected = True
Else
box = MsgBox("EZbot is eather connected, or trying to connect. Press Disconnect to close the link to Furcadia.", vbOKOnly, "Error")
End If
End Sub
Private Sub cmdDisconnect_Click()
If Connected = True Then
sckFurc.Close
Connected = False
txtCont = "False"
End If
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
If chkServCode.Value = Checked Then txtFromFurc = txtFromFurc & Txt & vbCrLf
If chkServCode.Value = Unchecked Then
If Left(Txt, 1) = "(" Then txtFromFurc = txtFromFurc & Right(Txt, Len(Txt) - 1) & vbCrLf
End If
If Txt = "END" Then
sckFurc.SendData "connect " & BotName & " " & BotPass & vbLf & "color " & ColorCode & vbLf & "desc " & Desc & vbLf
txtCont = "True"
End If
If Txt = "]ccmarbled.pcx" Then sckFurc.SendData "vascodagama" & vbLf

If Left(Txt, 3) = "([ " And Right(Txt, 10) = " to you. ]" Then
    Tmsg = Split(Txt, " whispers, " & Chr(34), 2)
    Furre = Right(Tmsg(0), Len(Tmsg(0)) - 3)
    Msg = Left(Tmsg(1), Len(Tmsg(1)) - 11)
    DoWhisper Furre, Msg
End If

If Left(Txt, 10) = "((You see " Then Furre = Mid(Txt, 11, Len(Txt) - 12)
End Sub
Sub DoWhisper(Furre, Msg)
msent = 0
Msg = LCase(Msg)
Open "C:/EZbot/whisper.lst" For Input As #1
Do Until mssg = Msg Or EOF(1)
Input #1, mssg, back
If Msg Like mssg Then
    back = Replace(back, "%1", Furre)
    sckFurc.SendData "wh " & Furre & " " & back & vbLf
    msent = 1
End If
Loop
Close #1
If msent = 0 Then sckFurc.SendData "wh " & Furre & " I'm sorry. I don't understand that." & vbLf
End Sub
Private Sub cmdExit_Click()
box = MsgBox("Are you shure you want to close EZbot?", vbOKCancel, "Exit Program")
If box = vbOK Then End
End Sub
Private Sub StayOnline_Timer()
Minute = Minute + 1
If Minute >= 60 Then
    Hour = Hour + 1
    Minute = 0
End If
If Hour >= 24 Then
    Day = Day + 1
    Hour = 0
End If
txtTime = Day & ":" & Hour & ":" & Minute
sckFurc.SendData "desc " & Desca & " [Uptime: "
If Day >= 1 Then sckFurc.SendData Day & " Day(s) "
If Hour >= 1 Then sckFurc.SendData Hour & " Hour(s) "
sckFurc.SendData Minute & " Minute(s)]" & vbLf
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
