VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmBot 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Example Bot"
   ClientHeight    =   4200
   ClientLeft      =   150
   ClientTop       =   150
   ClientWidth     =   5385
   Icon            =   "frmBot.frx":0000
   LinkTopic       =   "Form1"
   MouseIcon       =   "frmBot.frx":0442
   ScaleHeight     =   4200
   ScaleWidth      =   5385
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtcon 
      Height          =   285
      Left            =   4200
      Locked          =   -1  'True
      TabIndex        =   22
      Text            =   "False"
      Top             =   2880
      Width           =   855
   End
   Begin VB.TextBox timon 
      Height          =   285
      Left            =   4200
      Locked          =   -1  'True
      TabIndex        =   21
      Text            =   "0:0:0"
      Top             =   2520
      Width           =   855
   End
   Begin VB.CommandButton cmdGet 
      Caption         =   "&Get"
      Height          =   495
      Left            =   1440
      TabIndex        =   20
      Top             =   2880
      Width           =   615
   End
   Begin VB.CommandButton cmduse 
      Caption         =   "&Use"
      Height          =   495
      Left            =   2040
      TabIndex        =   19
      Top             =   2880
      Width           =   615
   End
   Begin VB.Frame Frame2 
      Caption         =   "System"
      Height          =   855
      Left            =   3360
      TabIndex        =   16
      Top             =   1200
      Width           =   1215
      Begin VB.CheckBox chkServtxt 
         Caption         =   "SText"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CheckBox chkServCode 
         Caption         =   "SCode"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Fighting"
      Height          =   975
      Left            =   3360
      TabIndex        =   12
      Top             =   120
      Width           =   1215
      Begin VB.CheckBox chkFight 
         Caption         =   "On/Off"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CommandButton cmdrest 
         Caption         =   "&Reset"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   975
      End
      Begin VB.CheckBox chkDream 
         Caption         =   "Dream"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdGoVinca 
      Caption         =   "Vinca"
      Height          =   255
      Left            =   2760
      TabIndex        =   11
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdGoAlleg 
      Caption         =   "Allegria"
      Height          =   255
      Left            =   2760
      TabIndex        =   10
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmdWho 
      Caption         =   "Who"
      Height          =   495
      Left            =   2040
      TabIndex        =   9
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton cmdLay 
      Caption         =   "Lay"
      Height          =   495
      Left            =   1440
      TabIndex        =   8
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton cmdSE 
      Caption         =   "SE"
      Height          =   495
      Left            =   720
      TabIndex        =   7
      Top             =   2880
      Width           =   615
   End
   Begin VB.CommandButton cmdSW 
      Caption         =   "SW"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   2880
      Width           =   615
   End
   Begin VB.CommandButton cmdNE 
      Caption         =   "NE"
      Height          =   495
      Left            =   720
      TabIndex        =   5
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton cmdNW 
      Caption         =   "NW"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton cmdDisconnect 
      Caption         =   "Disconnect"
      Height          =   255
      Left            =   2760
      TabIndex        =   3
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      Height          =   255
      Left            =   2760
      TabIndex        =   2
      Top             =   2400
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
      Left            =   0
      TabIndex        =   1
      Top             =   2040
      Width           =   3255
   End
   Begin VB.TextBox txtFromFurc 
      Height          =   1935
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   3255
   End
End
Attribute VB_Name = "frmBot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Minute, Hour, Day, furhost, furport As Integer
Public ColorCode, Desc, BotName, BotPass, descp, fight As String
Public Connected As Boolean

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
Private Sub cmdWho_Click()
sckFurc.SendData "who" & vbLf
End Sub

Sub Form_Load()
setting
Desc = descp & " [Uptime: 0 Minute(s)]"
End Sub

Private Sub cmdConnect_Click()
If Connected = False Then
sckFurc.RemoteHost = furhost
sckFurc.RemotePort = furport
sckFurc.Connect
End If
End Sub

Private Sub cmdDisconnect_Click()
If Connected = True Then
sckFurc.Close
Connected = False
txtcon.Text = "False"
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
Connected = True
txtcon.Text = "True"
End If
If Txt = "]ccmarbled.pcx" Then sckFurc.SendData "vascodagama" & vbLf

If Left(Txt, 1) = "<" Then
pace = Right(Txt, Len(Txt) - 11)
pace = Left(pace, Len(pace) - 2)
If pace = " 7 Y" Then
sckFurc.SendData "l  7 Y" & vbLf
fight = 1
ElseIf pace = " 9 ]" Then
sckFurc.SendData "l  9 ]" & vbLf
fight = 2
End If

End If
If Left(Txt, 3) = "([ " And Right(Txt, 10) = " to you. ]" Then
    Tmsg = Split(Txt, " whispers, " & Chr(34), 2)
    Furre = Right(Tmsg(0), Len(Tmsg(0)) - 3)
    Msg = Left(Tmsg(1), Len(Tmsg(1)) - 11)
    DoWhisper Furre, Msg
End If

If Left(Txt, 10) = "((You see " Then
Furre = Mid(Txt, 11, Len(Txt) - 12)
sckFurc.SendData Chr(34) & Furre & " has entered the arena." & vbLf
End If
End Sub

Sub DoWhisper(Furre, Msg)

If Msg Like "*" Then sckFurc.SendData "wh " & Furre & " I have nothing to say to you." & vbLf
End Sub

Private Sub cmdExit_Click()
End
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
    timon.Text = Day & ":" & Hour & ":" & Minute
    sckFurc.SendData "desc " & cDesc & " [Uptime: "
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
