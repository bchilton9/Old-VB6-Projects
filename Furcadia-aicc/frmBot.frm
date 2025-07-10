VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmBot 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AICC"
   ClientHeight    =   3720
   ClientLeft      =   150
   ClientTop       =   150
   ClientWidth     =   3585
   Icon            =   "frmBot.frx":0000
   LinkTopic       =   "Form1"
   MouseIcon       =   "frmBot.frx":0442
   ScaleHeight     =   3720
   ScaleWidth      =   3585
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   14
      Top             =   0
      Width           =   3375
      Begin VB.TextBox timon 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000007&
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   16
         Text            =   "0:0:0"
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtcnt 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000007&
         Height          =   285
         Left            =   600
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   "False"
         Top             =   480
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
         Left            =   1680
         TabIndex        =   18
         Top             =   240
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
         Left            =   600
         TabIndex        =   17
         Top             =   240
         Width           =   975
      End
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
      Left            =   3720
      TabIndex        =   13
      Top             =   360
      Width           =   975
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
      Left            =   3720
      TabIndex        =   12
      Top             =   600
      Value           =   2  'Grayed
      Width           =   975
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
      Left            =   3720
      TabIndex        =   11
      Top             =   120
      Value           =   1  'Checked
      Width           =   855
   End
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
      Left            =   3720
      TabIndex        =   10
      Top             =   840
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.CommandButton cmdButton 
      BackColor       =   &H8000000A&
      Caption         =   "&Connect"
      Height          =   375
      Index           =   0
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton cmdButton 
      BackColor       =   &H8000000A&
      Caption         =   "&Disconnect"
      Height          =   375
      Index           =   1
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton cmdButton 
      BackColor       =   &H8000000A&
      Caption         =   "&Allegria"
      Height          =   255
      Index           =   3
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton cmdButton 
      BackColor       =   &H8000000A&
      Caption         =   "&Vinca"
      Height          =   255
      Index           =   4
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton cmdButton 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      Caption         =   "Clear"
      Height          =   255
      Index           =   12
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3360
      Width           =   735
   End
   Begin VB.CommandButton cmdButton 
      BackColor       =   &H8000000A&
      Caption         =   "Movement"
      Height          =   255
      Index           =   2
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton cmdButton 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      Caption         =   "Send"
      Height          =   255
      Index           =   11
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3360
      Width           =   735
   End
   Begin VB.CommandButton cmdPanel 
      BackColor       =   &H8000000A&
      Caption         =   "More >"
      Height          =   255
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3360
      Width           =   735
   End
   Begin VB.Timer StayOnline 
      Interval        =   60000
      Left            =   720
      Top             =   3840
   End
   Begin MSWinsockLib.Winsock sckFurc 
      Left            =   240
      Top             =   3840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtSend 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   3000
      Width           =   3375
   End
   Begin VB.TextBox txtFromFurc 
      Height          =   1920
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   960
      Width           =   3375
   End
End
Attribute VB_Name = "frmBot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Minute, Hour, Day, frcHost, frcPort As Integer
Public Connected As Boolean
Public BotName, BotPass, descrip, ColorCode, Desc As String

Sub Form_Load()
dosettings
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
    sckFurc.SendData "connect " & BotName & " " & BotPass & vbLf & "color " & ColorCode & vbLf & "desc " & Desc & vbLf
    Connected = "True"
    txtcnt = "True"
    'sckFurc.SendData "goalleg" & vbLf
End If
If Txt = "]ccmarbled.pcx" Then
    sckFurc.SendData "vascodagama" & vbLf
End If
incomeingtxt Txt
End Sub

Private Sub StayOnline_Timer()
If Connected = True Then
    timeon
End If
If Connected = False And chkcnt = 1 Then
    On Error GoTo tcerr
    sckFurc.Close
    Connected = False
    txtcnt = "False"
    sckFurc.RemoteHost = frcHost
    sckFurc.RemotePort = frcPort
    sckFurc.Connect
End If
Exit Sub
tcerr:
    Close #1, #3
    Resume stoptrying
stoptrying:
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
Private Sub cmdPanel_Click()
If frmBot.Width = 3675 Then
    cmdPanel.Caption = "Less <"
    frmBot.Width = 5010
Else
    cmdPanel.Caption = "More >"
    frmBot.Width = 3675
End If
End Sub


Private Sub cmdButton_Click(Index As Integer)
    Select Case Index
        Case 0 'Connect
            If Connected = False Then
                On Error GoTo conerr
                sckFurc.RemoteHost = frcHost
                sckFurc.RemotePort = frcPort
                sckFurc.Connect
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
