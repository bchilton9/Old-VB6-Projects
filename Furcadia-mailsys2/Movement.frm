VERSION 5.00
Begin VB.Form frmMovement 
   BackColor       =   &H00404040&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   3135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Caption         =   "Movement Controls"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2895
      Begin VB.Image cmdNW 
         Height          =   555
         Left            =   960
         Picture         =   "Movement.frx":0000
         Top             =   360
         Width           =   540
      End
      Begin VB.Image cmdNE 
         Height          =   555
         Left            =   1500
         Picture         =   "Movement.frx":0FDE
         Top             =   360
         Width           =   525
      End
      Begin VB.Image cmdSW 
         Height          =   510
         Left            =   960
         Picture         =   "Movement.frx":1FBC
         Top             =   915
         Width           =   540
      End
      Begin VB.Image cmdSE 
         Height          =   510
         Left            =   1500
         Picture         =   "Movement.frx":2E56
         Top             =   915
         Width           =   525
      End
      Begin VB.Image cmdturnl 
         Height          =   210
         Left            =   1560
         Picture         =   "Movement.frx":3CF0
         Top             =   1560
         Width           =   180
      End
      Begin VB.Image cmdturnr 
         Height          =   210
         Left            =   1200
         Picture         =   "Movement.frx":3F2A
         Top             =   1560
         Width           =   180
      End
      Begin VB.Image cmdLie 
         Height          =   450
         Left            =   120
         Picture         =   "Movement.frx":4164
         Top             =   360
         Width           =   735
      End
      Begin VB.Image cmdGet 
         Height          =   375
         Left            =   2160
         Picture         =   "Movement.frx":52FE
         Top             =   360
         Width           =   480
      End
      Begin VB.Image cmduse 
         Height          =   495
         Left            =   2160
         Picture         =   "Movement.frx":5CA0
         Top             =   840
         Width           =   270
      End
      Begin VB.Image cmdwing 
         Height          =   585
         Left            =   360
         Picture         =   "Movement.frx":641A
         Top             =   840
         Width           =   420
      End
      Begin VB.Image cmdwho 
         Height          =   210
         Left            =   240
         Picture         =   "Movement.frx":7128
         Top             =   1560
         Width           =   570
      End
   End
End
Attribute VB_Name = "frmMovement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_SYSCOMMAND = &H112
Private ontop As New clsOnTop
Private Sub cmdGet_Click()
If frmBot.Connected = True Then frmBot.sckFurc.SendData "get" & vbLf
End Sub

Private Sub cmdLie_Click()
If frmBot.Connected = True Then frmBot.sckFurc.SendData "lie" & vbLf
End Sub

Private Sub cmdNE_Click()
If frmBot.Connected = True Then frmBot.sckFurc.SendData "m 9" & vbLf
End Sub

Private Sub cmdNW_Click()
If frmBot.Connected = True Then frmBot.sckFurc.SendData "m 7" & vbLf
End Sub

Private Sub cmdSE_Click()
If frmBot.Connected = True Then frmBot.sckFurc.SendData "m 3" & vbLf
End Sub

Private Sub cmdSW_Click()
If frmBot.Connected = True Then frmBot.sckFurc.SendData "m 1" & vbLf
End Sub

Private Sub cmdturnl_Click()
If frmBot.Connected = True Then frmBot.sckFurc.SendData ">" & vbLf
End Sub

Private Sub cmdturnr_Click()
If frmBot.Connected = True Then frmBot.sckFurc.SendData "<" & vbLf
End Sub

Private Sub cmduse_Click()
If frmBot.Connected = True Then frmBot.sckFurc.SendData "use" & vbLf
End Sub

Private Sub cmdwho_Click()
If frmBot.Connected = True Then frmBot.sckFurc.SendData "who" & vbLf
End Sub

Private Sub cmdwing_Click()
If frmBot.Connected = True Then frmBot.sckFurc.SendData "wings" & vbLf
End Sub


Private Sub Form_Load()
  Set ontop = New clsOnTop
  ontop.MakeTopMost hWnd
End Sub


Private Sub Form_Unload(Cancel As Integer)
frmBot.Movement.Checked = False
End Sub
