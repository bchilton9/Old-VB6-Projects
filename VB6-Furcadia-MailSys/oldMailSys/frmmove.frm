VERSION 5.00
Begin VB.Form frmmove 
   BackColor       =   &H8000000A&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   1200
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2745
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1200
   ScaleWidth      =   2745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdNW 
      BackColor       =   &H8000000A&
      Caption         =   "NW"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton cmdNE 
      BackColor       =   &H8000000A&
      Caption         =   "NE"
      Height          =   495
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton cmdSW 
      BackColor       =   &H8000000A&
      Caption         =   "SW"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton cmdSE 
      BackColor       =   &H8000000A&
      Caption         =   "SE"
      Height          =   495
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton cmdLie 
      BackColor       =   &H8000000A&
      Caption         =   "&Lie"
      Height          =   255
      Index           =   0
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton cmdWho 
      BackColor       =   &H8000000A&
      Caption         =   "&Who"
      Height          =   255
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton cmdGet 
      BackColor       =   &H8000000A&
      Caption         =   "&Get"
      Height          =   255
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   840
      Width           =   615
   End
   Begin VB.CommandButton cmduse 
      BackColor       =   &H8000000A&
      Caption         =   "&Use"
      Height          =   255
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   840
      Width           =   615
   End
   Begin VB.CommandButton cmdturnr 
      BackColor       =   &H8000000A&
      Caption         =   "< Turn"
      Height          =   375
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton cmdturnl 
      BackColor       =   &H8000000A&
      Caption         =   "Turn >"
      Height          =   375
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmmove"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private ontop As New clsOnTop
Private Sub cmdGet_Click()
If frmBot.Connected = True Then frmBot.sckFurc.SendData "get" & vbLf
End Sub
Private Sub cmdLie_Click(Index As Integer)
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
Private Sub cmdWho_Click()
If frmBot.Connected = True Then frmBot.sckFurc.SendData "who" & vbLf
End Sub
Private Sub Form_Load()
  Set ontop = New clsOnTop
  ontop.MakeTopMost hWnd
End Sub
