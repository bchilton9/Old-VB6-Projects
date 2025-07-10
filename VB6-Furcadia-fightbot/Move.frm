VERSION 5.00
Begin VB.Form frmMove 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   1200
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1200
   ScaleWidth      =   4110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdNW 
      Caption         =   "NW"
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton cmdNE 
      Caption         =   "NE"
      Height          =   495
      Left            =   720
      TabIndex        =   8
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton cmdSW 
      Caption         =   "SW"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton cmdSE 
      Caption         =   "SE"
      Height          =   495
      Left            =   720
      TabIndex        =   6
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton cmdLay 
      Caption         =   "Lay"
      Height          =   495
      Index           =   0
      Left            =   1440
      TabIndex        =   5
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton cmdWho 
      Caption         =   "Who"
      Height          =   495
      Left            =   2040
      TabIndex        =   4
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton cmdGoAlleg 
      Caption         =   "Allegria"
      Height          =   495
      Left            =   2760
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdGoVinca 
      Caption         =   "Vinca"
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdGet 
      Caption         =   "Get"
      Height          =   495
      Index           =   1
      Left            =   1440
      TabIndex        =   1
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton cmduse 
      Caption         =   "Use"
      Height          =   495
      Index           =   1
      Left            =   2040
      TabIndex        =   0
      Top             =   600
      Width           =   615
   End
End
Attribute VB_Name = "frmMove"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdGet_Click(Index As Integer)
frmBot.sckFurc.SendData "get" & vbLf
End Sub

Private Sub cmdGoAlleg_Click()
frmBot.sckFurc.SendData "goalleg" & vbLf
End Sub

Private Sub cmdGoVinca_Click()
frmBot.sckFurc.SendData "gostart" & vbLf
End Sub

Private Sub cmdlie_Click()
frmBot.sckFurc.SendData "lie" & vbLf
End Sub

Private Sub cmdLay_Click(Index As Integer)
frmBot.sckFurc.SendData "lay" & vbLf
End Sub

Private Sub cmdNE_Click()
frmBot.sckFurc.SendData "m 9" & vbLf
End Sub

Private Sub cmdNW_Click()
frmBot.sckFurc.SendData "m 7" & vbLf
End Sub

Private Sub cmdSE_Click()
frmBot.sckFurc.SendData "m 3" & vbLf
End Sub

Private Sub cmdSW_Click()
frmBot.sckFurc.SendData "m 1" & vbLf
End Sub

Private Sub cmduse_Click(Index As Integer)
frmBot.sckFurc.SendData "use" & vbLf
End Sub

Private Sub cmdWho_Click()
frmBot.sckFurc.SendData "who" & vbLf
End Sub

