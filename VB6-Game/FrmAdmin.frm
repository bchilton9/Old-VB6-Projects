VERSION 5.00
Begin VB.Form FrmAdmin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Admin Panel"
   ClientHeight    =   2220
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7665
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   7665
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   360
      Top             =   2400
   End
   Begin VB.ListBox lstOnline 
      Height          =   2010
      Left            =   5880
      TabIndex        =   15
      Top             =   120
      Width           =   1695
   End
   Begin VB.Frame Frame3 
      Caption         =   "Offline Message"
      Height          =   1335
      Left            =   0
      TabIndex        =   8
      Top             =   840
      Width           =   5775
      Begin VB.CommandButton cmdSendMsg 
         Caption         =   "Send Message"
         Height          =   255
         Left            =   2160
         TabIndex        =   14
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox txtOffMsg 
         Height          =   285
         Left            =   1440
         TabIndex        =   13
         Top             =   600
         Width           =   4095
      End
      Begin VB.TextBox txtOffMsgName 
         Height          =   285
         Left            =   1440
         TabIndex        =   10
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Message:"
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Name:"
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Kick"
      Height          =   855
      Left            =   2880
      TabIndex        =   5
      Top             =   0
      Width           =   2895
      Begin VB.CommandButton cmdKick 
         Caption         =   "Kick"
         Default         =   -1  'True
         Enabled         =   0   'False
         Height          =   255
         Left            =   720
         TabIndex        =   7
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtKick 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   6
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Name:"
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "System"
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2895
      Begin VB.CommandButton cmdDebugger 
         Caption         =   "Debugger"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdLoad 
         Caption         =   "Load Skin"
         Height          =   255
         Left            =   1440
         TabIndex        =   3
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton cmdReload 
         Caption         =   "Reload Server"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Edit Members"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1440
         TabIndex        =   1
         Top             =   240
         Width           =   1335
      End
   End
End
Attribute VB_Name = "FrmAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private ontop As New clsOnTop

Private Sub cmdDebugger_Click()
Load frmDebugger
frmDebugger.Show
End Sub

Private Sub cmdLoad_Click()
Dim skinfile As String
skinfile = InputBox("Load skin name?", "Load Skin")
If skinfile = "" Then
Exit Sub
Else
loadskin skinfile
End If
End Sub

Private Sub cmdReload_Click()
frmTCP.wskServer.SendData "reload" & pEnd
End Sub


Private Sub cmdSendMsg_Click()
frmTCP.wskServer.SendData "adminmsg" & pChar & txtOffMsgName.Text & pChar & txtOffMsg & pEnd
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Unload Me

End Sub

Private Sub Form_Load()

  Set ontop = New clsOnTop
  ontop.MakeTopMost hWnd
  
  
    For i = 1 To MaxPlayers
    If player(i).Name <> "" Then
    lstOnline.AddItem player(i).Name
    End If
    Next i

  
End Sub

Private Sub Timer1_Timer()
lstOnline.clear
    For i = 1 To MaxPlayers
    If player(i).Name <> "" Then
    lstOnline.AddItem player(i).Name
    End If
    Next i
End Sub
