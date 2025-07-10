VERSION 5.00
Begin VB.Form frmmove 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   1290
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   1935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1290
   ScaleWidth      =   1935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   1695
      Begin VB.CommandButton cmdAction 
         BackColor       =   &H00FFC0C0&
         Height          =   375
         Index           =   3
         Left            =   600
         Picture         =   "move.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton cmdAction 
         BackColor       =   &H00FFC0C0&
         Height          =   375
         Index           =   2
         Left            =   120
         Picture         =   "move.frx":0382
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton cmdAction 
         BackColor       =   &H00FFC0C0&
         Height          =   375
         Index           =   1
         Left            =   600
         Picture         =   "move.frx":0705
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdAction 
         BackColor       =   &H00FFC0C0&
         Height          =   375
         Index           =   0
         Left            =   120
         Picture         =   "move.frx":0A83
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdAction 
         BackColor       =   &H00FFC0C0&
         Height          =   375
         Index           =   4
         Left            =   1200
         Picture         =   "move.frx":0E00
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdAction 
         BackColor       =   &H00FFC0C0&
         Height          =   375
         Index           =   5
         Left            =   1200
         Picture         =   "move.frx":1167
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   720
         Width           =   375
      End
   End
End
Attribute VB_Name = "frmmove"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private ontop As New clsOnTop
Private Sub Form_Load()
  Set ontop = New clsOnTop
  ontop.MakeTopMost hWnd
End Sub
Private Sub cmdAction_Click(Index As Integer)
    Dim action          As String
    Select Case Index
        Case 0 'West
            action = "m 7"
        Case 1 'North
            action = "m 9"
        Case 2 'South
            action = "m 1"
        Case 3 'East
            action = "m 3"
        Case 4 'Turn <
            action = "<"
        Case 5 'Turn >
            action = ">"
            
    End Select
    frmBot.sckFurc.SendData action & vbLf
End Sub
