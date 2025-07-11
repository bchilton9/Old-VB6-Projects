VERSION 5.00
Begin VB.Form frmopen 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   1425
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   1425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdLoad 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Load"
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1680
      Width           =   975
   End
   Begin VB.ListBox lstBots 
      Height          =   1425
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmopen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Open "bots.ini" For Input As #1
Do Until EOF(1)
Input #1, nam
lstBots.AddItem nam
Loop
Close #1
End Sub
Private Sub cmdLoad_Click()
If lstBots.Text <> "" Then
frmBot.bload = lstBots.Text
frmBot.lblLoad = lstBots.Text
Open "settings.ini" For Output As #1
Write #1, lstBots.Text
Write #1, frmBot.frcHost
Write #1, frmBot.frcPort
Close #1
If lstBots.Text <> "Defalt" Then
frmBot.cmdConnect.Enabled = True
frmBot.cmdecit.Enabled = True
frmBot.cmmd.Enabled = True
Else
frmBot.cmdConnect.Enabled = False
frmBot.cmdecit.Enabled = False
frmBot.cmmd.Enabled = False
End If
Unload frmopen
Else
box = MsgBox("Please chose a bot to load.", vbOKOnly)
End If
End Sub

