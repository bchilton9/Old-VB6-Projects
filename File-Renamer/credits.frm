VERSION 5.00
Begin VB.Form Credits 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Credits"
   ClientHeight    =   2970
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5175
   Icon            =   "credits.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   5175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   600
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   3720
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000D&
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2400
      Width           =   975
   End
End
Attribute VB_Name = "Credits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub Command1_Click()
End
End Sub


Private Sub Form_Activate()
Call Credits
End Sub

Private Sub Form_Load()
Text1.Text = "Created by Misael Morales, Copy, Rename, Delete Any Files as you Wish.. Enjoy it...  Email: johnsurfer21@hotmail.com"
End Sub

Private Sub Form_Unload(Cancel As Integer)

Set Form1 = Nothing

End Sub


Private Sub Credits()



start:

For i = 3100 To -800 Step -15

DoEvents
Text1.Top = i
DoEvents
Sleep (50)

Next i
DoEvents
GoTo start

End Sub
