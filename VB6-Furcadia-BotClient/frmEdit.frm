VERSION 5.00
Begin VB.Form frmEdit 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   1560
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   3360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtbotpass 
      BackColor       =   &H00FFC0C0&
      Height          =   285
      Left            =   1680
      TabIndex        =   3
      Top             =   360
      Width           =   1575
   End
   Begin VB.TextBox txtbotdesc 
      BackColor       =   &H00FFC0C0&
      Height          =   285
      Left            =   1680
      TabIndex        =   2
      Top             =   600
      Width           =   1575
   End
   Begin VB.TextBox txtbotcode 
      BackColor       =   &H00FFC0C0&
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      Top             =   840
      Width           =   1575
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Save"
      Height          =   255
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label lblNam 
      BackColor       =   &H00FFC0C0&
      Height          =   255
      Left            =   1680
      TabIndex        =   8
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Caption         =   "Bot Name:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Caption         =   "Bot Password:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Caption         =   "Bot Description:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Caption         =   "Bot ColorCode:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1455
   End
End
Attribute VB_Name = "frmEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private ontop As New clsOnTop
Private Sub Form_Load()
  Set ontop = New clsOnTop
  ontop.MakeTopMost hWnd
  lblNam = frmBot.bload
Open frmBot.bload & ".bot" For Input As #1
Input #1, bname
Input #1, bpass
Input #1, bdesc
Input #1, bcode
Close #1
txtbotpass = bpass
txtbotdesc = bdesc
txtbotcode = bcode
End Sub
Private Sub cmdSave_Click()
Open lblNam & ".bot" For Output As #1
Write #1, lblNam
Write #1, txtbotpass
Write #1, txtbotdesc
Write #1, txtbotcode
Close #1
Unload frmEdit
End Sub

