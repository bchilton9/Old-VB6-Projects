VERSION 5.00
Begin VB.Form frmabout 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   1995
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   1860
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1995
   ScaleWidth      =   1860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Close"
      Height          =   255
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Caption         =   "Unknown"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Interface:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Caption         =   "Byron Chilton"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Programing:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "Virtual Furre"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmabout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private ontop As New clsOnTop
Private Sub Form_Load()
  Set ontop = New clsOnTop
  ontop.MakeTopMost hWnd
End Sub
Private Sub cmdClose_Click()
Unload frmabout
End Sub
