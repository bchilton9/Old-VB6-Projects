VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00800000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "by hazim rihani"
   ClientHeight    =   1995
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4650
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1995
   ScaleWidth      =   4650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FF0000&
      Caption         =   "Exit"
      Height          =   375
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FF0000&
      Caption         =   "About"
      Height          =   375
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FF0000&
      ForeColor       =   &H80000009&
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   720
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FF0000&
      ForeColor       =   &H80000009&
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Text            =   "hazim.exe"
      Top             =   120
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF0000&
      Caption         =   "add to start up"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "File path :"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "File name :"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''
''by hazim hani rihani''
'' hazim_3@hotmail.com''
''www.87.8m.net       ''
'' tested win xp      ''
''''''''''''''''''''''''

Private Function RegWrite(Key1, SValue As String)

Set WSHShell = CreateObject("WScript.Shell")
WSHShell.RegWrite Key1, SValue

End Function


Private Sub Command1_Click()
 
RegWrite "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Run\" & Text1.Text, Text2.Text

End Sub

Private Sub Command2_Click()
MsgBox "by hazim rihani"
End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub Form_Load()

End Sub
