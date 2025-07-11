VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dragon Speek 2000"
   ClientHeight    =   1545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4290
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   4290
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtFrom 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   3120
      Width           =   1455
   End
   Begin VB.TextBox txtTemp 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   2880
      Width           =   1455
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Switch"
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Furcadian Dragon Speek Generator!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOpen_Click()
Open "temp1.txt" For Output As #1
Write #1, ""
Close #1
Open "temp2.txt" For Output As #1
Write #1, ""
Close #1
txtFrom = ""
txtTemp = ""

Switch.Show
End Sub
