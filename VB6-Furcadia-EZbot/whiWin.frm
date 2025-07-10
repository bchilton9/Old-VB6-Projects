VERSION 5.00
Begin VB.Form whiWin 
   Caption         =   "Add Whisper"
   ClientHeight    =   3240
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8820
   LinkTopic       =   "Form1"
   ScaleHeight     =   3240
   ScaleWidth      =   8820
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox txtBack 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   4215
   End
   Begin VB.TextBox txtWhis 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "Whisper Back:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "to you. ]"
      Height          =   255
      Left            =   3480
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "[ Furre Whispers,"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "whiWin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSave_Click()
txtwhis = LCase(txtwhis)
Open "C:/EZbot/whisper.lst" For Append As #1
Write #1, txtwhis, txtBack
Close #1
whiWin.Hide
End Sub
