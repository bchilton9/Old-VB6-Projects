VERSION 5.00
Begin VB.Form whisper 
   Caption         =   "Whispers"
   ClientHeight    =   4215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6600
   LinkTopic       =   "Form1"
   ScaleHeight     =   4215
   ScaleWidth      =   6600
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   3720
      TabIndex        =   9
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox txtDel 
      Height          =   285
      Left            =   3720
      TabIndex        =   7
      Top             =   2160
      Width           =   2175
   End
   Begin VB.TextBox txtWhis 
      Height          =   285
      Left            =   3720
      TabIndex        =   3
      Top             =   240
      Width           =   1935
   End
   Begin VB.TextBox txtBack 
      Height          =   285
      Left            =   2280
      TabIndex        =   2
      Top             =   960
      Width           =   4215
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add New"
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox txtwh 
      Height          =   3975
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "Use %1 to make your bot say the furrys name"
      Height          =   255
      Left            =   2760
      TabIndex        =   10
      Top             =   1320
      Width           =   3375
   End
   Begin VB.Label lbl4 
      Caption         =   "Whisper to delete:"
      Height          =   255
      Left            =   2280
      TabIndex        =   8
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Line Line1 
      X1              =   2160
      X2              =   6480
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Label Label1 
      Caption         =   "[ Furre Whispers,"
      Height          =   255
      Left            =   2400
      TabIndex        =   6
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "to you. ]"
      Height          =   255
      Left            =   5760
      TabIndex        =   5
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Whisper Back:"
      Height          =   255
      Left            =   2280
      TabIndex        =   4
      Top             =   720
      Width           =   1335
   End
End
Attribute VB_Name = "whisper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
txtWhis = LCase(txtWhis)
Open "C:/EZbot/whisper.lst" For Append As #1
Write #1, txtWhis, txtBack
Close #1
txtwh = txtwh & txtWhis & vbCrLf
txtWhis = ""
txtBack = ""
End Sub

Private Sub cmdDelete_Click()
txtDel = "Not active yet"
End Sub

Private Sub Form_Load()
On Error GoTo stoptrying
Open "C:/EZbot/whisper.lst" For Input As #1
Input #1, ame, back
txtwh = txtwh & ame & vbCrLf
Do Until EOF(1)
Input #1, ame, back
txtwh = txtwh & ame & vbCrLf
Loop
Close #1
Exit Sub
stoptrying:
Close #1
End Sub
