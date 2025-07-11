VERSION 5.00
Begin VB.Form Switch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Switch"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4425
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   4425
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClAl 
      Caption         =   "Clear All"
      Height          =   375
      Left            =   2280
      TabIndex        =   20
      Top             =   3960
      Width           =   1335
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "< Hide"
      Height          =   255
      Left            =   4560
      TabIndex        =   19
      Top             =   4080
      Width           =   615
   End
   Begin VB.CommandButton cmdGen 
      Caption         =   "Generate"
      Height          =   375
      Left            =   720
      TabIndex        =   18
      Top             =   3960
      Width           =   1335
   End
   Begin VB.CommandButton cmdAddOff 
      Caption         =   "Add"
      Height          =   255
      Left            =   3840
      TabIndex        =   17
      Top             =   2880
      Width           =   495
   End
   Begin VB.TextBox txtOff 
      Height          =   735
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   15
      Top             =   3120
      Width           =   4215
   End
   Begin VB.CommandButton cmdAddOn 
      Caption         =   "Add"
      Height          =   255
      Left            =   3840
      TabIndex        =   14
      Top             =   1800
      Width           =   495
   End
   Begin VB.TextBox txtOn 
      Height          =   735
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Top             =   2040
      Width           =   4215
   End
   Begin VB.TextBox txtCode 
      Height          =   3735
      Left            =   4560
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   360
      Width           =   4215
   End
   Begin VB.TextBox txtX 
      Height          =   285
      Left            =   2040
      TabIndex        =   5
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox txtY 
      Height          =   285
      Left            =   2760
      TabIndex        =   4
      Top             =   120
      Width           =   375
   End
   Begin VB.OptionButton optSwitch 
      Alignment       =   1  'Right Justify
      Height          =   255
      Index           =   0
      Left            =   2040
      TabIndex        =   3
      Top             =   720
      Value           =   -1  'True
      Width           =   255
   End
   Begin VB.OptionButton optSwitch 
      Alignment       =   1  'Right Justify
      Height          =   255
      Index           =   1
      Left            =   2040
      TabIndex        =   2
      Top             =   1320
      Width           =   255
   End
   Begin VB.PictureBox p161 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   2400
      Picture         =   "Switch.frx":0000
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   1
      Top             =   600
      Width           =   615
   End
   Begin VB.PictureBox p163 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   2400
      Picture         =   "Switch.frx":0515
      ScaleHeight     =   495
      ScaleWidth      =   735
      TabIndex        =   0
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "When Turned off:"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "When Turned on:"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label lblSwitch1 
      Alignment       =   1  'Right Justify
      Caption         =   "Switch Location:"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lblSwitch2 
      Alignment       =   2  'Center
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   10
      Top             =   120
      Width           =   255
   End
   Begin VB.Label lblSwitch3 
      Alignment       =   2  'Center
      Caption         =   "Y"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      TabIndex        =   9
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "Code:"
      Height          =   255
      Left            =   4560
      TabIndex        =   8
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lblSwitch4 
      Alignment       =   1  'Right Justify
      Caption         =   "Switch Type:"
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "Switch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim swon, swoff As String

Private Sub cmdAddOff_Click()
frmMain.txtTemp = "temp2.txt"
frmMain.txtFrom = "switchoff"
doAdd.Show
End Sub

Private Sub cmdAddOn_Click()
frmMain.txtTemp = "temp1.txt"
frmMain.txtFrom = "switchon"
doAdd.Show
End Sub

Private Sub cmdClAl_Click()
txtCode = ""
txtX = ""
txtY = ""
Switch.Width = "4515"
txtOn.Text = ""
txtOff.Text = ""
Open "temp1.txt" For Output As #1
Write #1, ""
Close #1
Open "temp2.txt" For Output As #1
Write #1, ""
Close #1
End Sub

Private Sub cmdClear_Click()
txtCode = ""
Switch.Width = "4515"
End Sub

Private Sub cmdGen_Click()
If txtX = "" Or txtY = "" Then
messg = MsgBox("Please enter the X,Y location!", vbOKOnly)
Else
pos = txtX & "," & txtY
Switch.Width = "8970"
txtCode = "SWITCH" & vbCrLf
txtCode = txtCode & "(0:7) When somebody moves into position (" & pos & ")," & vbCrLf
txtCode = txtCode & "    (1:3) and they move into object type " & swon & "," & vbCrLf
txtCode = txtCode & "    (3:2) at position (" & pos & ") on the map," & vbCrLf
txtCode = txtCode & "    (5:4) place object type " & swoff & "." & vbCrLf
Open "temp1.txt" For Input As #1
Do Until EOF(1)
Input #1, lin
If lin <> "" Then txtCode = txtCode & lin & vbCrLf
Loop
Close #1
txtCode = txtCode & "(0:7) When somebody moves into position (" & pos & ")," & vbCrLf
txtCode = txtCode & "    (1:3) and they move into object type " & swoff & "," & vbCrLf
txtCode = txtCode & "    (3:2) at position (" & pos & ") on the map," & vbCrLf
txtCode = txtCode & "    (5:4) place object type " & swon & "." & vbCrLf
Open "temp2.txt" For Input As #1
Do Until EOF(1)
Input #1, lin
If lin <> "" Then txtCode = txtCode & lin & vbCrLf
Loop
Close #1
End If
End Sub

Private Sub Form_Load()
swon = "161"
swoff = "162"
End Sub

Private Sub optSwitch_Click(Index As Integer)
Select Case Index
Case 0
swon = "161"
swoff = "162"
Case 1
swon = "163"
swoff = "164"
End Select
End Sub
