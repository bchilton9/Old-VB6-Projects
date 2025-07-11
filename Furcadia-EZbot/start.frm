VERSION 5.00
Begin VB.Form start 
   Caption         =   "EZbot"
   ClientHeight    =   3825
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4455
   LinkTopic       =   "Form1"
   ScaleHeight     =   3825
   ScaleWidth      =   4455
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtCode 
      Height          =   285
      Left            =   1800
      TabIndex        =   8
      Top             =   1440
      Width           =   2055
   End
   Begin VB.TextBox txtDesc 
      Height          =   285
      Left            =   1800
      TabIndex        =   7
      Top             =   1080
      Width           =   2055
   End
   Begin VB.TextBox txtPass 
      Height          =   285
      Left            =   1800
      TabIndex        =   6
      Top             =   720
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      Begin VB.TextBox txtRace 
         Height          =   285
         Left            =   1680
         TabIndex        =   18
         Top             =   2040
         Width           =   2055
      End
      Begin VB.TextBox txtSex 
         Height          =   285
         Left            =   1680
         TabIndex        =   17
         Top             =   1680
         Width           =   2055
      End
      Begin VB.TextBox txtPort 
         Height          =   285
         Left            =   1680
         TabIndex        =   12
         Top             =   2760
         Width           =   2055
      End
      Begin VB.TextBox txtAddr 
         Height          =   285
         Left            =   1680
         TabIndex        =   11
         Top             =   2400
         Width           =   2055
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Connect"
         Height          =   375
         Left            =   2160
         TabIndex        =   10
         Top             =   3120
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Save"
         Height          =   375
         Left            =   720
         TabIndex        =   9
         Top             =   3120
         Width           =   1335
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   1680
         TabIndex        =   5
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Race:"
         Height          =   255
         Left            =   480
         TabIndex        =   16
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Sex:"
         Height          =   255
         Left            =   360
         TabIndex        =   15
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Port:"
         Height          =   255
         Left            =   600
         TabIndex        =   14
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Server:"
         Height          =   255
         Left            =   480
         TabIndex        =   13
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Bot's ColorCode:"
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Bot's Description:"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Bot's Password:"
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Bot's Name:"
         Height          =   255
         Left            =   720
         TabIndex        =   1
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "start"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Open "C:/EZbot/setting.ini" For Output As #1
Write #1, txtName
Write #1, txtPass
Write #1, txtCode
Write #1, txtSex
Write #1, txtRace
Write #1, txtDesc
Write #1, txtAddr
Write #1, txtPort
Close #1
box = MsgBox("Setting's have been saved.", vbOKOnly, "Saved")
End Sub

Private Sub Command2_Click()
Load frmBot
frmBot.Show
start.Hide
End Sub

Private Sub Form_Load()
Open "C:/EZbot/setting.ini" For Input As #1
Input #1, botNam
Input #1, botPas
Input #1, botCod
Input #1, botSex
Input #1, botRace
Input #1, botDes
Input #1, furAddr
Input #1, furPort
Close #1
txtName = botNam
txtPass = botPas
txtCode = botCod
txtSex = botSex
txtRace = botRace
txtDesc = botDes
txtAddr = furAddr
txtPort = furPort
End Sub
