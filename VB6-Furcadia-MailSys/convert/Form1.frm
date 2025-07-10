VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2625
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   2625
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Members"
      Height          =   1815
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2415
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   "C:\Program Files\Microsoft Visual Studio\VB98\projects\MailSys\convert\mailsys.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   0  'Table
         RecordSource    =   "Member"
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox txtStatus 
         DataField       =   "Stats"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   720
         TabIndex        =   4
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox txtDate 
         DataField       =   "Joined"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   720
         TabIndex        =   3
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtName 
         DataField       =   "Name"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   720
         TabIndex        =   2
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label7 
         Caption         =   "Stat's"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Date"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Name"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "start"
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   2040
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdStart_Click()

Open "members.txt" For Input As #1
Input #1, nam, num, rank
Do Until EOF(1)


Data1.Recordset.AddNew

txtName.Text = nam
txtDate.Text = "Unknown"
txtStatus.Text = "1"
Input #1, nam, num, rank

Loop
Close #1

End Sub
