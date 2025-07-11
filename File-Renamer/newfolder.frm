VERSION 5.00
Begin VB.Form newfolder 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create New Folder"
   ClientHeight    =   975
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4470
   Icon            =   "newfolder.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   975
   ScaleWidth      =   4470
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Create"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   4215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Text            =   "Enter New Folder Name"
      Top             =   240
      Width           =   3135
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "New Folder:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "newfolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim direrctory1
directory1 = Len(fileeditor.Dir1.Path)
If directory1 <= 3 Then
directory = "C:\" & fileeditor.File1.FileName
Else
directory = fileeditor.Dir1.Path & "\" & fileeditor.File1.FileName
End If
MkDir (directory & Text1.Text)

fileeditor.Dir1.Refresh
Unload Me
End Sub

Private Sub Text1_Click()
Text1.Text = ""
End Sub
