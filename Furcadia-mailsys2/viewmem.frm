VERSION 5.00
Begin VB.Form viewmem 
   BackColor       =   &H00404040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Members"
   ClientHeight    =   3285
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   3480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   3480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtmem 
      BackColor       =   &H00404000&
      ForeColor       =   &H008080FF&
      Height          =   3015
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "viewmem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Open "C:\mailsys\members.txt" For Input As #6
    Line Input #6, dat
    txtmem = txtmem & dat & vbCrLf
    Do Until EOF(6)
    Line Input #6, dat
    txtmem = txtmem & dat & vbCrLf
    Loop
    Close #6
End Sub
