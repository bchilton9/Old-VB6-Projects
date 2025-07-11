VERSION 5.00
Begin VB.Form viewerr 
   BackColor       =   &H00404040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Errors"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   6390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   6390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txterr 
      BackColor       =   &H00404000&
      ForeColor       =   &H008080FF&
      Height          =   3015
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   6135
   End
End
Attribute VB_Name = "viewerr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Open "C:\mailsys\errorlog.txt" For Input As #6
    Line Input #6, dat
    txterr = txterr & dat & vbCrLf
    Do Until EOF(6)
    Line Input #6, dat
    txterr = txterr & dat & vbCrLf
    Loop
    Close #6
End Sub
