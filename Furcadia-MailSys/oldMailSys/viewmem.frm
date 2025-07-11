VERSION 5.00
Begin VB.Form viewmem 
   BackColor       =   &H8000000A&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3285
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3510
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   3510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtmem 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000007&
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
