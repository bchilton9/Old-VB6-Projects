VERSION 5.00
Begin VB.Form frmDebugger 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Debugger"
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7995
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   7995
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkEnable 
      Caption         =   "Enable"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   4920
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   4935
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   8055
   End
End
Attribute VB_Name = "frmDebugger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Text1_Change()
Text1.SelStart = Len(Text1)
'If Len(Text1) > 10000 Then Text1 = Right(Text1, 9000)
End Sub


Private Sub Form_Unload(Cancel As Integer)

    Unload Me

End Sub
