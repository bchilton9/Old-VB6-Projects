VERSION 5.00
Begin VB.Form frmItem 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "ITEM NAME"
   ClientHeight    =   2700
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5205
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   5205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   255
      Left            =   4440
      TabIndex        =   0
      Top             =   2400
      Width           =   735
   End
End
Attribute VB_Name = "frmItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
Me.Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
Me.Hide
End Sub
