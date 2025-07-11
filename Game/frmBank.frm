VERSION 5.00
Begin VB.Form frmBank 
   BackColor       =   &H008080FF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8385
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11520
   LinkTopic       =   "Form1"
   ScaleHeight     =   8385
   ScaleWidth      =   11520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close Bank"
      Height          =   375
      Left            =   3840
      TabIndex        =   0
      Top             =   2400
      Width           =   1215
   End
End
Attribute VB_Name = "frmBank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
Unload Me
End Sub
