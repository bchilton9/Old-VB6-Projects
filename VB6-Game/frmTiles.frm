VERSION 5.00
Begin VB.Form frmTiles 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2205
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2745
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   2745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picSetA 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   615
      TabIndex        =   0
      Top             =   0
      Width           =   615
   End
End
Attribute VB_Name = "frmTiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

  With picSetA
    .Width = 12 * 32
    .Height = 23 * 32
    .Picture = LoadPicture(App.Path + "\tiles\a.bmp")
  End With
  
End Sub
