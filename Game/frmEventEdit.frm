VERSION 5.00
Begin VB.Form frmEventEdit 
   Caption         =   "Event Editor"
   ClientHeight    =   5730
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8220
   LinkTopic       =   "Form1"
   ScaleHeight     =   5730
   ScaleWidth      =   8220
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load"
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add Event"
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox txtMapLoaded 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   2520
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtEvent 
      Height          =   2415
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   4815
   End
End
Attribute VB_Name = "frmEventEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

