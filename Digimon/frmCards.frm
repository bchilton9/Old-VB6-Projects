VERSION 5.00
Begin VB.Form frmCards 
   Caption         =   "Form1"
   ClientHeight    =   13725
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   16395
   LinkTopic       =   "Form1"
   ScaleHeight     =   13725
   ScaleWidth      =   16395
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox imgViewClear 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7800
      Left            =   9600
      ScaleHeight     =   520
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   374
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   480
      Width           =   5610
   End
   Begin VB.PictureBox imgDMsm 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2355
      Index           =   0
      Left            =   120
      Picture         =   "frmCards.frx":0000
      ScaleHeight     =   157
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   113
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   8760
      Width           =   1695
   End
   Begin VB.PictureBox imgDM 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7800
      Index           =   4
      Left            =   9600
      Picture         =   "frmCards.frx":698E
      ScaleHeight     =   520
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   374
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3240
      Width           =   5610
   End
   Begin VB.PictureBox imgDM 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7800
      Index           =   3
      Left            =   7200
      Picture         =   "frmCards.frx":250BA
      ScaleHeight     =   520
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   374
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2640
      Width           =   5610
   End
   Begin VB.PictureBox imgDM 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7800
      Index           =   2
      Left            =   5040
      Picture         =   "frmCards.frx":40106
      ScaleHeight     =   520
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   374
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1920
      Width           =   5610
   End
   Begin VB.PictureBox imgDM 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7800
      Index           =   1
      Left            =   2400
      Picture         =   "frmCards.frx":5C374
      ScaleHeight     =   520
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   374
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1320
      Width           =   5610
   End
   Begin VB.PictureBox imgDM 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7800
      Index           =   0
      Left            =   120
      Picture         =   "frmCards.frx":74218
      ScaleHeight     =   520
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   374
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   840
      Width           =   5610
   End
End
Attribute VB_Name = "frmCards"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
