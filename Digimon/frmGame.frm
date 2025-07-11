VERSION 5.00
Begin VB.Form frmGame 
   Caption         =   "Form1"
   ClientHeight    =   14265
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   20070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   14265
   ScaleWidth      =   20070
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmGame 
      Caption         =   "Game"
      Height          =   14175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   19815
      Begin VB.PictureBox imgCard 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2355
         Index           =   15
         Left            =   11880
         ScaleHeight     =   157
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   113
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   8400
         Width           =   1695
      End
      Begin VB.PictureBox imgCard 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2355
         Index           =   14
         Left            =   9960
         ScaleHeight     =   157
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   113
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   8400
         Width           =   1695
      End
      Begin VB.PictureBox imgCard 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2355
         Index           =   13
         Left            =   8040
         ScaleHeight     =   157
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   113
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   8400
         Width           =   1695
      End
      Begin VB.PictureBox imgCard 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2355
         Index           =   12
         Left            =   6120
         ScaleHeight     =   157
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   113
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   8400
         Width           =   1695
      End
      Begin VB.PictureBox imgCard 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2355
         Index           =   11
         Left            =   11880
         ScaleHeight     =   157
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   113
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   6000
         Width           =   1695
      End
      Begin VB.PictureBox imgCard 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2355
         Index           =   10
         Left            =   9960
         ScaleHeight     =   157
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   113
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   6000
         Width           =   1695
      End
      Begin VB.PictureBox imgCard 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2355
         Index           =   9
         Left            =   8040
         ScaleHeight     =   157
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   113
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   6000
         Width           =   1695
      End
      Begin VB.PictureBox imgCard 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2355
         Index           =   8
         Left            =   6120
         ScaleHeight     =   157
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   113
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   6000
         Width           =   1695
      End
      Begin VB.PictureBox imgCard 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2355
         Index           =   7
         Left            =   11880
         ScaleHeight     =   157
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   113
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   2880
         Width           =   1695
      End
      Begin VB.PictureBox imgCard 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2355
         Index           =   6
         Left            =   9960
         ScaleHeight     =   157
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   113
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   2880
         Width           =   1695
      End
      Begin VB.PictureBox imgCard 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2355
         Index           =   5
         Left            =   8040
         ScaleHeight     =   157
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   113
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   2880
         Width           =   1695
      End
      Begin VB.PictureBox imgCard 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2355
         Index           =   4
         Left            =   6120
         ScaleHeight     =   157
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   113
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   2880
         Width           =   1695
      End
      Begin VB.PictureBox imgCard 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2355
         Index           =   3
         Left            =   11880
         ScaleHeight     =   157
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   113
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   480
         Width           =   1695
      End
      Begin VB.PictureBox imgCard 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2355
         Index           =   2
         Left            =   9960
         ScaleHeight     =   157
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   113
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   480
         Width           =   1695
      End
      Begin VB.PictureBox imgCard 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2355
         Index           =   1
         Left            =   8040
         ScaleHeight     =   157
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   113
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   480
         Width           =   1695
      End
      Begin VB.PictureBox imgCard 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2355
         Index           =   0
         Left            =   6120
         ScaleHeight     =   157
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   113
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   480
         Width           =   1695
      End
      Begin VB.PictureBox imgOffline 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2355
         Left            =   14040
         ScaleHeight     =   157
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   113
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   6000
         Width           =   1695
      End
      Begin VB.PictureBox imgView 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   7800
         Left            =   120
         ScaleHeight     =   520
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   374
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   240
         Width           =   5610
      End
      Begin VB.PictureBox imgOnline 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2355
         Left            =   14040
         Picture         =   "frmGame.frx":0000
         ScaleHeight     =   157
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   113
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   8400
         Width           =   1695
      End
      Begin VB.PictureBox imgHand 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2355
         Index           =   9
         Left            =   18000
         ScaleHeight     =   157
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   113
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   10320
         Width           =   1695
      End
      Begin VB.PictureBox imgHand 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2355
         Index           =   8
         Left            =   16200
         ScaleHeight     =   157
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   113
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   10320
         Width           =   1695
      End
      Begin VB.PictureBox imgHand 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2355
         Index           =   7
         Left            =   18000
         ScaleHeight     =   157
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   113
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   7920
         Width           =   1695
      End
      Begin VB.PictureBox imgHand 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2355
         Index           =   6
         Left            =   16200
         ScaleHeight     =   157
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   113
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   7920
         Width           =   1695
      End
      Begin VB.PictureBox imgHand 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2355
         Index           =   5
         Left            =   18000
         ScaleHeight     =   157
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   113
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   5520
         Width           =   1695
      End
      Begin VB.PictureBox imgHand 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2355
         Index           =   4
         Left            =   16200
         ScaleHeight     =   157
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   113
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   5520
         Width           =   1695
      End
      Begin VB.PictureBox imgHand 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2355
         Index           =   3
         Left            =   18000
         ScaleHeight     =   157
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   113
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   3120
         Width           =   1695
      End
      Begin VB.PictureBox imgHand 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2355
         Index           =   2
         Left            =   16200
         ScaleHeight     =   157
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   113
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   3120
         Width           =   1695
      End
      Begin VB.PictureBox imgHand 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2355
         Index           =   1
         Left            =   18000
         ScaleHeight     =   157
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   113
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   720
         Width           =   1695
      End
      Begin VB.PictureBox imgHand 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2355
         Index           =   0
         Left            =   16200
         ScaleHeight     =   157
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   113
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   720
         Width           =   1695
      End
      Begin VB.Frame frmGameInfo 
         Caption         =   "Game Info"
         Height          =   2655
         Left            =   120
         TabIndex        =   1
         Top             =   8040
         Width           =   5655
      End
   End
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Load frmCards
frmCards.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub frmGame_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgView.Picture = frmCards.imgViewClear.Picture
End Sub

Private Sub imgHand_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

Call StretchBlt(imgView.hdc, 0, 0, 374, 512, frmCards.imgDM(0).hdc, 0, 0, 374, 512, &HCC0020)

End Sub

Private Sub imgOnline_Click()

Call StretchBlt(imgHand(0).hdc, 0, 0, 113, 157, frmCards.imgDMsm(0).hdc, 0, 0, 113, 157, &HCC0020)

End Sub
