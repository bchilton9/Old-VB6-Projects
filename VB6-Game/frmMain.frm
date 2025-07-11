VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   Caption         =   "The Land's of Tidel"
   ClientHeight    =   9855
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12045
   LinkTopic       =   "Form1"
   Picture         =   "frmMain.frx":0000
   ScaleHeight     =   9855
   ScaleWidth      =   12045
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmBattle 
      Caption         =   "Battle"
      Height          =   5775
      Left            =   6480
      TabIndex        =   8
      Top             =   480
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame frmStore 
      Caption         =   "Store"
      Height          =   5775
      Left            =   8040
      TabIndex        =   7
      Top             =   960
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.Frame frmBank 
      Caption         =   "Bank"
      Height          =   5775
      Left            =   7680
      TabIndex        =   5
      Top             =   600
      Visible         =   0   'False
      Width           =   4695
      Begin VB.PictureBox picEquip 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   44
         Left            =   2160
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   52
         Top             =   2400
         Width           =   495
      End
      Begin VB.PictureBox picEquip 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   43
         Left            =   1560
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   51
         Top             =   2400
         Width           =   495
      End
      Begin VB.PictureBox picEquip 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   42
         Left            =   960
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   50
         Top             =   2400
         Width           =   495
      End
      Begin VB.PictureBox picEquip 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   41
         Left            =   360
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   49
         Top             =   2400
         Width           =   495
      End
      Begin VB.PictureBox picEquip 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   40
         Left            =   2160
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   48
         Top             =   1800
         Width           =   495
      End
      Begin VB.PictureBox picEquip 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   39
         Left            =   1560
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   47
         Top             =   1800
         Width           =   495
      End
      Begin VB.PictureBox picEquip 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   38
         Left            =   960
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   46
         Top             =   1800
         Width           =   495
      End
      Begin VB.PictureBox picEquip 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   37
         Left            =   360
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   45
         Top             =   1800
         Width           =   495
      End
      Begin VB.PictureBox picEquip 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   36
         Left            =   2160
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   44
         Top             =   1200
         Width           =   495
      End
      Begin VB.PictureBox picEquip 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   35
         Left            =   1560
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   43
         Top             =   1200
         Width           =   495
      End
      Begin VB.PictureBox picEquip 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   34
         Left            =   960
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   42
         Top             =   1200
         Width           =   495
      End
      Begin VB.PictureBox picEquip 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   33
         Left            =   360
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   41
         Top             =   1200
         Width           =   495
      End
      Begin VB.PictureBox picEquip 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   32
         Left            =   2160
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   40
         Top             =   600
         Width           =   495
      End
      Begin VB.PictureBox picEquip 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   31
         Left            =   1560
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   39
         Top             =   600
         Width           =   495
      End
      Begin VB.PictureBox picEquip 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   30
         Left            =   960
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   38
         Top             =   600
         Width           =   495
      End
      Begin VB.PictureBox picEquip 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   29
         Left            =   360
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   37
         Top             =   600
         Width           =   495
      End
   End
   Begin VB.PictureBox cmgInv3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   8400
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   77
      Top             =   7440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox cmgInv2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   7920
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   76
      Top             =   7440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox cmgInv 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   7440
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   75
      Top             =   7440
      Width           =   375
   End
   Begin VB.PictureBox cmdWho3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   8400
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   74
      Top             =   6960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox cmdWho2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   7920
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   73
      Top             =   6960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox cmdWho 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   7440
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   72
      Top             =   6960
      Width           =   375
   End
   Begin VB.PictureBox cmdSend3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   8400
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   71
      Top             =   6480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox cmdSend2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   7920
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   70
      Top             =   6480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox cmdSend 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   7440
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   69
      Top             =   6480
      Width           =   375
   End
   Begin VB.PictureBox picObjectSelectB 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   18780
      Left            =   2040
      Picture         =   "frmMain.frx":120C6
      ScaleHeight     =   18720
      ScaleWidth      =   5760
      TabIndex        =   67
      Top             =   9000
      Width           =   5820
   End
   Begin VB.PictureBox picObjectSelectMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   18780
      Left            =   1560
      Picture         =   "frmMain.frx":171108
      ScaleHeight     =   18720
      ScaleWidth      =   5760
      TabIndex        =   66
      Top             =   9000
      Width           =   5820
   End
   Begin VB.PictureBox picNpcSelectB 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   18780
      Left            =   1080
      Picture         =   "frmMain.frx":2D014A
      ScaleHeight     =   18720
      ScaleWidth      =   5760
      TabIndex        =   65
      Top             =   9000
      Width           =   5820
   End
   Begin VB.PictureBox picNpcSelectMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   18780
      Left            =   480
      Picture         =   "frmMain.frx":42F18C
      ScaleHeight     =   18720
      ScaleWidth      =   5760
      TabIndex        =   64
      Top             =   9000
      Width           =   5820
   End
   Begin VB.PictureBox picSelect 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   18720
      Left            =   120
      Picture         =   "frmMain.frx":58E1CE
      ScaleHeight     =   1248
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   384
      TabIndex        =   63
      Top             =   9000
      Width           =   5760
   End
   Begin VB.Frame frmBlank 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Blank"
      Height          =   5775
      Left            =   7080
      TabIndex        =   6
      Top             =   240
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.Frame frmInventory 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Inventory"
      Height          =   5775
      Left            =   0
      TabIndex        =   4
      Top             =   240
      Visible         =   0   'False
      Width           =   6855
      Begin VB.TextBox txtGold 
         Height          =   375
         Left            =   3480
         TabIndex        =   68
         Text            =   "Text1"
         Top             =   4440
         Width           =   1455
      End
      Begin VB.PictureBox cmdDistory3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   4200
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   62
         Top             =   5280
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox cmdDistory2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3720
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   61
         Top             =   5280
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox cmdDistory 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3240
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   60
         Top             =   5280
         Width           =   375
      End
      Begin VB.PictureBox cmdInspect3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2640
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   59
         Top             =   5280
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox cmdInspect2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2160
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   58
         Top             =   5280
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox cmdInspect 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1680
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   57
         Top             =   5280
         Width           =   375
      End
      Begin VB.PictureBox cmdCloseInventory3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1080
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   56
         Top             =   5280
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox cmdCloseInventory2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   600
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   55
         Top             =   5280
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox cmdCloseInventory 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   54
         Top             =   5280
         Width           =   375
      End
      Begin VB.PictureBox picEquip 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   28
         Left            =   2400
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   36
         Top             =   3720
         Width           =   495
      End
      Begin VB.PictureBox picEquip 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   27
         Left            =   1800
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   35
         Top             =   3720
         Width           =   495
      End
      Begin VB.PictureBox picEquip 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   26
         Left            =   1200
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   34
         Top             =   3720
         Width           =   495
      End
      Begin VB.PictureBox picEquip 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   25
         Left            =   600
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   33
         Top             =   3720
         Width           =   495
      End
      Begin VB.PictureBox picEquip 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   24
         Left            =   3000
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   32
         Top             =   3120
         Width           =   495
      End
      Begin VB.PictureBox picEquip 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   23
         Left            =   2400
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   31
         Top             =   3120
         Width           =   495
      End
      Begin VB.PictureBox picEquip 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   22
         Left            =   1800
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   30
         Top             =   3120
         Width           =   495
      End
      Begin VB.PictureBox picEquip 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   21
         Left            =   1200
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   29
         Top             =   3120
         Width           =   495
      End
      Begin VB.PictureBox picEquip 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   20
         Left            =   600
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   28
         Top             =   3120
         Width           =   495
      End
      Begin VB.PictureBox picEquip 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   19
         Left            =   3000
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   27
         Top             =   2520
         Width           =   495
      End
      Begin VB.PictureBox picEquip 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   18
         Left            =   2400
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   26
         Top             =   2520
         Width           =   495
      End
      Begin VB.PictureBox picEquip 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   17
         Left            =   1800
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   25
         Top             =   2520
         Width           =   495
      End
      Begin VB.PictureBox picEquip 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   16
         Left            =   1200
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   24
         Top             =   2520
         Width           =   495
      End
      Begin VB.PictureBox picEquip 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   15
         Left            =   600
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   23
         Top             =   2520
         Width           =   495
      End
      Begin VB.PictureBox picEquip 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   14
         Left            =   3000
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   22
         Top             =   1920
         Width           =   495
      End
      Begin VB.PictureBox picEquip 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   13
         Left            =   2400
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   21
         Top             =   1920
         Width           =   495
      End
      Begin VB.PictureBox picEquip 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   12
         Left            =   1800
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   20
         Top             =   1920
         Width           =   495
      End
      Begin VB.PictureBox picEquip 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   11
         Left            =   1200
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   19
         Top             =   1920
         Width           =   495
      End
      Begin VB.PictureBox picEquip 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   10
         Left            =   600
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   18
         Top             =   1920
         Width           =   495
      End
      Begin VB.PictureBox picEquip 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   9
         Left            =   3000
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   17
         Top             =   1320
         Width           =   495
      End
      Begin VB.PictureBox picEquip 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   8
         Left            =   2400
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   16
         Top             =   1320
         Width           =   495
      End
      Begin VB.PictureBox picEquip 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   7
         Left            =   1800
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   15
         Top             =   1320
         Width           =   495
      End
      Begin VB.PictureBox picEquip 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   6
         Left            =   1200
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   14
         Top             =   1320
         Width           =   495
      End
      Begin VB.PictureBox picEquip 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   5
         Left            =   600
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   13
         Top             =   1320
         Width           =   495
      End
      Begin VB.PictureBox picEquip 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   4
         Left            =   3000
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   12
         Top             =   720
         Width           =   495
      End
      Begin VB.PictureBox picEquip 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   3
         Left            =   2400
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   11
         Top             =   720
         Width           =   495
      End
      Begin VB.PictureBox picEquip 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   2
         Left            =   1800
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   10
         Top             =   720
         Width           =   495
      End
      Begin VB.PictureBox picEquip 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   1
         Left            =   1200
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   9
         Top             =   720
         Width           =   495
      End
      Begin VB.PictureBox picInvBack 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   0
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   78
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.Frame frmSign 
      Caption         =   "Sign"
      Height          =   1575
      Left            =   120
      TabIndex        =   3
      Top             =   3240
      Visible         =   0   'False
      Width           =   1455
      Begin VB.TextBox txtSign 
         Height          =   615
         Left            =   240
         TabIndex        =   53
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.TextBox txtSend 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   240
      TabIndex        =   2
      Top             =   8040
      Width           =   5535
   End
   Begin RichTextLib.RichTextBox txtChat 
      Height          =   1575
      Left            =   240
      TabIndex        =   1
      Top             =   6360
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   2778
      _Version        =   393217
      BackColor       =   0
      ReadOnly        =   -1  'True
      TextRTF         =   $"frmMain.frx":6ED212
   End
   Begin VB.PictureBox picBuffer 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   5775
      Left            =   240
      ScaleHeight     =   383
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   767
      TabIndex        =   0
      Top             =   240
      Width           =   11535
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCloseInventory_Click()
frmInventory.Visible = False
frmBlank.Visible = False
frmBank.Visible = False
frmStore.Visible = False
End Sub

Private Sub cmgInv_Click()
frmInventory.Visible = True
frmBlank.Visible = True
End Sub

Private Sub Form_Load()
loadskin "Defalt"
End Sub

Private Sub Form_unLoad(cancel As Integer)
frmTCP.Show
frmTCP.wskServer.SendData "logout" & pEnd
Unload Me
End Sub
