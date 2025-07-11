VERSION 5.00
Begin VB.Form frmtex 
   Caption         =   "Tile Selector"
   ClientHeight    =   9150
   ClientLeft      =   8430
   ClientTop       =   1305
   ClientWidth     =   6525
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   610
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   435
   Visible         =   0   'False
   Begin VB.PictureBox char 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   495
      Index           =   15
      Left            =   3360
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   79
      Top             =   9840
      Width           =   495
   End
   Begin VB.PictureBox char 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   495
      Index           =   14
      Left            =   2760
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   78
      Top             =   9840
      Width           =   495
   End
   Begin VB.PictureBox char 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   495
      Index           =   13
      Left            =   2160
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   77
      Top             =   9840
      Width           =   495
   End
   Begin VB.PictureBox char 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   495
      Index           =   12
      Left            =   1560
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   76
      Top             =   9840
      Width           =   495
   End
   Begin VB.PictureBox char 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   495
      Index           =   11
      Left            =   960
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   75
      Top             =   9840
      Width           =   495
   End
   Begin VB.PictureBox char 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   495
      Index           =   10
      Left            =   360
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   74
      Top             =   9840
      Width           =   495
   End
   Begin VB.PictureBox tex 
      Height          =   495
      Index           =   63
      Left            =   480
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   73
      Top             =   7920
      Width           =   495
   End
   Begin VB.PictureBox tex 
      Height          =   495
      Index           =   62
      Left            =   5280
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   72
      Top             =   6840
      Width           =   495
   End
   Begin VB.PictureBox tex 
      Height          =   495
      Index           =   61
      Left            =   4680
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   71
      Top             =   6840
      Width           =   495
   End
   Begin VB.PictureBox tex 
      Height          =   495
      Index           =   60
      Left            =   4080
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   70
      Top             =   6840
      Width           =   495
   End
   Begin VB.PictureBox tex 
      Height          =   495
      Index           =   59
      Left            =   3480
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   69
      Top             =   6840
      Width           =   495
   End
   Begin VB.PictureBox tex 
      Height          =   495
      Index           =   58
      Left            =   2880
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   68
      Top             =   6840
      Width           =   495
   End
   Begin VB.PictureBox tex 
      Height          =   495
      Index           =   57
      Left            =   2280
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   67
      Top             =   6840
      Width           =   495
   End
   Begin VB.PictureBox tex 
      Height          =   495
      Index           =   56
      Left            =   1680
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   66
      Top             =   6840
      Width           =   495
   End
   Begin VB.PictureBox tex 
      Height          =   495
      Index           =   55
      Left            =   1080
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   65
      Top             =   6840
      Width           =   495
   End
   Begin VB.PictureBox tex 
      Height          =   495
      Index           =   54
      Left            =   480
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   64
      Top             =   6840
      Width           =   495
   End
   Begin VB.PictureBox tex 
      Height          =   495
      Index           =   53
      Left            =   5280
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   63
      Top             =   5760
      Width           =   495
   End
   Begin VB.PictureBox tex 
      Height          =   495
      Index           =   52
      Left            =   4680
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   62
      Top             =   5760
      Width           =   495
   End
   Begin VB.PictureBox tex 
      Height          =   495
      Index           =   51
      Left            =   4080
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   61
      Top             =   5760
      Width           =   495
   End
   Begin VB.PictureBox char 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   495
      Index           =   9
      Left            =   5760
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   60
      Top             =   9240
      Width           =   495
   End
   Begin VB.PictureBox char 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   495
      Index           =   8
      Left            =   5160
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   59
      Top             =   9240
      Width           =   495
   End
   Begin VB.PictureBox tex 
      Height          =   495
      Index           =   50
      Left            =   3480
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   58
      Top             =   5760
      Width           =   495
   End
   Begin VB.PictureBox tex 
      Height          =   495
      Index           =   49
      Left            =   2880
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   57
      Top             =   5760
      Width           =   495
   End
   Begin VB.PictureBox tex 
      Height          =   495
      Index           =   48
      Left            =   2280
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   56
      Top             =   5760
      Width           =   495
   End
   Begin VB.PictureBox tex 
      Height          =   495
      Index           =   47
      Left            =   1680
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   55
      Top             =   5760
      Width           =   495
   End
   Begin VB.PictureBox tex 
      Height          =   495
      Index           =   46
      Left            =   1080
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   54
      Top             =   5760
      Width           =   495
   End
   Begin VB.PictureBox tex 
      Height          =   495
      Index           =   45
      Left            =   480
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   53
      Top             =   5760
      Width           =   495
   End
   Begin VB.PictureBox tex 
      Height          =   495
      Index           =   44
      Left            =   5280
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   52
      Top             =   4680
      Width           =   495
   End
   Begin VB.PictureBox tex 
      Height          =   495
      Index           =   43
      Left            =   4680
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   51
      Top             =   4680
      Width           =   495
   End
   Begin VB.PictureBox tex 
      Height          =   495
      Index           =   42
      Left            =   4080
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   50
      Top             =   4680
      Width           =   495
   End
   Begin VB.PictureBox tex 
      Height          =   495
      Index           =   41
      Left            =   3480
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   49
      Top             =   4680
      Width           =   495
   End
   Begin VB.PictureBox tex 
      Height          =   495
      Index           =   40
      Left            =   2880
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   48
      Top             =   4680
      Width           =   495
   End
   Begin VB.PictureBox tex 
      Height          =   495
      Index           =   39
      Left            =   2280
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   47
      Top             =   4680
      Width           =   495
   End
   Begin VB.PictureBox tex 
      Height          =   495
      Index           =   38
      Left            =   1680
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   46
      Top             =   4680
      Width           =   495
   End
   Begin VB.PictureBox tex 
      Height          =   495
      Index           =   37
      Left            =   1080
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   45
      Top             =   4680
      Width           =   495
   End
   Begin VB.PictureBox tex 
      Height          =   495
      Index           =   36
      Left            =   480
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   44
      Top             =   4680
      Width           =   495
   End
   Begin VB.PictureBox tex 
      Height          =   495
      Index           =   35
      Left            =   5280
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   43
      Top             =   3600
      Width           =   495
   End
   Begin VB.PictureBox tex 
      Height          =   495
      Index           =   34
      Left            =   4680
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   42
      Top             =   3600
      Width           =   495
   End
   Begin VB.PictureBox tex 
      Height          =   495
      Index           =   33
      Left            =   4080
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   41
      Top             =   3600
      Width           =   495
   End
   Begin VB.PictureBox tex 
      Height          =   495
      Index           =   32
      Left            =   3480
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   40
      Top             =   3600
      Width           =   495
   End
   Begin VB.PictureBox tex 
      Height          =   495
      Index           =   31
      Left            =   2880
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   39
      Top             =   3600
      Width           =   495
   End
   Begin VB.PictureBox tex 
      Height          =   495
      Index           =   30
      Left            =   2280
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   38
      Top             =   3600
      Width           =   495
   End
   Begin VB.PictureBox tex 
      Height          =   495
      Index           =   29
      Left            =   1680
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   37
      Top             =   3600
      Width           =   495
   End
   Begin VB.PictureBox tex 
      Height          =   495
      Index           =   28
      Left            =   1080
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   36
      Top             =   3600
      Width           =   495
   End
   Begin VB.PictureBox tex 
      Height          =   495
      Index           =   27
      Left            =   480
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   35
      Top             =   3600
      Width           =   495
   End
   Begin VB.PictureBox tex 
      Height          =   495
      Index           =   26
      Left            =   5280
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   34
      Top             =   2520
      Width           =   495
   End
   Begin VB.PictureBox tex 
      Height          =   495
      Index           =   25
      Left            =   4680
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   33
      Top             =   2520
      Width           =   495
   End
   Begin VB.PictureBox tex 
      Height          =   495
      Index           =   24
      Left            =   4080
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   32
      Top             =   2520
      Width           =   495
   End
   Begin VB.PictureBox tex 
      Height          =   495
      Index           =   23
      Left            =   3480
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   31
      Top             =   2520
      Width           =   495
   End
   Begin VB.PictureBox tex 
      Height          =   495
      Index           =   22
      Left            =   2880
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   30
      Top             =   2520
      Width           =   495
   End
   Begin VB.PictureBox tex 
      Height          =   495
      Index           =   21
      Left            =   2280
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   29
      Top             =   2520
      Width           =   495
   End
   Begin VB.PictureBox tex 
      Height          =   495
      Index           =   20
      Left            =   1680
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   28
      Top             =   2520
      Width           =   495
   End
   Begin VB.PictureBox char 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   495
      Index           =   7
      Left            =   4560
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   27
      Top             =   9240
      Width           =   495
   End
   Begin VB.PictureBox char 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   495
      Index           =   6
      Left            =   3960
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   26
      Top             =   9240
      Width           =   495
   End
   Begin VB.PictureBox char 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   495
      Index           =   5
      Left            =   3360
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   25
      Top             =   9240
      Width           =   495
   End
   Begin VB.PictureBox char 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   495
      Index           =   4
      Left            =   2760
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   24
      Top             =   9240
      Width           =   495
   End
   Begin VB.PictureBox char 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   495
      Index           =   3
      Left            =   2160
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   23
      Top             =   9240
      Width           =   495
   End
   Begin VB.PictureBox char 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   495
      Index           =   2
      Left            =   1560
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   22
      Top             =   9240
      Width           =   495
   End
   Begin VB.PictureBox char 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   495
      Index           =   1
      Left            =   960
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   21
      Top             =   9240
      Width           =   495
   End
   Begin VB.PictureBox char 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   495
      Index           =   0
      Left            =   360
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   20
      Top             =   9240
      Width           =   495
   End
   Begin VB.PictureBox tex 
      Height          =   495
      Index           =   19
      Left            =   1080
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   19
      Top             =   2520
      Width           =   495
   End
   Begin VB.PictureBox tex 
      Height          =   495
      Index           =   18
      Left            =   480
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   18
      Top             =   2520
      Width           =   495
   End
   Begin VB.PictureBox tex 
      Height          =   495
      Index           =   17
      Left            =   5280
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   17
      Top             =   1440
      Width           =   495
   End
   Begin VB.PictureBox tex 
      Height          =   495
      Index           =   16
      Left            =   4680
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   16
      Top             =   1440
      Width           =   495
   End
   Begin VB.PictureBox tex 
      Height          =   495
      Index           =   15
      Left            =   4080
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   15
      Top             =   1440
      Width           =   495
   End
   Begin VB.PictureBox tex 
      Height          =   495
      Index           =   14
      Left            =   3480
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   14
      Top             =   1440
      Width           =   495
   End
   Begin VB.PictureBox tex 
      Height          =   495
      Index           =   13
      Left            =   2880
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   13
      Top             =   1440
      Width           =   495
   End
   Begin VB.PictureBox tex 
      Height          =   495
      Index           =   12
      Left            =   2280
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   12
      Top             =   1440
      Width           =   495
   End
   Begin VB.PictureBox tex 
      Height          =   495
      Index           =   11
      Left            =   1680
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   11
      Top             =   1440
      Width           =   495
   End
   Begin VB.PictureBox tex 
      Height          =   495
      Index           =   10
      Left            =   1080
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   10
      Top             =   1440
      Width           =   495
   End
   Begin VB.PictureBox tex 
      Height          =   495
      Index           =   9
      Left            =   480
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   9
      Top             =   1440
      Width           =   495
   End
   Begin VB.PictureBox tex 
      Height          =   495
      Index           =   8
      Left            =   5280
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   8
      Top             =   360
      Width           =   495
   End
   Begin VB.PictureBox tex 
      Height          =   495
      Index           =   7
      Left            =   4680
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   7
      Top             =   360
      Width           =   495
   End
   Begin VB.PictureBox tex 
      Height          =   495
      Index           =   6
      Left            =   4080
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   6
      Top             =   360
      Width           =   495
   End
   Begin VB.PictureBox tex 
      Height          =   495
      Index           =   5
      Left            =   3480
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   5
      Top             =   360
      Width           =   495
   End
   Begin VB.PictureBox tex 
      Height          =   495
      Index           =   4
      Left            =   2880
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   4
      Top             =   360
      Width           =   495
   End
   Begin VB.PictureBox tex 
      Height          =   495
      Index           =   3
      Left            =   2280
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   3
      Top             =   360
      Width           =   495
   End
   Begin VB.PictureBox tex 
      Height          =   495
      Index           =   2
      Left            =   1680
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   2
      Top             =   360
      Width           =   495
   End
   Begin VB.PictureBox tex 
      Height          =   495
      Index           =   1
      Left            =   1080
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   1
      Top             =   360
      Width           =   495
   End
   Begin VB.PictureBox tex 
      Height          =   495
      Index           =   0
      Left            =   480
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   0
      Top             =   360
      Width           =   495
   End
   Begin VB.Shape ball 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   600
      Shape           =   3  'Circle
      Top             =   960
      Width           =   255
   End
End
Attribute VB_Name = "frmtex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub tex_Click(Index As Integer)
On Error Resume Next
cur = Index
ball.Top = tex(cur).Top + 40
ball.Left = tex(cur).Left + 10


End Sub
