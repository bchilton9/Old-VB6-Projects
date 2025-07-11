VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmBot 
   Caption         =   "The Land's of Tidel Bots"
   ClientHeight    =   5265
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   5625
   LinkTopic       =   "Form1"
   ScaleHeight     =   5265
   ScaleWidth      =   5625
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmBoT 
      Caption         =   "Dwain"
      Height          =   4815
      Index           =   2
      Left            =   120
      TabIndex        =   85
      Top             =   360
      Visible         =   0   'False
      Width           =   5415
      Begin VB.CommandButton cmdConnect 
         Caption         =   "&Connect"
         Height          =   375
         Index           =   2
         Left            =   2760
         TabIndex        =   124
         Top             =   3240
         Width           =   1215
      End
      Begin VB.CommandButton cmdDisconnect 
         Caption         =   "&Disconnect"
         Height          =   375
         Index           =   2
         Left            =   2760
         TabIndex        =   123
         Top             =   3720
         Width           =   1215
      End
      Begin VB.CommandButton cmdNW 
         Caption         =   "NW"
         Height          =   495
         Index           =   2
         Left            =   120
         TabIndex        =   122
         Top             =   3480
         Width           =   615
      End
      Begin VB.CommandButton cmdNE 
         Caption         =   "NE"
         Height          =   495
         Index           =   2
         Left            =   720
         TabIndex        =   121
         Top             =   3480
         Width           =   615
      End
      Begin VB.CommandButton cmdSW 
         Caption         =   "SW"
         Height          =   495
         Index           =   2
         Left            =   120
         TabIndex        =   120
         Top             =   3960
         Width           =   615
      End
      Begin VB.CommandButton cmdSE 
         Caption         =   "SE"
         Height          =   495
         Index           =   2
         Left            =   720
         TabIndex        =   119
         Top             =   3960
         Width           =   615
      End
      Begin VB.CommandButton cmdLay 
         Caption         =   "&Lay"
         Height          =   495
         Index           =   2
         Left            =   1440
         TabIndex        =   118
         Top             =   3240
         Width           =   615
      End
      Begin VB.CommandButton cmdWho 
         Caption         =   "&Who"
         Height          =   495
         Index           =   2
         Left            =   2040
         TabIndex        =   117
         Top             =   3240
         Width           =   615
      End
      Begin VB.CommandButton cmdGet 
         Caption         =   "&Get"
         Height          =   495
         Index           =   2
         Left            =   1440
         TabIndex        =   116
         Top             =   3720
         Width           =   615
      End
      Begin VB.CommandButton cmduse 
         Caption         =   "&Use"
         Height          =   495
         Index           =   2
         Left            =   2040
         TabIndex        =   115
         Top             =   3720
         Width           =   615
      End
      Begin VB.CommandButton cmdGoAlleg 
         Caption         =   "&Allegria"
         Height          =   375
         Index           =   2
         Left            =   2760
         TabIndex        =   114
         Top             =   4200
         Width           =   1215
      End
      Begin VB.CommandButton cmdturnr 
         Caption         =   "< Turn"
         Height          =   495
         Index           =   2
         Left            =   1440
         TabIndex        =   113
         Top             =   4200
         Width           =   615
      End
      Begin VB.CommandButton cmdturnl 
         Caption         =   "Turn >"
         Height          =   495
         Index           =   2
         Left            =   2040
         TabIndex        =   112
         Top             =   4200
         Width           =   615
      End
      Begin VB.Frame Frame4 
         Caption         =   "System"
         Height          =   1335
         Index           =   2
         Left            =   4080
         TabIndex        =   107
         Top             =   360
         Width           =   1215
         Begin VB.CheckBox chkServCode 
            Caption         =   "SCode"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   111
            Top             =   480
            Width           =   975
         End
         Begin VB.CheckBox chkFollow 
            Caption         =   "Follow"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   110
            Top             =   720
            Width           =   975
         End
         Begin VB.CheckBox chkWhisp 
            Caption         =   "Whispers"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   109
            Top             =   960
            Value           =   1  'Checked
            Width           =   975
         End
         Begin VB.CheckBox chkServtxt 
            Caption         =   "SText"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   108
            Top             =   240
            Value           =   1  'Checked
            Width           =   855
         End
      End
      Begin VB.TextBox txtFromFurc 
         Height          =   2655
         Index           =   2
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   106
         Top             =   480
         Width           =   3855
      End
      Begin VB.Frame Frame1 
         Caption         =   "Other"
         Height          =   615
         Index           =   2
         Left            =   4080
         TabIndex        =   104
         Top             =   1680
         Width           =   1215
         Begin VB.CheckBox chkbar 
            Caption         =   "Bartend"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   105
            Top             =   240
            Value           =   1  'Checked
            Width           =   975
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Fighting"
         Height          =   975
         Index           =   2
         Left            =   4080
         TabIndex        =   86
         Top             =   2280
         Visible         =   0   'False
         Width           =   1215
         Begin VB.TextBox fnum 
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   127
            Text            =   "0"
            Top             =   1200
            Width           =   975
         End
         Begin VB.TextBox turn 
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   103
            Text            =   "0"
            Top             =   3000
            Width           =   975
         End
         Begin VB.TextBox f2class 
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   102
            Top             =   2880
            Width           =   975
         End
         Begin VB.TextBox f1class 
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   101
            Top             =   2760
            Width           =   975
         End
         Begin VB.TextBox f2armor 
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   100
            Top             =   2640
            Width           =   975
         End
         Begin VB.TextBox f1armor 
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   99
            Top             =   2520
            Width           =   975
         End
         Begin VB.TextBox f2weapon 
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   98
            Top             =   2400
            Width           =   975
         End
         Begin VB.TextBox f1weapon 
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   97
            Top             =   2280
            Width           =   975
         End
         Begin VB.TextBox f2mana 
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   96
            Top             =   2160
            Width           =   975
         End
         Begin VB.TextBox f1mana 
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   95
            Top             =   2040
            Width           =   975
         End
         Begin VB.TextBox f2hp 
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   94
            Top             =   1920
            Width           =   975
         End
         Begin VB.TextBox f1hp 
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   93
            Top             =   1800
            Width           =   975
         End
         Begin VB.TextBox f2lvl 
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   92
            Top             =   1680
            Width           =   975
         End
         Begin VB.TextBox f1lvl 
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   91
            Top             =   1560
            Width           =   975
         End
         Begin VB.TextBox f2name 
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   90
            Top             =   1440
            Width           =   975
         End
         Begin VB.TextBox f1name 
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   89
            Top             =   1320
            Width           =   975
         End
         Begin VB.CheckBox chkFight 
            Caption         =   "On/Off"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   88
            Top             =   240
            UseMaskColor    =   -1  'True
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.CommandButton cmdrest 
            Caption         =   "&Reset"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   87
            Top             =   600
            Width           =   975
         End
      End
      Begin VB.Label lblConected 
         Caption         =   "False"
         Height          =   255
         Index           =   2
         Left            =   1080
         TabIndex        =   126
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Conected 
         Caption         =   "Connected:"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   125
         Top             =   240
         Width           =   855
      End
   End
   Begin MSWinsockLib.Winsock sckFurc 
      Index           =   0
      Left            =   120
      Top             =   5400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer StayOnline 
      Index           =   0
      Interval        =   60000
      Left            =   1080
      Top             =   5400
   End
   Begin VB.Timer tfight 
      Index           =   0
      Interval        =   2000
      Left            =   600
      Top             =   5400
   End
   Begin VB.Frame frmBoT 
      Caption         =   "Valka"
      Height          =   4815
      Index           =   1
      Left            =   120
      TabIndex        =   26
      Top             =   360
      Visible         =   0   'False
      Width           =   5415
      Begin VB.Frame Frame3 
         Caption         =   "Fighting"
         Height          =   975
         Index           =   1
         Left            =   4080
         TabIndex        =   67
         Top             =   1680
         Width           =   1215
         Begin VB.TextBox fnum 
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   128
            Text            =   "0"
            Top             =   3120
            Width           =   975
         End
         Begin VB.CommandButton cmdrest 
            Caption         =   "&Reset"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   84
            Top             =   600
            Width           =   975
         End
         Begin VB.CheckBox chkFight 
            Caption         =   "On/Off"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   83
            Top             =   240
            Value           =   1  'Checked
            Width           =   855
         End
         Begin VB.TextBox f1name 
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   82
            Top             =   1320
            Width           =   975
         End
         Begin VB.TextBox f2name 
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   81
            Top             =   1680
            Width           =   975
         End
         Begin VB.TextBox f1lvl 
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   80
            Top             =   1440
            Width           =   975
         End
         Begin VB.TextBox f2lvl 
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   79
            Top             =   1680
            Width           =   975
         End
         Begin VB.TextBox f1hp 
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   78
            Top             =   2280
            Width           =   975
         End
         Begin VB.TextBox f2hp 
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   77
            Top             =   2160
            Width           =   975
         End
         Begin VB.TextBox f1mana 
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   76
            Top             =   2160
            Width           =   975
         End
         Begin VB.TextBox f2mana 
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   75
            Top             =   2160
            Width           =   975
         End
         Begin VB.TextBox f1weapon 
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   74
            Top             =   2280
            Width           =   975
         End
         Begin VB.TextBox f2weapon 
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   73
            Top             =   2400
            Width           =   975
         End
         Begin VB.TextBox f1armor 
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   72
            Top             =   2520
            Width           =   975
         End
         Begin VB.TextBox f2armor 
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   71
            Top             =   2640
            Width           =   975
         End
         Begin VB.TextBox f1class 
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   70
            Top             =   2760
            Width           =   975
         End
         Begin VB.TextBox f2class 
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   69
            Top             =   2880
            Width           =   975
         End
         Begin VB.TextBox turn 
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   68
            Text            =   "1"
            Top             =   3000
            Width           =   975
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Other"
         Height          =   615
         Index           =   1
         Left            =   4080
         TabIndex        =   50
         Top             =   2640
         Visible         =   0   'False
         Width           =   1215
         Begin VB.CheckBox chkbar 
            Caption         =   "Bartend"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   51
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.TextBox txtFromFurc 
         Height          =   2655
         Index           =   1
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   45
         Top             =   480
         Width           =   3855
      End
      Begin VB.Frame Frame4 
         Caption         =   "System"
         Height          =   1335
         Index           =   1
         Left            =   4080
         TabIndex        =   40
         Top             =   360
         Width           =   1215
         Begin VB.CheckBox chkServtxt 
            Caption         =   "SText"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   44
            Top             =   240
            Value           =   1  'Checked
            Width           =   855
         End
         Begin VB.CheckBox chkWhisp 
            Caption         =   "Whispers"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   43
            Top             =   960
            Value           =   1  'Checked
            Width           =   975
         End
         Begin VB.CheckBox chkFollow 
            Caption         =   "Follow"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   42
            Top             =   720
            Width           =   975
         End
         Begin VB.CheckBox chkServCode 
            Caption         =   "SCode"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   41
            Top             =   480
            Width           =   975
         End
      End
      Begin VB.CommandButton cmdturnl 
         Caption         =   "Turn >"
         Height          =   495
         Index           =   1
         Left            =   2040
         TabIndex        =   39
         Top             =   4200
         Width           =   615
      End
      Begin VB.CommandButton cmdturnr 
         Caption         =   "< Turn"
         Height          =   495
         Index           =   1
         Left            =   1440
         TabIndex        =   38
         Top             =   4200
         Width           =   615
      End
      Begin VB.CommandButton cmdGoAlleg 
         Caption         =   "&Allegria"
         Height          =   375
         Index           =   1
         Left            =   2760
         TabIndex        =   37
         Top             =   4200
         Width           =   1215
      End
      Begin VB.CommandButton cmduse 
         Caption         =   "&Use"
         Height          =   495
         Index           =   1
         Left            =   2040
         TabIndex        =   36
         Top             =   3720
         Width           =   615
      End
      Begin VB.CommandButton cmdGet 
         Caption         =   "&Get"
         Height          =   495
         Index           =   1
         Left            =   1440
         TabIndex        =   35
         Top             =   3720
         Width           =   615
      End
      Begin VB.CommandButton cmdWho 
         Caption         =   "&Who"
         Height          =   495
         Index           =   1
         Left            =   2040
         TabIndex        =   34
         Top             =   3240
         Width           =   615
      End
      Begin VB.CommandButton cmdLay 
         Caption         =   "&Lay"
         Height          =   495
         Index           =   1
         Left            =   1440
         TabIndex        =   33
         Top             =   3240
         Width           =   615
      End
      Begin VB.CommandButton cmdSE 
         Caption         =   "SE"
         Height          =   495
         Index           =   1
         Left            =   720
         TabIndex        =   32
         Top             =   3960
         Width           =   615
      End
      Begin VB.CommandButton cmdSW 
         Caption         =   "SW"
         Height          =   495
         Index           =   1
         Left            =   120
         TabIndex        =   31
         Top             =   3960
         Width           =   615
      End
      Begin VB.CommandButton cmdNE 
         Caption         =   "NE"
         Height          =   495
         Index           =   1
         Left            =   720
         TabIndex        =   30
         Top             =   3480
         Width           =   615
      End
      Begin VB.CommandButton cmdNW 
         Caption         =   "NW"
         Height          =   495
         Index           =   1
         Left            =   120
         TabIndex        =   29
         Top             =   3480
         Width           =   615
      End
      Begin VB.CommandButton cmdDisconnect 
         Caption         =   "&Disconnect"
         Height          =   375
         Index           =   1
         Left            =   2760
         TabIndex        =   28
         Top             =   3720
         Width           =   1215
      End
      Begin VB.CommandButton cmdConnect 
         Caption         =   "&Connect"
         Height          =   375
         Index           =   1
         Left            =   2760
         TabIndex        =   27
         Top             =   3240
         Width           =   1215
      End
      Begin VB.Label Conected 
         Caption         =   "Connected:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   47
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblConected 
         Caption         =   "False"
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   46
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame frmBoT 
      Caption         =   "Pena"
      Height          =   4815
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   5415
      Begin VB.Frame Frame1 
         Caption         =   "Other"
         Height          =   615
         Index           =   0
         Left            =   4080
         TabIndex        =   48
         Top             =   2640
         Visible         =   0   'False
         Width           =   1215
         Begin VB.CheckBox chkbar 
            Caption         =   "Bartend"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   49
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Fighting"
         Height          =   975
         Index           =   0
         Left            =   4080
         TabIndex        =   21
         Top             =   1680
         Visible         =   0   'False
         Width           =   1215
         Begin VB.TextBox fnum 
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   129
            Text            =   "0"
            Top             =   1800
            Width           =   975
         End
         Begin VB.TextBox turn 
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   66
            Text            =   "0"
            Top             =   3000
            Width           =   975
         End
         Begin VB.TextBox f2class 
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   65
            Top             =   2880
            Width           =   975
         End
         Begin VB.TextBox f1class 
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   64
            Top             =   2760
            Width           =   975
         End
         Begin VB.TextBox f2armor 
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   63
            Top             =   2640
            Width           =   975
         End
         Begin VB.TextBox f1armor 
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   62
            Top             =   2520
            Width           =   975
         End
         Begin VB.TextBox f2weapon 
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   61
            Top             =   2400
            Width           =   975
         End
         Begin VB.TextBox f1weapon 
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   60
            Top             =   2280
            Width           =   975
         End
         Begin VB.TextBox f2mana 
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   59
            Top             =   2160
            Width           =   975
         End
         Begin VB.TextBox f1mana 
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   58
            Top             =   2040
            Width           =   975
         End
         Begin VB.TextBox f2hp 
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   57
            Top             =   1920
            Width           =   975
         End
         Begin VB.TextBox f1hp 
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   56
            Top             =   1800
            Width           =   975
         End
         Begin VB.TextBox f2lvl 
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   55
            Top             =   1680
            Width           =   975
         End
         Begin VB.TextBox f1lvl 
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   54
            Top             =   1560
            Width           =   975
         End
         Begin VB.TextBox f2name 
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   53
            Top             =   1440
            Width           =   975
         End
         Begin VB.TextBox f1name 
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   52
            Top             =   1320
            Width           =   975
         End
         Begin VB.CheckBox chkFight 
            Caption         =   "On/Off"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   23
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton cmdrest 
            Caption         =   "&Reset"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   22
            Top             =   600
            Width           =   975
         End
      End
      Begin VB.CommandButton cmdConnect 
         Caption         =   "&Connect"
         Height          =   375
         Index           =   0
         Left            =   2760
         TabIndex        =   20
         Top             =   3240
         Width           =   1215
      End
      Begin VB.CommandButton cmdDisconnect 
         Caption         =   "&Disconnect"
         Height          =   375
         Index           =   0
         Left            =   2760
         TabIndex        =   19
         Top             =   3720
         Width           =   1215
      End
      Begin VB.CommandButton cmdNW 
         Caption         =   "NW"
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   18
         Top             =   3480
         Width           =   615
      End
      Begin VB.CommandButton cmdNE 
         Caption         =   "NE"
         Height          =   495
         Index           =   0
         Left            =   720
         TabIndex        =   17
         Top             =   3480
         Width           =   615
      End
      Begin VB.CommandButton cmdSW 
         Caption         =   "SW"
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   3960
         Width           =   615
      End
      Begin VB.CommandButton cmdSE 
         Caption         =   "SE"
         Height          =   495
         Index           =   0
         Left            =   720
         TabIndex        =   15
         Top             =   3960
         Width           =   615
      End
      Begin VB.CommandButton cmdLay 
         Caption         =   "&Lay"
         Height          =   495
         Index           =   0
         Left            =   1440
         TabIndex        =   14
         Top             =   3240
         Width           =   615
      End
      Begin VB.CommandButton cmdWho 
         Caption         =   "&Who"
         Height          =   495
         Index           =   0
         Left            =   2040
         TabIndex        =   13
         Top             =   3240
         Width           =   615
      End
      Begin VB.CommandButton cmdGet 
         Caption         =   "&Get"
         Height          =   495
         Index           =   0
         Left            =   1440
         TabIndex        =   12
         Top             =   3720
         Width           =   615
      End
      Begin VB.CommandButton cmduse 
         Caption         =   "&Use"
         Height          =   495
         Index           =   0
         Left            =   2040
         TabIndex        =   11
         Top             =   3720
         Width           =   615
      End
      Begin VB.CommandButton cmdGoAlleg 
         Caption         =   "&Allegria"
         Height          =   375
         Index           =   0
         Left            =   2760
         TabIndex        =   10
         Top             =   4200
         Width           =   1215
      End
      Begin VB.CommandButton cmdturnr 
         Caption         =   "< Turn"
         Height          =   495
         Index           =   0
         Left            =   1440
         TabIndex        =   9
         Top             =   4200
         Width           =   615
      End
      Begin VB.CommandButton cmdturnl 
         Caption         =   "Turn >"
         Height          =   495
         Index           =   0
         Left            =   2040
         TabIndex        =   8
         Top             =   4200
         Width           =   615
      End
      Begin VB.Frame Frame4 
         Caption         =   "System"
         Height          =   1335
         Index           =   0
         Left            =   4080
         TabIndex        =   3
         Top             =   360
         Width           =   1215
         Begin VB.CheckBox chkServCode 
            Caption         =   "SCode"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   7
            Top             =   480
            Width           =   975
         End
         Begin VB.CheckBox chkFollow 
            Caption         =   "Follow"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   6
            Top             =   720
            Width           =   975
         End
         Begin VB.CheckBox chkWhisp 
            Caption         =   "Whispers"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   5
            Top             =   960
            Value           =   1  'Checked
            Width           =   975
         End
         Begin VB.CheckBox chkServtxt 
            Caption         =   "SText"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Value           =   1  'Checked
            Width           =   855
         End
      End
      Begin VB.TextBox txtFromFurc 
         Height          =   2655
         Index           =   0
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   480
         Width           =   3855
      End
      Begin VB.Label lblConected 
         Caption         =   "False"
         Height          =   255
         Index           =   0
         Left            =   1080
         TabIndex        =   25
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Conected 
         Caption         =   "Connected:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   855
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   5295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   9340
      MultiRow        =   -1  'True
      MultiSelect     =   -1  'True
      Separators      =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Pena (Main)"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Valka (Arena 1)"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Dwain (Bartender)"
            ImageVarType    =   2
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "frmBot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public BotName, BotPass, descrip, ColorCode, fName, mnum, lvl, clas, gold, xp, weap, armo, sone, stwo, sthe, sfor, sfiv, hp, man

Private Sub Form_Load()
bot(0).Minute = 0

Dim i As Long

  For i = 1 To MaxBot
    Load sckFurc(i)
    Load StayOnline(i)
    Load tfight(i)
    bot(i).Minute = 0
  Next i

bot(0).Name = "Pena"
bot(0).Pass = "0519aa"
bot(0).Desc = "Tidel Governer. Im here to make sure everything is ran right here in The Lands of Tidel"
bot(0).Color = "56J B999=9 " & Chr(34) & "! " & Chr(34) & "!"

bot(1).Name = "Valka"
bot(1).Pass = "0519aa"
bot(1).Desc = "Arena 1 Officent."
bot(1).Color = "! G2+88888!#!!#!"
bot(1).LookA = ") 1"
bot(1).LookB = "( 3"
bot(1).TriggerA = ") + 6"
bot(1).TriggerB = ") * 8"

bot(2).Name = "Dwain"
bot(2).Pass = "0519aa"
bot(2).Desc = "I am the Bartender here at The Tidel Pub. Whisper me HELP for more info."
bot(2).Color = Chr(34) & Chr(34) & "H2G88888!#!! !"

End Sub

Private Sub TabStrip1_Click()
If TabStrip1.SelectedItem.Index = 1 Then
frmBoT(0).Visible = True
frmBoT(1).Visible = False
frmBoT(2).Visible = False
ElseIf TabStrip1.SelectedItem.Index = 2 Then
frmBoT(0).Visible = False
frmBoT(1).Visible = True
frmBoT(2).Visible = False
ElseIf TabStrip1.SelectedItem.Index = 3 Then
frmBoT(0).Visible = False
frmBoT(1).Visible = False
frmBoT(2).Visible = True
End If
End Sub

Private Sub chkServCode_Click(Index As Integer)
If chkServCode(Index) = 1 Then
chkServtxt(Index) = 2
chkServtxt(Index).Enabled = False
End If
If chkServCode(Index) = 0 Then
chkServtxt(Index) = 1
chkServtxt(Index).Enabled = True
End If
End Sub

Private Sub cmdGet_Click(Index As Integer)
sckFurc(Index).SendData "get" & vbLf
End Sub

Private Sub cmdGoAlleg_Click(Index As Integer)
sckFurc(Index).SendData "goalleg" & vbLf
End Sub

Private Sub cmdlie_Click(Index As Integer)
sckFurc(Index).SendData "lie" & vbLf
End Sub

Private Sub cmdNE_Click(Index As Integer)
sckFurc(Index).SendData "m 9" & vbLf
End Sub

Private Sub cmdNW_Click(Index As Integer)
sckFurc(Index).SendData "m 7" & vbLf
End Sub

Private Sub cmdrest_Click(Index As Integer)
reset Index
sckFurc(Index).SendData Chr(34) & "emitloud Arena " & Index & " was reset!" & vbLf
End Sub

Sub reset(Index As Integer)
If Index = 1 Then
sckFurc(Index).SendData "m 7" & vbLf & "m 3" & vbLf
End If
   f1name(Index).Text = ""
   f1lvl(Index).Text = ""
   f1hp(Index).Text = ""
   f1mana(Index).Text = ""
   f1weapon(Index).Text = ""
   f1armor(Index).Text = ""
   f1class(Index).Text = ""
   
   f2name(Index).Text = ""
   f2lvl(Index).Text = ""
   f2hp(Index).Text = ""
   f2mana(Index).Text = ""
   f2weapon(Index).Text = ""
   f2armor(Index).Text = ""
   f2class(Index).Text = ""
   
   fnum(Index).Text = "0"
End Sub

Private Sub cmdSE_Click(Index As Integer)
sckFurc(Index).SendData "m 3" & vbLf
End Sub

Private Sub cmdSW_Click(Index As Integer)
sckFurc(Index).SendData "m 1" & vbLf
End Sub

Private Sub cmdturnl_Click(Index As Integer)
sckFurc(Index).SendData ">" & vbLf
End Sub

Private Sub cmdturnr_Click(Index As Integer)
sckFurc(Index).SendData "<" & vbLf
End Sub

Private Sub cmduse_Click(Index As Integer)
sckFurc(Index).SendData "use" & vbLf
End Sub

Private Sub cmdWho_Click(Index As Integer)
sckFurc(Index).SendData "who" & vbLf
End Sub

Private Sub cmdConnect_Click(Index As Integer)
If lblConected(Index).Caption = False Then
sckFurc(Index).RemoteHost = "64.191.51.88"
sckFurc(Index).RemotePort = "6000"
sckFurc(Index).Connect
lblConected(Index).Caption = "Connecting..."
bot(Index).lastwalk = "none"
txtFromFurc(Index) = txtFromFurc(Index) & "Connecting..." & vbCrLf
End If
End Sub

Private Sub cmdDisconnect_Click(Index As Integer)
If lblConected(Index).Caption = "True" Or lblConected(Index).Caption = "Connecting..." Then
sckFurc(Index).Close
lblConected(Index).Caption = "False"
txtFromFurc(Index) = txtFromFurc(Index) & "Disconected" & vbCrLf
End If
End Sub

Private Sub sckFurc_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim s As String
Dim Packet() As String
Dim i As Long


  sckFurc(Index).GetData s
  Packet = Split(s, vbLf)
  For i = 0 To UBound(Packet) - 1
    RealText Packet(i), Index
  Next i
End Sub
Sub RealText(Txt, Index As Integer)
Dim tmsg, Furre, NMsg, Msg, frl, ord, Order As String
On Error Resume Next
If chkServtxt(Index).Value = Checked Or chkServtxt(Index).Enabled = False Then
'If the checkbox with the Server Code is checked then you see all of the server
'code
If chkServCode(Index).Value = Checked Then txtFromFurc(0) = txtFromFurc(0) & Txt & vbCrLf
'If Left(Txt, 1) = "/" Then txtFromFurc = txtFromFurc & Left(Txt, Len(Txt) - 6) & vbCrLf
'If the checkbox with the Server Code label is not checked you do not see any of
'the server code. You'll only see what you would see in the Furcadia client.
If chkServCode(Index).Value = Unchecked Then
If Left(Txt, 1) = "(" Then txtFromFurc(Index) = txtFromFurc(Index) & Right(Txt, Len(Txt) - 1) & vbCrLf
End If
End If
'When the text "END" is sent to the bot, the bot sends the information to login
'to Furcadia
If Txt = "END" Then
sckFurc(Index).SendData "connect " & bot(Index).Name & " " & bot(Index).Pass & vbLf & "color " & bot(Index).Color & vbLf & "desc " & bot(Index).Desc & " [Uptime: 0 Minute(s)]" & vbLf
lblConected(Index).Caption = "True"
End If
'When your bot enters a dream, it sends "vascodagama" to Furcadia to let it into
'the dream.
If Txt = "]ccmarbled.pcx" Then
chkFollow(Index).Value = 0
chkWhisp(Index).Value = 1
sckFurc(Index).SendData "vascodagama" & vbLf
sckFurc(Index).SendData "use" & vbLf
End If
'When someone whispers the bot, it gets there name and message and calls the
'DoWhisper(Furre, Msg) sub which is used to respond to whispers.
If chkWhisp(Index).Value = Checked Then
If Left(Txt, 3) = "([ " And Right(Txt, 10) = " to you. ]" Then
    tmsg = Split(Txt, " whispers, " & Chr(34), 2)
    Furre = Right(tmsg(0), Len(tmsg(0)) - 3)
    NMsg = Left(tmsg(1), Len(tmsg(1)) - 11)
    Msg = LCase(NMsg)
    DoWhisper Furre, Msg, Index
End If
End If 'chkWhisp

'make the bot follow it owner
If chkFollow(Index).Value = Checked Then
If Left(Txt, 11) = "/!!8++<<)<<" Then
        frl = Mid(Txt, 17, Len(Txt) - 0)
        bot(Index).whatwalk = Mid(frl, 1, Len(frl) - 4)
        'whatwalk = LCase(wwalk)
    dowalk bot(Index).whatwalk, bot(Index).lastwalk, Index
End If
End If 'chkFollow


'make the bot act like a bartender
If chkbar(Index).Value = Checked Then
If Left(Txt, 1) = "(" And Right(Txt, 1) = "#" Then
    tmsg = Split(Txt, ": #", Len(Txt), 2)
    Furre = Right(tmsg(0), Len(tmsg(0)) - 1)
    ord = Left(tmsg(1), Len(tmsg(1)) - 1)
    Order = LCase(ord)
    doserve Furre, Order, Index
End If
End If 'chkBar

'watch for fight data
If chkFight(Index).Value = Checked Then

        
        'If Txt = ") + 6" Then sckFurc(Index).SendData "l  ) 1" & vbLf
        'If Txt = ") * 8" Then sckFurc(Index).SendData "l  ( 3" & vbLf
        If Txt = bot(Index).TriggerA Then sckFurc(Index).SendData "l  " & bot(Index).LookA & vbLf
        If Txt = bot(Index).TriggerB Then sckFurc(Index).SendData "l  " & bot(Index).LookB & vbLf
            
            
            'Gets the furres name when the bot looks at them.
            If Left(Txt, 10) = "((You see " Then
            Furre = Mid(Txt, 11, Len(Txt) - 12)
            
            Open "members.txt" For Input As #1
            Input #1, fName, mnum
            Do Until (fName = Furre) Or (EOF(1))
            Input #1, fName, mnum
            Loop
            Close #1
            
            If fName = Furre Then
                Open "memfiles\" & mnum & ".txt" For Input As #1
                Input #1, fName, mnum, lvl, clas, gold, xp, weap, armo, sone, stwo, sthe, sfor, sfiv
                Close #1
                
    If clas = "Wizard" Then
    hp = lvl * 15
    man = lvl * 20
    End If
    If clas = "Fighter" Then
    hp = lvl * 20
    man = 0
    End If
    If clas = "Thief" Then
    hp = lvl * 15
    man = lvl * 10
    End If
    If clas = "Paladin" Then
    hp = lvl * 20
    man = lvl * 10
    End If
    If clas = "Priest" Then
    hp = lvl * 10
    man = lvl * 10
    End If
                
    If fnum(Index).Text = "0" Then
                        
    fnum(Index).Text = "1"

   f1name(Index).Text = Furre
   f1lvl(Index).Text = lvl
   f1hp(Index).Text = hp
   f1mana(Index).Text = man
   f1weapon(Index).Text = weap
   f1armor(Index).Text = armo
   f1class(Index).Text = clas

                            
    ElseIf fnum(Index).Text = 1 Then
    fnum(Index).Text = "2"

   f2name(Index).Text = Furre
   f2lvl(Index).Text = lvl
   f2hp(Index).Text = hp
   f2mana(Index).Text = man
   f2weapon(Index).Text = weap
   f2armor(Index).Text = armo
   f2class(Index).Text = clas

    End If
                        
                        
If weap = 0 Then weap = "Paws"
If weap = 1 Then weap = "Dagger"
If weap = 2 Then weap = "Knife"
If weap = 3 Then weap = "Hand ax"
If weap = 4 Then weap = "Quarterstaff"
If weap = 5 Then weap = "Spear"
If weap = 6 Then weap = "Warhammer"
If weap = 7 Then weap = "Battle ax"
If weap = 8 Then weap = "Morneing Star"
If weap = 9 Then weap = "Flail"
If weap = 10 Then weap = "Mace"
If weap = 11 Then weap = "Broad Sword"
If weap = 12 Then weap = "Short Bow"
If weap = 13 Then weap = "Crossbow"
If weap = 14 Then weap = "Shord Sword"
If weap = 15 Then weap = "Long Sword"
If weap = 16 Then weap = "TwoHand Sword"

If armo = 0 Then armo = "Fir"
If armo = 1 Then armo = "Padded"
If armo = 2 Then armo = "Leather"
If armo = 3 Then armo = "Chain Mail"
If armo = 4 Then armo = "Splint Mail"
If armo = 5 Then armo = "Ring Mail"
If armo = 6 Then armo = "Scale Mail"
If armo = 7 Then armo = "Banded Mail"
If armo = 8 Then armo = "Plate Mail"
        sckFurc(Index).SendData Chr(34) & "emit " & Furre & " The " & clas & " of level " & lvl & " Welding " & weap & " and " & armo & " has entered the fighting area." & vbLf
            
            
            
            
            If fnum(Index).Text = "2" Then
                sckFurc(Index).SendData Chr(34) & "emit let the fight begine" & vbLf
                turn(Index).Text = 1
                dofdream turn, Index
            End If
            
        Else
            sckFurc(Index).SendData Chr(34) & "emit " & Furre & " is not a registered member." & vbLf
        reset Index
        End If
        
        End If
        End If


End Sub

Private Sub StayOnline_Timer(Index As Integer)
'Each minute the timer is set off. The Minute variable is increased by one. Your
'bot changes its desc to add the Minute which is an Uptimer.
If lblConected(Index).Caption = "True" Then
bot(Index).Minute = bot(Index).Minute + 1
sckFurc(Index).SendData "desc " & bot(Index).Desc & " [Uptime: " & bot(Index).Minute & " Minute(s)]" & vbLf
End If
End Sub

Private Sub tfight_Timer(Index As Integer)

        If fnum(Index).Text = "2" Then
        dofdream turn, Index
        End If
End Sub

Private Sub txtFromFurc_Change(Index As Integer)
txtFromFurc(Index).SelStart = Len(txtFromFurc(Index))
If Len(txtFromFurc(Index)) > 10000 Then txtFromFurc(Index) = Right(txtFromFurc(Index), 9000)
End Sub
