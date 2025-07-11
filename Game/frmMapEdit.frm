VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Map Editor"
   ClientHeight    =   10620
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15060
   Icon            =   "frmMapEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   708
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1004
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkNpc 
      Caption         =   "Npcs"
      Height          =   255
      Left            =   8280
      TabIndex        =   84
      Top             =   8640
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.PictureBox picBoxMask32 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   12720
      Picture         =   "frmMapEdit.frx":0E42
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   81
      Top             =   7680
      Width           =   540
   End
   Begin VB.PictureBox picBox32 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   12120
      Picture         =   "frmMapEdit.frx":1A84
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   80
      Top             =   7680
      Width           =   540
   End
   Begin VB.PictureBox picNpcSelectMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   18780
      Left            =   1680
      Picture         =   "frmMapEdit.frx":26C8
      ScaleHeight     =   18720
      ScaleWidth      =   5760
      TabIndex        =   78
      Top             =   9360
      Width           =   5820
   End
   Begin VB.PictureBox picNpcSelectB 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   18780
      Left            =   1320
      Picture         =   "frmMapEdit.frx":16170A
      ScaleHeight     =   18720
      ScaleWidth      =   5760
      TabIndex        =   79
      Top             =   9360
      Width           =   5820
   End
   Begin VB.PictureBox picObjectSelectMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   18780
      Left            =   840
      Picture         =   "frmMapEdit.frx":2C074C
      ScaleHeight     =   18720
      ScaleWidth      =   5760
      TabIndex        =   63
      Top             =   9360
      Width           =   5820
   End
   Begin VB.Frame frmNpc 
      Caption         =   "Npcs"
      Height          =   2775
      Left            =   1080
      TabIndex        =   72
      Top             =   2160
      Visible         =   0   'False
      Width           =   7935
      Begin VB.PictureBox picNpc 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   7320
         ScaleHeight     =   31
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   31
         TabIndex        =   76
         Top             =   240
         Width           =   495
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2415
         Left            =   120
         ScaleHeight     =   2415
         ScaleWidth      =   6015
         TabIndex        =   73
         Top             =   240
         Width           =   6015
         Begin VB.PictureBox picNpcSelect 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   18720
            Left            =   0
            Picture         =   "frmMapEdit.frx":41F78E
            ScaleHeight     =   1248
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   384
            TabIndex        =   75
            Top             =   0
            Width           =   5760
         End
         Begin VB.VScrollBar scrlNpc 
            Height          =   2415
            LargeChange     =   30
            Left            =   5760
            Max             =   510
            Min             =   1
            SmallChange     =   15
            TabIndex        =   74
            Top             =   0
            Value           =   1
            Width           =   255
         End
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Preview:"
         Height          =   255
         Left            =   6240
         TabIndex        =   77
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdAddMap 
      Caption         =   "Add New Map"
      Height          =   375
      Left            =   9960
      TabIndex        =   71
      Top             =   8520
      Width           =   1455
   End
   Begin VB.FileListBox fileMaps 
      Height          =   1455
      Left            =   9960
      Normal          =   0   'False
      Pattern         =   "map*.dat"
      TabIndex        =   70
      Top             =   5880
      Width           =   1455
   End
   Begin VB.CheckBox chkText 
      Caption         =   "Text"
      Height          =   255
      Left            =   8280
      TabIndex        =   69
      Top             =   7920
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.CheckBox chkobjunder 
      Caption         =   "Objects (Under)"
      Height          =   255
      Left            =   8280
      TabIndex        =   67
      Top             =   8400
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.CheckBox chkObjover 
      Caption         =   "Objects (Over)"
      Height          =   255
      Left            =   8280
      TabIndex        =   66
      Top             =   8160
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.CheckBox chkGrid 
      Caption         =   "Grid"
      Height          =   255
      Left            =   8280
      TabIndex        =   65
      Top             =   7680
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.PictureBox picObjectSelectB 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   18780
      Left            =   480
      Picture         =   "frmMapEdit.frx":57E7D0
      ScaleHeight     =   18720
      ScaleWidth      =   5760
      TabIndex        =   64
      Top             =   9360
      Width           =   5820
   End
   Begin VB.PictureBox picBox 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   12120
      Picture         =   "frmMapEdit.frx":6DD812
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   58
      Top             =   7320
      Width           =   300
   End
   Begin VB.PictureBox picBoxMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   12720
      Picture         =   "frmMapEdit.frx":6DDB56
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   57
      Top             =   7320
      Width           =   300
   End
   Begin VB.PictureBox picGridMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   12720
      Picture         =   "frmMapEdit.frx":6DDE9A
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   56
      Top             =   6720
      Width           =   540
   End
   Begin VB.PictureBox picGrid 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   12120
      Picture         =   "frmMapEdit.frx":6DEADC
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   55
      Top             =   6720
      Width           =   540
   End
   Begin VB.PictureBox picBackSelectB 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   18720
      Left            =   120
      Picture         =   "frmMapEdit.frx":6DF720
      ScaleHeight     =   1248
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   384
      TabIndex        =   54
      Top             =   9360
      Width           =   5760
   End
   Begin VB.Frame frmObjects 
      Caption         =   "Objects"
      Height          =   2775
      Left            =   2160
      TabIndex        =   48
      Top             =   2040
      Visible         =   0   'False
      Width           =   7935
      Begin VB.OptionButton optOver 
         Caption         =   "Under"
         Height          =   255
         Index           =   1
         Left            =   6240
         TabIndex        =   62
         Top             =   1440
         Width           =   1095
      End
      Begin VB.OptionButton optOver 
         Caption         =   "Over"
         Height          =   255
         Index           =   0
         Left            =   6240
         TabIndex        =   61
         Top             =   1200
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.PictureBox picObject 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   7320
         ScaleHeight     =   31
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   31
         TabIndex        =   52
         Top             =   240
         Width           =   495
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2415
         Left            =   120
         ScaleHeight     =   2415
         ScaleWidth      =   6015
         TabIndex        =   49
         Top             =   240
         Width           =   6015
         Begin VB.VScrollBar scrlObject 
            Height          =   2415
            LargeChange     =   30
            Left            =   5760
            Max             =   510
            Min             =   1
            SmallChange     =   15
            TabIndex        =   51
            Top             =   0
            Value           =   1
            Width           =   255
         End
         Begin VB.PictureBox picObjectSelect 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   18720
            Left            =   0
            Picture         =   "frmMapEdit.frx":83E764
            ScaleHeight     =   1248
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   384
            TabIndex        =   50
            Top             =   0
            Width           =   5760
         End
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Preview:"
         Height          =   255
         Left            =   6240
         TabIndex        =   53
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame frmEvents 
      Caption         =   "Events"
      Height          =   2775
      Left            =   2520
      TabIndex        =   14
      Top             =   2040
      Visible         =   0   'False
      Width           =   7935
      Begin VB.TextBox txtWarpMap 
         Height          =   285
         Left            =   5520
         TabIndex        =   34
         Text            =   "1"
         Top             =   1440
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtStoreNum 
         Height          =   285
         Left            =   5760
         TabIndex        =   33
         Text            =   "1"
         Top             =   1440
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtX 
         Height          =   285
         Left            =   5760
         TabIndex        =   24
         Text            =   "0"
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox txtY 
         Height          =   285
         Left            =   6840
         TabIndex        =   23
         Text            =   "0"
         Top             =   240
         Width           =   615
      End
      Begin VB.ComboBox cmbType 
         Height          =   315
         ItemData        =   "frmMapEdit.frx":99D7A6
         Left            =   5520
         List            =   "frmMapEdit.frx":99D7B6
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   600
         Width           =   2295
      End
      Begin VB.CommandButton cmbWarpMap 
         Caption         =   "M"
         Height          =   255
         Left            =   7560
         TabIndex        =   21
         Top             =   1800
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox txtWarpY 
         Height          =   285
         Left            =   6840
         TabIndex        =   20
         Text            =   "0"
         Top             =   1800
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtWarpX 
         Height          =   285
         Left            =   5760
         TabIndex        =   19
         Text            =   "0"
         Top             =   1800
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdSaveEvent 
         Caption         =   "<  Add Event"
         Height          =   375
         Left            =   5520
         TabIndex        =   18
         Top             =   2280
         Width           =   1455
      End
      Begin VB.ComboBox cmbTrigger 
         Height          =   315
         ItemData        =   "frmMapEdit.frx":99D7D3
         Left            =   5520
         List            =   "frmMapEdit.frx":99D7E0
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   960
         Width           =   2295
      End
      Begin VB.TextBox txtSign 
         Height          =   285
         Left            =   5520
         TabIndex        =   16
         Text            =   "This sign is Blank"
         Top             =   1440
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.TextBox txtEvent 
         Height          =   2415
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   15
         Top             =   240
         Width           =   3855
      End
      Begin VB.Label lblwarp1 
         Alignment       =   1  'Right Justify
         Caption         =   "Event Location:"
         Height          =   255
         Index           =   4
         Left            =   3840
         TabIndex        =   45
         Top             =   240
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label lblwarp1 
         Alignment       =   1  'Right Justify
         Caption         =   "Warp to map #:"
         Height          =   255
         Index           =   0
         Left            =   4080
         TabIndex        =   36
         Top             =   1440
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lblStore 
         Alignment       =   1  'Right Justify
         Caption         =   "Store #:"
         Height          =   255
         Left            =   3840
         TabIndex        =   35
         Top             =   1440
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "X"
         Height          =   255
         Left            =   5520
         TabIndex        =   32
         Top             =   240
         Width           =   135
      End
      Begin VB.Label Label3 
         Caption         =   "Y"
         Height          =   255
         Left            =   6600
         TabIndex        =   31
         Top             =   240
         Width           =   135
      End
      Begin VB.Label lblwarp1 
         Caption         =   "Y"
         Height          =   255
         Index           =   3
         Left            =   6600
         TabIndex        =   30
         Top             =   1800
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblwarp1 
         Caption         =   "X"
         Height          =   255
         Index           =   2
         Left            =   5520
         TabIndex        =   29
         Top             =   1800
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblwarp1 
         Alignment       =   1  'Right Justify
         Caption         =   "Location:"
         Height          =   255
         Index           =   1
         Left            =   3840
         TabIndex        =   28
         Top             =   1800
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Triggered By:"
         Height          =   255
         Left            =   3840
         TabIndex        =   27
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Type:"
         Height          =   255
         Left            =   3840
         TabIndex        =   26
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label lblSign 
         Alignment       =   1  'Right Justify
         Caption         =   "Sign Text:"
         Height          =   255
         Left            =   3840
         TabIndex        =   25
         Top             =   1440
         Visible         =   0   'False
         Width           =   1455
      End
   End
   Begin VB.Frame frmTiles 
      Caption         =   "Tiles"
      Height          =   2775
      Left            =   120
      TabIndex        =   4
      Top             =   6120
      Width           =   7935
      Begin VB.PictureBox picSelect 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   7320
         ScaleHeight     =   31
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   31
         TabIndex        =   12
         Top             =   240
         Width           =   495
      End
      Begin VB.OptionButton optWalk 
         Caption         =   "Walkable"
         Height          =   255
         Index           =   0
         Left            =   6240
         TabIndex        =   11
         Top             =   840
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton optWalk 
         Caption         =   "Blocked"
         Height          =   255
         Index           =   1
         Left            =   6240
         TabIndex        =   10
         Top             =   1080
         Width           =   1215
      End
      Begin VB.OptionButton optWalk 
         Caption         =   "Water"
         Height          =   255
         Index           =   2
         Left            =   6240
         TabIndex        =   9
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CommandButton cmdFill 
         Caption         =   "Fill"
         Height          =   375
         Left            =   6240
         TabIndex        =   8
         Top             =   2160
         Width           =   1575
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2415
         Left            =   120
         ScaleHeight     =   2415
         ScaleWidth      =   6015
         TabIndex        =   5
         Top             =   240
         Width           =   6015
         Begin VB.PictureBox picBackSelect 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   18720
            Left            =   0
            Picture         =   "frmMapEdit.frx":99D809
            ScaleHeight     =   1248
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   384
            TabIndex        =   7
            Top             =   0
            Width           =   5760
         End
         Begin VB.VScrollBar scrlPicture 
            Height          =   2415
            LargeChange     =   30
            Left            =   5760
            Max             =   510
            Min             =   1
            SmallChange     =   15
            TabIndex        =   6
            Top             =   0
            Value           =   1
            Width           =   255
         End
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Preview:"
         Height          =   255
         Left            =   6240
         TabIndex        =   13
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdEditItems 
      Caption         =   "Store/Item Editor"
      Height          =   375
      Left            =   9960
      TabIndex        =   46
      Top             =   7440
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Map Data"
      Height          =   1335
      Left            =   8280
      TabIndex        =   38
      Top             =   6000
      Width           =   1575
      Begin VB.TextBox txtMapLoaded 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   39
         Text            =   "0"
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Map Y:"
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Map X:"
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Map Loaded:"
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblX 
         Caption         =   "0"
         Height          =   255
         Left            =   1200
         TabIndex        =   41
         Top             =   600
         Width           =   255
      End
      Begin VB.Label lblY 
         Caption         =   "0"
         Height          =   255
         Left            =   1200
         TabIndex        =   40
         Top             =   960
         Width           =   255
      End
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh Map"
      Height          =   375
      Left            =   9960
      TabIndex        =   37
      Top             =   7800
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save Map"
      Height          =   375
      Left            =   9960
      TabIndex        =   0
      Top             =   8160
      Width           =   1455
   End
   Begin VB.PictureBox picBuffer 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   5775
      Left            =   0
      ScaleHeight     =   383
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   767
      TabIndex        =   1
      Top             =   0
      Width           =   11535
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   3255
      Left            =   0
      TabIndex        =   47
      Top             =   5760
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   5741
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Tiles"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Events"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Objects"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Npcs"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label lblNpcY 
      Caption         =   "0"
      Height          =   255
      Left            =   11760
      TabIndex        =   83
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label lblNpcX 
      Caption         =   "0"
      Height          =   255
      Left            =   11760
      TabIndex        =   82
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label11 
      Caption         =   "Show:"
      Height          =   255
      Left            =   8400
      TabIndex        =   68
      Top             =   7440
      Width           =   855
   End
   Begin VB.Label lblObjectX 
      Caption         =   "0"
      Height          =   255
      Left            =   11760
      TabIndex        =   60
      Top             =   960
      Width           =   855
   End
   Begin VB.Label lblObjectY 
      Caption         =   "0"
      Height          =   255
      Left            =   11760
      TabIndex        =   59
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label lblTileY 
      Caption         =   "0"
      Height          =   255
      Left            =   11760
      TabIndex        =   3
      Top             =   600
      Width           =   855
   End
   Begin VB.Label lblTileX 
      Caption         =   "0"
      Height          =   255
      Left            =   11760
      TabIndex        =   2
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "frmEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, _
ByVal X As Long, _
ByVal Y As Long, _
ByVal nWidth As Long, _
ByVal nHeight As Long, _
ByVal hSrcDC As Long, _
ByVal xSrc As Long, _
ByVal ySrc As Long, _
ByVal nSrcWidth As Long, _
ByVal nSrcHeight As Long, _
ByVal dwRop As Long) As Long

Public isSaved As Boolean

Private Sub chkGrid_Click()
refreshmap
End Sub

Private Sub chkNpc_Click()
refreshmap
End Sub

Private Sub chkObjover_Click()
refreshmap
End Sub

Private Sub chkobjunder_Click()
refreshmap
End Sub

Private Sub chkText_Click()
refreshmap
End Sub

Private Sub cmbWarpMap_Click()
Load frmAddMap
frmAddMap.Show
End Sub

Private Sub cmdAddMap_Click()
On Error Resume Next
MapNum = InputBox("Create map number?", "New Map", 0)

If MapNum = 0 Then

Else
  For l = 1 To 9
    For Y = 0 To MapY
      For X = 0 To MapX
     
        mapsrt = mapsrt & "0,"
      
      Next X
    Next Y
  Next l
  
Open App.Path & "\maps\map" & MapNum & ".dat" For Output As #1
    Print #1, mapsrt
Close #1

Open App.Path & "\events\event" & MapNum & ".dat" For Output As #1
Print #1, ""
Close #1

fileMaps.Refresh
End If
End Sub

Private Sub cmdEditItems_Click()
Load frmItem
frmItem.Show
End Sub

Private Sub cmdEditStore_Click()
Load frmStore
frmStore.Show
End Sub

Private Sub cmdFill_Click()

msg = MsgBox("Are you sure you want to fill the map?", vbOKCancel)

If msg = vbOK Then
    TileX = lblTileX.Caption
    TileY = lblTileY.Caption

If optWalk(1) = True Then
walkable = 1
ElseIf optWalk(2) = True Then
walkable = 2
Else
walkable = 0
End If

mytileset = 0


    For Y = 0 To MapY
      For X = 0 To MapX

     With LocalMap.Tile(X, Y)
       .TileX = TileX
       .TileY = TileY
       .Walk = walkable
     End With


      Next X
    Next Y

     refreshmap
     isSaved = False
     
End If

End Sub

Sub loadMap(MapNum As Integer)
On Error GoTo stopit

restart:

Open App.Path & "\maps\map" & MapNum & ".dat" For Input As #1

  For l = 1 To 9
    For Y = 0 To MapY
      For X = 0 To MapX

        Input #1, mapsrt
        
     With LocalMap.Tile(X, Y)
     
     If l = 1 Then
        .TileX = mapsrt
     ElseIf l = 2 Then
        .TileY = mapsrt
     ElseIf l = 3 Then
        .Walk = mapsrt
     ElseIf l = 4 Then
        .objectX = mapsrt
     ElseIf l = 5 Then
        .objectY = mapsrt
     ElseIf l = 6 Then
        .objectoverX = mapsrt
     ElseIf l = 7 Then
        .objectoverY = mapsrt
     ElseIf l = 8 Then
        .npcX = mapsrt
     ElseIf l = 9 Then
        .npcY = mapsrt
     End If
     
     End With

      Next X
    Next Y
  Next l
  
Close #1

  picBackSelect.Picture = picBackSelectB.Picture
  
  StretchBlt picSelect.hdc, 0, 0, 32, 32, picBackSelectB.hdc, 0, 0, 16, 16, &HCC0020

  BitBlt picBackSelect.hdc, TileX * 16, TileY * 16, 16, 16, picBoxMask.hdc, 0, 0, vbMergePaint
  BitBlt picBackSelect.hdc, TileX * 16, TileY * 16, 16, 16, picBox.hdc, 0, 0, vbSrcAnd

  picObjectSelect.Picture = picObjectSelectB.Picture
  
  StretchBlt picObject.hdc, 0, 0, 32, 32, picObjectSelectB.hdc, 0, 0, 16, 16, &HCC0020

  BitBlt picObjectSelect.hdc, TileX * 16, TileY * 16, 16, 16, picBoxMask.hdc, 0, 0, vbMergePaint
  BitBlt picObjectSelect.hdc, TileX * 16, TileY * 16, 16, 16, picBox.hdc, 0, 0, vbSrcAnd

  picNpcSelect.Picture = picNpcSelectB.Picture
  
  StretchBlt picNpc.hdc, 0, 0, 32, 32, picNpcSelectB.hdc, 0, 0, 32, 32, &HCC0020

  BitBlt picNpcSelect.hdc, TileX * 32, TileY * 32, 32, 32, picBoxMask32.hdc, 0, 0, vbMergePaint
  BitBlt picNpcSelect.hdc, TileX * 32, TileY * 32, 32, 32, picBox32.hdc, 0, 0, vbSrcAnd

  
    lblTileX.Caption = 0
    lblTileY.Caption = 0
lblObjectX.Caption = 0
lblObjectY.Caption = 0
lblNpcX.Caption = 0
lblNpcY.Caption = 0

txtMapLoaded.Text = MapNum

refreshmap

On Error GoTo stoptrying
Dim eventdata As String
txtEvent.Text = ""


Open App.Path & "\events\event" & MapNum & ".dat" For Input As #1
Do Until (EOF(1))
Input #1, eventdata
If eventdata <> "" Then
txtEvent.Text = txtEvent.Text & eventdata & vbCrLf
End If
Loop
Close #1

isSaved = True

Exit Sub

stoptrying:
Exit Sub

stopit:

For i = 1 To 1152
mapline = mapline & "0,"
Next i
Open App.Path & "\maps\map" & MapNum & ".dat" For Output As #1
Print #1, mapline
Close #1
GoTo restart
End Sub

Private Sub cmdRefresh_Click()
refreshmap
End Sub

Private Sub cmdSave_Click()

save = MsgBox("Save changes to map " & txtMapLoaded.Text & "?", vbOKCancel, "Save Changes")

If save = vbOK Then
MapNum = txtMapLoaded.Text 'InputBox("Save as map number?", "Save Map")

  For l = 1 To 9
    For Y = 0 To MapY
      For X = 0 To MapX

     With LocalMap.Tile(X, Y)
     
     If l = 1 Then
        mapsrt = mapsrt & .TileX & ","
     ElseIf l = 2 Then
        mapsrt = mapsrt & .TileY & ","
     ElseIf l = 3 Then
        mapsrt = mapsrt & .Walk & ","
     ElseIf l = 4 Then
        mapsrt = mapsrt & .objectX & ","
     ElseIf l = 5 Then
        mapsrt = mapsrt & .objectY & ","
     ElseIf l = 6 Then
        mapsrt = mapsrt & .objectoverX & ","
     ElseIf l = 7 Then
        mapsrt = mapsrt & .objectoverY & ","
     ElseIf l = 8 Then
        mapsrt = mapsrt & .npcX & ","
     ElseIf l = 9 Then
        mapsrt = mapsrt & .npcY & ","
     End If
     
     End With
      
      Next X
    Next Y
  Next l
  
Open App.Path & "\maps\map" & MapNum & ".dat" For Output As #1
    Print #1, mapsrt
Close #1

Open App.Path & "\events\event" & MapNum & ".dat" For Output As #1
Print #1, txtEvent.Text
Close #1

msg = MsgBox("Map saved as number " & MapNum & "!", vbOKOnly, "Saved")

refreshmap

isSaved = True

End If

End Sub


Private Sub cmdSaveEvent_Click()
Dim myLine As String
MapNum = txtMapLoaded.Text

If cmbType.ListIndex = 0 Then
myLine = txtX.Text & pChar & txtY.Text & pChar & "warp" & pChar & cmbTrigger.ListIndex & pChar & txtWarpMap.Text & pChar & txtWarpX.Text & pChar & txtWarpY.Text & pChar & "0"
ElseIf cmbType.ListIndex = 1 Then
myLine = txtX.Text & pChar & txtY.Text & pChar & "store" & pChar & cmbTrigger.ListIndex & pChar & "0" & pChar & "0" & pChar & "0" & pChar & txtStoreNum.Text
ElseIf cmbType.ListIndex = 2 Then
myLine = txtX.Text & pChar & txtY.Text & pChar & "bank" & pChar & cmbTrigger.ListIndex & pChar & "0" & pChar & "0" & pChar & "0" & pChar & "0"
ElseIf cmbType.ListIndex = 3 Then
myLine = txtX.Text & pChar & txtY.Text & pChar & "sign" & pChar & cmbTrigger.ListIndex & pChar & "0" & pChar & "0" & pChar & "0" & pChar & txtSign.Text

End If

txtEvent.Text = txtEvent.Text & myLine & vbCrLf

Open App.Path & "\events\event" & MapNum & ".dat" For Output As #1
Print #1, txtEvent.Text
Close #1

refreshmap

isSaved = False

End Sub

Private Sub fileMaps_Click()

If isSaved = False Then
save = MsgBox("Save changes to current map?", vbOKCancel, "Save Changes")

If save = vbOK Then
cmdSave_Click
Exit Sub
End If
End If

mymapnum = Left(fileMaps.FileName, Len(fileMaps.FileName) - 4)
mymapnum = Right(mymapnum, Len(mymapnum) - 3)

loadMap (mymapnum)
End Sub

Private Sub Form_Load()

fileMaps.Path = App.Path & "\maps"

    For Y = 0 To MapY
      For X = 0 To MapX

     With LocalMap.Tile(X, Y)
       .TileX = 0
       .TileY = 0
       .Walk = 1
       .objectoverX = 0
       .objectoverY = 0
       .objectX = 0
       .objectY = 0
       .npcX = 0
       .npcY = 0
     End With

      Next X
    Next Y

Me.Show

Call SetParce
Call SetTextColor(picBuffer.hdc, vbWhite)

cmbTrigger.ListIndex = 1
cmbType.ListIndex = 0

loadMap (1)

End Sub


Private Sub picBackSelect_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 1 Then
    TileX = Int(X / 16)
    TileY = Int(Y / 16)
    lblTileX.Caption = TileX
    lblTileY.Caption = TileY

  StretchBlt picSelect.hdc, 0, 0, 32, 32, picBackSelectB.hdc, TileX * 16, TileY * 16, 16, 16, &HCC0020
  
  picBackSelect.Picture = picBackSelectB.Picture
  
  BitBlt picBackSelect.hdc, TileX * 16, TileY * 16, 16, 16, picBoxMask.hdc, 0, 0, vbMergePaint
  BitBlt picBackSelect.hdc, TileX * 16, TileY * 16, 16, 16, picBox.hdc, 0, 0, vbSrcAnd

  End If
End Sub

Private Sub picBuffer_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim x1, y1 As Long

If frmTiles.Visible = True Then

  x1 = Int(X / PicX)
  y1 = Int(Y / PicY)

    TileX = lblTileX.Caption
    TileY = lblTileY.Caption

If optWalk(1) = True Then
walkable = 1
ElseIf optWalk(2) = True Then
walkable = 2
Else
walkable = 0
End If

mytileset = 0

   If (Button = 1) And Shift = 2 And (x1 >= 0) And (x1 <= MapX) And (y1 >= 0) And (y1 <= MapY) Then

     With LocalMap.Tile(x1, y1)
       .TileX = TileX
       .TileY = TileY
       .Walk = walkable
     End With
     refreshmap
     isSaved = False
   End If
   If (Button = 2) And Shift = 2 And (x1 >= 0) And (x1 <= MapX) And (y1 >= 0) And (y1 <= MapY) Then
     With LocalMap.Tile(x1, y1)
       .TileX = 0
       .TileY = 0
       .Walk = 0
     End With
     refreshmap
     isSaved = False
   End If



ElseIf frmObjects.Visible = True Then
  
  x1 = Int(X / PicX)
  y1 = Int(Y / PicY)

    TileX = lblObjectX.Caption
    TileY = lblObjectY.Caption



mytileset = 0

   If (Button = 1) And Shift = 2 And (x1 >= 0) And (x1 <= MapX) And (y1 >= 0) And (y1 <= MapY) Then

     With LocalMap.Tile(x1, y1)
If optOver(1) = True Then
      .objectX = TileX
      .objectY = TileY
Else
      .objectoverX = TileX
      .objectoverY = TileY
End If
     
     End With
     refreshmap
     isSaved = False
   End If
   
   If (Button = 2) And Shift = 2 And (x1 >= 0) And (x1 <= MapX) And (y1 >= 0) And (y1 <= MapY) Then
     With LocalMap.Tile(x1, y1)
      .objectX = 0
      .objectY = 0
      .objectoverX = 0
      .objectoverY = 0
     End With
     refreshmap
     isSaved = False
   End If



ElseIf frmNpc.Visible = True Then
  
  x1 = Int(X / PicX)
  y1 = Int(Y / PicY)

    TileX = lblNpcX.Caption
    TileY = lblNpcY.Caption



mytileset = 0

   If (Button = 1) And Shift = 2 And (x1 >= 0) And (x1 <= MapX) And (y1 >= 0) And (y1 <= MapY) Then

     With LocalMap.Tile(x1, y1)
        .npcX = TileX
        .npcY = TileY
     End With
     refreshmap
     isSaved = False
   End If
   
   If (Button = 2) And Shift = 2 And (x1 >= 0) And (x1 <= MapX) And (y1 >= 0) And (y1 <= MapY) Then
     With LocalMap.Tile(x1, y1)
        .npcX = 0
        .npcY = 0
     End With
     refreshmap
     isSaved = False
   End If
   
ElseIf frmEvents.Visible = True Then

  x1 = Int(X / PicX)
  y1 = Int(Y / PicY)

    txtX = lblX.Caption
    txtY = lblY.Caption

End If

End Sub

Private Sub refreshmap()

Dim X, Y As Long

    For Y = 0 To MapY
      For X = 0 To MapX
        With LocalMap.Tile(X, Y)
                        
StretchBlt picBuffer.hdc, X * PicX, Y * PicY, 32, 32, picBackSelectB.hdc, .TileX * 16, .TileY * 16, 16, 16, &HCC0020

If chkobjunder.Value = 1 Then
  If .objectX = 0 And .objectY = 0 Then
  'donuthing
  Else
StretchBlt picBuffer.hdc, X * PicX, Y * PicY, 32, 32, picObjectSelectMask.hdc, .objectX * 16, .objectY * 16, 16, 16, vbMergePaint
StretchBlt picBuffer.hdc, X * PicX, Y * PicY, 32, 32, picObjectSelectB.hdc, .objectX * 16, .objectY * 16, 16, 16, vbSrcAnd
           If chkText.Value = 1 Then
           Call SetTextColor(picBuffer.hdc, vbYellow)
           Call TextOut(picBuffer.hdc, X * PicX + 1, Y * PicY + 19, "U", 1)
           End If
   End If
End If

If chkNpc.Value = 1 Then
If .npcX = 0 And .npcY = 0 Then
  'do nuthing
  Else
  StretchBlt picBuffer.hdc, X * PicX, Y * PicY, 32, 32, picNpcSelectMask.hdc, .npcX * 32, .npcY * 32, 32, 32, vbMergePaint
  StretchBlt picBuffer.hdc, X * PicX, Y * PicY, 32, 32, picNpcSelectB.hdc, .npcX * 32, .npcY * 32, 32, 32, vbSrcAnd
           If chkText.Value = 1 Then
           Call SetTextColor(picBuffer.hdc, &HFFFF00)
           Call TextOut(picBuffer.hdc, X * PicX + 9, Y * PicY + 19, "NPC", 3)
           End If
End If
End If

If chkObjover.Value = 1 Then
If .objectoverX = 0 And .objectoverY = 0 Then
  'do nuthing
  Else
  StretchBlt picBuffer.hdc, X * PicX, Y * PicY, 32, 32, picObjectSelectMask.hdc, .objectoverX * 16, .objectoverY * 16, 16, 16, vbMergePaint
  StretchBlt picBuffer.hdc, X * PicX, Y * PicY, 32, 32, picObjectSelectB.hdc, .objectoverX * 16, .objectoverY * 16, 16, 16, vbSrcAnd
           If chkText.Value = 1 Then
           Call SetTextColor(picBuffer.hdc, vbYellow)
           Call TextOut(picBuffer.hdc, X * PicX + 1, Y * PicY + 19, "O", 1)
           End If
   End If
End If


If chkGrid.Value = 1 Then
BitBlt picBuffer.hdc, X * PicX, Y * PicY, 32, 32, picGridMask.hdc, 0, 0, vbMergePaint
BitBlt picBuffer.hdc, X * PicX, Y * PicY, 32, 32, picGrid.hdc, 0, 0, vbSrcAnd
End If
  
  If chkText.Value = 1 Then
           If .Walk = 1 Then
           Call SetTextColor(picBuffer.hdc, vbWhite)
             Call TextOut(picBuffer.hdc, X * PicX + 1, Y * PicY + 8, "Block", 5)
           End If
           
           If .Walk = 2 Then
           Call SetTextColor(picBuffer.hdc, vbWhite)
             Call TextOut(picBuffer.hdc, X * PicX + 1, Y * PicY + 8, "Water", 5)
           End If
  End If
           
        End With
      Next X
    Next Y



Open App.Path & "\events\event" & txtMapLoaded.Text & ".dat" For Input As #1
Do Until (EOF(1))

Line Input #1, eventdata

If eventdata <> "" Then
myevent = Split(eventdata, pChar)

If myevent(2) = "warp" Then
leter = "Warp "
ElseIf myevent(2) = "store" Then
leter = "Store"
ElseIf myevent(2) = "bank" Then
leter = "Bank "
ElseIf myevent(2) = "sign" Then
leter = "Sign "
End If

If chkText.Value = 1 Then
Call SetTextColor(picBuffer.hdc, vbGreen)
Call TextOut(picBuffer.hdc, myevent(0) * PicX + 1, myevent(1) * PicY - 2, leter, 5)
End If

End If
Loop

Close #1

End Sub

Private Sub picBuffer_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim x1, y1 As Long

  x1 = Int(X / PicX)
  y1 = Int(Y / PicY)

lblX.Caption = x1
lblY.Caption = y1

If frmTiles.Visible = True Then

Call picBuffer_MouseDown(Button, Shift, X, Y)

End If
End Sub

Private Sub picObjectSelect_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  If Button = 1 Then
    TileX = Int(X / 16)
    TileY = Int(Y / 16)
    lblObjectX.Caption = TileX
    lblObjectY.Caption = TileY

  StretchBlt picObject.hdc, 0, 0, 32, 32, picObjectSelectB.hdc, TileX * 16, TileY * 16, 16, 16, &HCC0020
  
  picObjectSelect.Picture = picObjectSelectB.Picture
  
  BitBlt picObjectSelect.hdc, TileX * 16, TileY * 16, 16, 16, picBoxMask.hdc, 0, 0, vbMergePaint
  BitBlt picObjectSelect.hdc, TileX * 16, TileY * 16, 16, 16, picBox.hdc, 0, 0, vbSrcAnd

  End If
End Sub

Private Sub picNpcSelect_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  If Button = 1 Then
    TileX = Int(X / 32)
    TileY = Int(Y / 32)
    lblNpcX.Caption = TileX
    lblNpcY.Caption = TileY

  StretchBlt picNpc.hdc, 0, 0, 32, 32, picNpcSelectB.hdc, TileX * 32, TileY * 32, 32, 32, &HCC0020
  
  picNpcSelect.Picture = picNpcSelectB.Picture
  
  BitBlt picNpcSelect.hdc, TileX * 32, TileY * 32, 32, 32, picBoxMask32.hdc, 0, 0, vbMergePaint
  BitBlt picNpcSelect.hdc, TileX * 32, TileY * 32, 32, 32, picBox32.hdc, 0, 0, vbSrcAnd

  End If
End Sub


Private Sub scrlObject_Change()
On Error Resume Next

picObjectSelect.Top = (scrlObject.Value * 32) * -1
End Sub

Private Sub scrlObject_Scroll()
Call scrlObject_Change
End Sub

Private Sub scrlPicture_Change()
On Error Resume Next

picBackSelect.Top = (scrlPicture.Value * 32) * -1
End Sub

Private Sub scrlPicture_Scroll()
Call scrlPicture_Change
End Sub

Private Sub Form_Unload(Cancel As Integer)

If isSaved = False Then
save = MsgBox("Save changes to current map?", vbOKCancel, "Save Changes")

If save = vbOK Then
cmdSave_Click
Exit Sub
End If
End If

End
End Sub

Private Sub cmbType_Click()

lblwarp1(0).Visible = False
lblwarp1(1).Visible = False
lblwarp1(2).Visible = False
lblwarp1(3).Visible = False
txtWarpMap.Visible = False
txtWarpX.Visible = False
txtWarpY.Visible = False
cmbWarpMap.Visible = False
lblStore.Visible = False
txtStoreNum.Visible = False
lblSign.Visible = False
txtSign.Visible = False

If cmbType.ListIndex = 0 Then
lblwarp1(0).Visible = True
lblwarp1(1).Visible = True
lblwarp1(2).Visible = True
lblwarp1(3).Visible = True
txtWarpMap.Visible = True
txtWarpX.Visible = True
txtWarpY.Visible = True
cmbWarpMap.Visible = True

ElseIf cmbType.ListIndex = 1 Then
lblStore.Visible = True
txtStoreNum.Visible = True

ElseIf cmbType.ListIndex = 3 Then
lblSign.Visible = True
txtSign.Visible = True
End If

End Sub

Private Sub TabStrip1_Click()

lblTileX.Caption = 0
lblTileY.Caption = 0
lblObjectX.Caption = 0
lblObjectY.Caption = 0
lblNpcX.Caption = 0
lblNpcY.Caption = 0

  picBackSelect.Picture = picBackSelectB.Picture
  
  StretchBlt picSelect.hdc, 0, 0, 32, 32, picBackSelectB.hdc, 0, 0, 16, 16, &HCC0020

  BitBlt picBackSelect.hdc, TileX * 16, TileY * 16, 16, 16, picBoxMask.hdc, 0, 0, vbMergePaint
  BitBlt picBackSelect.hdc, TileX * 16, TileY * 16, 16, 16, picBox.hdc, 0, 0, vbSrcAnd

  picObjectSelect.Picture = picObjectSelectB.Picture
  
  StretchBlt picObject.hdc, 0, 0, 32, 32, picObjectSelectB.hdc, 0, 0, 16, 16, &HCC0020

  BitBlt picObjectSelect.hdc, TileX * 16, TileY * 16, 16, 16, picBoxMask.hdc, 0, 0, vbMergePaint
  BitBlt picObjectSelect.hdc, TileX * 16, TileY * 16, 16, 16, picBox.hdc, 0, 0, vbSrcAnd

  picNpcSelect.Picture = picNpcSelectB.Picture
  
  StretchBlt picNpc.hdc, 0, 0, 32, 32, picNpcSelectB.hdc, 0, 0, 32, 32, &HCC0020

  BitBlt picNpcSelect.hdc, TileX * 32, TileY * 32, 32, 32, picBoxMask32.hdc, 0, 0, vbMergePaint
  BitBlt picNpcSelect.hdc, TileX * 32, TileY * 32, 32, 32, picBox32.hdc, 0, 0, vbSrcAnd

If TabStrip1.SelectedItem.Index = 1 Then
frmEvents.Visible = False
frmTiles.Visible = True
frmObjects.Visible = False
frmNpc.Visible = False
ElseIf TabStrip1.SelectedItem.Index = 2 Then
frmEvents.Visible = True
frmTiles.Visible = False
frmObjects.Visible = False
frmNpc.Visible = False
ElseIf TabStrip1.SelectedItem.Index = 3 Then
frmEvents.Visible = False
frmTiles.Visible = False
frmObjects.Visible = True
frmNpc.Visible = False
ElseIf TabStrip1.SelectedItem.Index = 4 Then
frmEvents.Visible = False
frmTiles.Visible = False
frmObjects.Visible = False
frmNpc.Visible = True
End If

End Sub
