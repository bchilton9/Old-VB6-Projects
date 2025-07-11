VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmItem 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Editor"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4575
   Icon            =   "frmItemEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   4575
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picItem 
      Height          =   495
      Index           =   0
      Left            =   4920
      Picture         =   "frmItemEdit.frx":0E42
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   114
      Top             =   0
      Width           =   495
   End
   Begin VB.Frame frmStore 
      Caption         =   "Store's"
      Height          =   5655
      Left            =   120
      TabIndex        =   31
      Top             =   600
      Visible         =   0   'False
      Width           =   4335
      Begin VB.TextBox txtPrice 
         DataField       =   "Price15"
         DataSource      =   "axsStore"
         Height          =   285
         Index           =   15
         Left            =   3720
         TabIndex        =   90
         Top             =   3480
         Width           =   495
      End
      Begin VB.TextBox txtPrice 
         DataField       =   "Price14"
         DataSource      =   "axsStore"
         Height          =   285
         Index           =   14
         Left            =   3720
         TabIndex        =   89
         Top             =   3120
         Width           =   495
      End
      Begin VB.TextBox txtPrice 
         DataField       =   "Price13"
         DataSource      =   "axsStore"
         Height          =   285
         Index           =   13
         Left            =   3720
         TabIndex        =   88
         Top             =   2760
         Width           =   495
      End
      Begin VB.TextBox txtPrice 
         DataField       =   "Price12"
         DataSource      =   "axsStore"
         Height          =   285
         Index           =   12
         Left            =   3720
         TabIndex        =   87
         Top             =   2400
         Width           =   495
      End
      Begin VB.TextBox txtPrice 
         DataField       =   "Price11"
         DataSource      =   "axsStore"
         Height          =   285
         Index           =   11
         Left            =   3720
         TabIndex        =   86
         Top             =   2040
         Width           =   495
      End
      Begin VB.TextBox txtPrice 
         DataField       =   "Price10"
         DataSource      =   "axsStore"
         Height          =   285
         Index           =   10
         Left            =   3720
         TabIndex        =   85
         Top             =   1680
         Width           =   495
      End
      Begin VB.TextBox txtPrice 
         DataField       =   "Price9"
         DataSource      =   "axsStore"
         Height          =   285
         Index           =   9
         Left            =   3720
         TabIndex        =   84
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox txtPrice 
         DataField       =   "Price8"
         DataSource      =   "axsStore"
         Height          =   285
         Index           =   8
         Left            =   1560
         TabIndex        =   83
         Top             =   3840
         Width           =   495
      End
      Begin VB.TextBox txtPrice 
         DataField       =   "Price7"
         DataSource      =   "axsStore"
         Height          =   285
         Index           =   7
         Left            =   1560
         TabIndex        =   82
         Top             =   3480
         Width           =   495
      End
      Begin VB.TextBox txtPrice 
         DataField       =   "Price6"
         DataSource      =   "axsStore"
         Height          =   285
         Index           =   6
         Left            =   1560
         TabIndex        =   81
         Top             =   3120
         Width           =   495
      End
      Begin VB.TextBox txtPrice 
         DataField       =   "Price5"
         DataSource      =   "axsStore"
         Height          =   285
         Index           =   5
         Left            =   1560
         TabIndex        =   80
         Top             =   2760
         Width           =   495
      End
      Begin VB.TextBox txtPrice 
         DataField       =   "Price4"
         DataSource      =   "axsStore"
         Height          =   285
         Index           =   4
         Left            =   1560
         TabIndex        =   79
         Top             =   2400
         Width           =   495
      End
      Begin VB.TextBox txtPrice 
         DataField       =   "Price3"
         DataSource      =   "axsStore"
         Height          =   285
         Index           =   3
         Left            =   1560
         TabIndex        =   78
         Top             =   2040
         Width           =   495
      End
      Begin VB.TextBox txtPrice 
         DataField       =   "Price2"
         DataSource      =   "axsStore"
         Height          =   285
         Index           =   2
         Left            =   1560
         TabIndex        =   77
         Top             =   1680
         Width           =   495
      End
      Begin VB.TextBox txtPrice 
         DataField       =   "Price1"
         DataSource      =   "axsStore"
         Height          =   285
         Index           =   1
         Left            =   1560
         TabIndex        =   76
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         DataField       =   "StoresID"
         DataSource      =   "axsStore"
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   73
         Top             =   720
         Width           =   615
      End
      Begin VB.CommandButton cmdDeleteStore 
         Caption         =   "Delete Store"
         Height          =   375
         Left            =   2880
         TabIndex        =   71
         Top             =   4800
         Width           =   1215
      End
      Begin VB.Data axsStore 
         Caption         =   "Stores"
         Connect         =   "Access 2000;"
         DatabaseName    =   "events\Item.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Stores"
         Top             =   240
         Width           =   4095
      End
      Begin VB.TextBox Store 
         DataField       =   "Store1"
         DataSource      =   "axsStore"
         Height          =   285
         Index           =   1
         Left            =   1080
         TabIndex        =   64
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox Store 
         DataField       =   "Store2"
         DataSource      =   "axsStore"
         Height          =   285
         Index           =   2
         Left            =   1080
         TabIndex        =   63
         Top             =   1680
         Width           =   375
      End
      Begin VB.TextBox Store 
         DataField       =   "Store3"
         DataSource      =   "axsStore"
         Height          =   285
         Index           =   3
         Left            =   1080
         TabIndex        =   62
         Top             =   2040
         Width           =   375
      End
      Begin VB.TextBox Store 
         DataField       =   "Store4"
         DataSource      =   "axsStore"
         Height          =   285
         Index           =   4
         Left            =   1080
         TabIndex        =   61
         Top             =   2400
         Width           =   375
      End
      Begin VB.TextBox Store 
         DataField       =   "Store5"
         DataSource      =   "axsStore"
         Height          =   285
         Index           =   5
         Left            =   1080
         TabIndex        =   60
         Top             =   2760
         Width           =   375
      End
      Begin VB.TextBox Store 
         DataField       =   "Store6"
         DataSource      =   "axsStore"
         Height          =   285
         Index           =   6
         Left            =   1080
         TabIndex        =   59
         Top             =   3120
         Width           =   375
      End
      Begin VB.TextBox Store 
         DataField       =   "Store7"
         DataSource      =   "axsStore"
         Height          =   285
         Index           =   7
         Left            =   1080
         TabIndex        =   58
         Top             =   3480
         Width           =   375
      End
      Begin VB.TextBox Store 
         DataField       =   "Store8"
         DataSource      =   "axsStore"
         Height          =   285
         Index           =   8
         Left            =   1080
         TabIndex        =   57
         Top             =   3840
         Width           =   375
      End
      Begin VB.TextBox Store 
         DataField       =   "Store9"
         DataSource      =   "axsStore"
         Height          =   285
         Index           =   9
         Left            =   3240
         TabIndex        =   56
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox Store 
         DataField       =   "Store10"
         DataSource      =   "axsStore"
         Height          =   285
         Index           =   10
         Left            =   3240
         TabIndex        =   55
         Top             =   1680
         Width           =   375
      End
      Begin VB.TextBox Store 
         DataField       =   "Store11"
         DataSource      =   "axsStore"
         Height          =   285
         Index           =   11
         Left            =   3240
         TabIndex        =   54
         Top             =   2040
         Width           =   375
      End
      Begin VB.TextBox Store 
         DataField       =   "Store12"
         DataSource      =   "axsStore"
         Height          =   285
         Index           =   12
         Left            =   3240
         TabIndex        =   53
         Top             =   2400
         Width           =   375
      End
      Begin VB.TextBox Store 
         DataField       =   "Store13"
         DataSource      =   "axsStore"
         Height          =   285
         Index           =   13
         Left            =   3240
         TabIndex        =   52
         Top             =   2760
         Width           =   375
      End
      Begin VB.TextBox Store 
         DataField       =   "Store14"
         DataSource      =   "axsStore"
         Height          =   285
         Index           =   14
         Left            =   3240
         TabIndex        =   51
         Top             =   3120
         Width           =   375
      End
      Begin VB.TextBox Store 
         DataField       =   "Store15"
         DataSource      =   "axsStore"
         Height          =   285
         Index           =   15
         Left            =   3240
         TabIndex        =   50
         Top             =   3480
         Width           =   375
      End
      Begin VB.OptionButton optItem 
         Caption         =   "Item#1:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   49
         Top             =   1320
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton optItem 
         Caption         =   "Item#2:"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   48
         Top             =   1680
         Width           =   1215
      End
      Begin VB.OptionButton optItem 
         Caption         =   "Item#3:"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   47
         Top             =   2040
         Width           =   1215
      End
      Begin VB.OptionButton optItem 
         Caption         =   "Item#4:"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   46
         Top             =   2400
         Width           =   1215
      End
      Begin VB.OptionButton optItem 
         Caption         =   "Item#5:"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   45
         Top             =   2760
         Width           =   1215
      End
      Begin VB.OptionButton optItem 
         Caption         =   "Item#6:"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   44
         Top             =   3120
         Width           =   1215
      End
      Begin VB.OptionButton optItem 
         Caption         =   "Item#7:"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   43
         Top             =   3480
         Width           =   1215
      End
      Begin VB.OptionButton optItem 
         Caption         =   "Item#8:"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   42
         Top             =   3840
         Width           =   1215
      End
      Begin VB.OptionButton optItem 
         Caption         =   "Item#9:"
         Height          =   255
         Index           =   9
         Left            =   2280
         TabIndex        =   41
         Top             =   1320
         Width           =   1215
      End
      Begin VB.OptionButton optItem 
         Caption         =   "Item#10:"
         Height          =   255
         Index           =   10
         Left            =   2280
         TabIndex        =   40
         Top             =   1680
         Width           =   1215
      End
      Begin VB.OptionButton optItem 
         Caption         =   "Item#11:"
         Height          =   255
         Index           =   11
         Left            =   2280
         TabIndex        =   39
         Top             =   2040
         Width           =   1215
      End
      Begin VB.OptionButton optItem 
         Caption         =   "Item#12:"
         Height          =   255
         Index           =   12
         Left            =   2280
         TabIndex        =   38
         Top             =   2400
         Width           =   1215
      End
      Begin VB.OptionButton optItem 
         Caption         =   "Item#13:"
         Height          =   255
         Index           =   13
         Left            =   2280
         TabIndex        =   37
         Top             =   2760
         Width           =   1215
      End
      Begin VB.OptionButton optItem 
         Caption         =   "Item#14:"
         Height          =   255
         Index           =   14
         Left            =   2280
         TabIndex        =   36
         Top             =   3120
         Width           =   1215
      End
      Begin VB.OptionButton optItem 
         Caption         =   "Item#15:"
         Height          =   255
         Index           =   15
         Left            =   2280
         TabIndex        =   35
         Top             =   3480
         Width           =   1215
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "Remove Item From Store"
         Height          =   255
         Left            =   1200
         TabIndex        =   34
         Top             =   4200
         Width           =   1935
      End
      Begin VB.CommandButton cmdSaveStore 
         Caption         =   "Save Store"
         Height          =   375
         Left            =   240
         TabIndex        =   33
         Top             =   4800
         Width           =   1215
      End
      Begin VB.CommandButton cmdNewStore 
         Caption         =   "New Store"
         Height          =   375
         Left            =   1560
         TabIndex        =   32
         Top             =   4800
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "#:      Price:"
         Height          =   255
         Left            =   3360
         TabIndex        =   116
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "#:      Price:"
         Height          =   255
         Left            =   1200
         TabIndex        =   115
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "StoreID #:"
         Height          =   255
         Left            =   1200
         TabIndex        =   72
         Top             =   720
         Width           =   855
      End
   End
   Begin VB.PictureBox picItem 
      Height          =   495
      Index           =   1
      Left            =   4920
      Picture         =   "frmItemEdit.frx":36E86
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   113
      Top             =   600
      Width           =   495
   End
   Begin VB.PictureBox picItem 
      Height          =   495
      Index           =   2
      Left            =   5400
      Picture         =   "frmItemEdit.frx":37ACA
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   112
      Top             =   600
      Width           =   495
   End
   Begin VB.PictureBox picItem 
      Height          =   495
      Index           =   3
      Left            =   5880
      Picture         =   "frmItemEdit.frx":3870E
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   111
      Top             =   600
      Width           =   495
   End
   Begin VB.PictureBox picItem 
      Height          =   495
      Index           =   4
      Left            =   6360
      Picture         =   "frmItemEdit.frx":39352
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   110
      Top             =   600
      Width           =   495
   End
   Begin VB.PictureBox picItem 
      Height          =   495
      Index           =   5
      Left            =   6840
      Picture         =   "frmItemEdit.frx":39F96
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   109
      Top             =   600
      Width           =   495
   End
   Begin VB.PictureBox picItem 
      Height          =   495
      Index           =   6
      Left            =   7320
      Picture         =   "frmItemEdit.frx":3ABDA
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   108
      Top             =   600
      Width           =   495
   End
   Begin VB.PictureBox picItem 
      Height          =   495
      Index           =   7
      Left            =   7800
      Picture         =   "frmItemEdit.frx":3B81E
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   107
      Top             =   600
      Width           =   495
   End
   Begin VB.PictureBox picItem 
      Height          =   495
      Index           =   8
      Left            =   8280
      Picture         =   "frmItemEdit.frx":3C462
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   106
      Top             =   600
      Width           =   495
   End
   Begin VB.PictureBox picItem 
      Height          =   495
      Index           =   9
      Left            =   8760
      Picture         =   "frmItemEdit.frx":3D0A6
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   105
      Top             =   600
      Width           =   495
   End
   Begin VB.PictureBox picItem 
      Height          =   495
      Index           =   10
      Left            =   9240
      Picture         =   "frmItemEdit.frx":3DCEA
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   104
      Top             =   600
      Width           =   495
   End
   Begin VB.PictureBox picItem 
      Height          =   495
      Index           =   11
      Left            =   4920
      Picture         =   "frmItemEdit.frx":3E92E
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   103
      Top             =   1080
      Width           =   495
   End
   Begin VB.PictureBox picItem 
      Height          =   495
      Index           =   12
      Left            =   5400
      Picture         =   "frmItemEdit.frx":3F572
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   102
      Top             =   1080
      Width           =   495
   End
   Begin VB.PictureBox picItem 
      Height          =   495
      Index           =   13
      Left            =   5880
      Picture         =   "frmItemEdit.frx":401B6
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   101
      Top             =   1080
      Width           =   495
   End
   Begin VB.PictureBox picItem 
      Height          =   495
      Index           =   14
      Left            =   6360
      Picture         =   "frmItemEdit.frx":40DFA
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   100
      Top             =   1080
      Width           =   495
   End
   Begin VB.PictureBox picItem 
      Height          =   495
      Index           =   15
      Left            =   6840
      Picture         =   "frmItemEdit.frx":41A3E
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   99
      Top             =   1080
      Width           =   495
   End
   Begin VB.PictureBox picItem 
      Height          =   495
      Index           =   16
      Left            =   7320
      Picture         =   "frmItemEdit.frx":42682
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   98
      Top             =   1080
      Width           =   495
   End
   Begin VB.PictureBox picItem 
      Height          =   495
      Index           =   17
      Left            =   7800
      Picture         =   "frmItemEdit.frx":432C6
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   97
      Top             =   1080
      Width           =   495
   End
   Begin VB.PictureBox picItem 
      Height          =   495
      Index           =   18
      Left            =   8280
      Picture         =   "frmItemEdit.frx":43F0A
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   96
      Top             =   1080
      Width           =   495
   End
   Begin VB.PictureBox picItem 
      Height          =   495
      Index           =   19
      Left            =   8760
      Picture         =   "frmItemEdit.frx":44B4E
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   95
      Top             =   1080
      Width           =   495
   End
   Begin VB.PictureBox picItem 
      Height          =   495
      Index           =   20
      Left            =   9240
      Picture         =   "frmItemEdit.frx":45792
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   94
      Top             =   1080
      Width           =   495
   End
   Begin VB.Frame frmItem 
      Caption         =   "Item's"
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   4335
      Begin VB.CommandButton cmdAddToStore 
         Caption         =   "Add to Selected Store"
         Height          =   375
         Left            =   1200
         TabIndex        =   66
         Top             =   5160
         Width           =   1935
      End
      Begin VB.Frame Frame3 
         Caption         =   "Image"
         Height          =   975
         Left            =   120
         TabIndex        =   91
         Top             =   3600
         Width           =   1815
         Begin VB.TextBox txtImg 
            DataField       =   "Image"
            DataSource      =   "axsItem"
            Height          =   285
            Left            =   840
            TabIndex        =   117
            Top             =   120
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.ComboBox cmbImg 
            DataField       =   "Image"
            DataSource      =   "axsItem"
            Height          =   315
            ItemData        =   "frmItemEdit.frx":463D6
            Left            =   840
            List            =   "frmItemEdit.frx":46424
            TabIndex        =   93
            Text            =   "0"
            Top             =   360
            Width           =   855
         End
         Begin VB.PictureBox picshow 
            Height          =   540
            Left            =   120
            Picture         =   "frmItemEdit.frx":46472
            ScaleHeight     =   480
            ScaleWidth      =   480
            TabIndex        =   92
            Top             =   240
            Width           =   540
         End
      End
      Begin VB.Data axsItem 
         Caption         =   "Items"
         Connect         =   "Access 2000;"
         DatabaseName    =   "events\Item.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Item"
         Top             =   240
         Width           =   4095
      End
      Begin VB.Frame Frame1 
         Caption         =   "Item"
         Height          =   1695
         Left            =   120
         TabIndex        =   20
         Top             =   600
         Width           =   4095
         Begin VB.TextBox txtName 
            DataField       =   "Name"
            DataSource      =   "axsItem"
            Height          =   285
            Left            =   1320
            TabIndex        =   24
            Top             =   600
            Width           =   2655
         End
         Begin VB.ComboBox cmbType 
            Height          =   315
            ItemData        =   "frmItemEdit.frx":470B6
            Left            =   1320
            List            =   "frmItemEdit.frx":470E4
            TabIndex        =   23
            Text            =   "Combo1"
            Top             =   960
            Width           =   2655
         End
         Begin VB.TextBox txtSlot 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            DataField       =   "slot"
            DataSource      =   "axsItem"
            Height          =   285
            Left            =   1320
            Locked          =   -1  'True
            MaxLength       =   255
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   1320
            Width           =   495
         End
         Begin VB.TextBox txtEquip 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            DataField       =   "equip"
            DataSource      =   "axsItem"
            Height          =   285
            Left            =   2640
            Locked          =   -1  'True
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   1320
            Width           =   495
         End
         Begin VB.Label lblItemId 
            DataField       =   "ItemID"
            DataSource      =   "axsItem"
            Height          =   255
            Left            =   1320
            TabIndex        =   30
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Item ID:"
            Height          =   255
            Index           =   0
            Left            =   600
            TabIndex        =   29
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Item Name:"
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   28
            Top             =   600
            Width           =   855
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Item Type:"
            Height          =   255
            Index           =   11
            Left            =   360
            TabIndex        =   27
            Top             =   960
            Width           =   855
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Slot:"
            Height          =   255
            Index           =   9
            Left            =   720
            TabIndex        =   26
            Top             =   1320
            Width           =   495
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Equip:"
            Height          =   255
            Index           =   10
            Left            =   1320
            TabIndex        =   25
            Top             =   1320
            Width           =   1215
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Stats"
         Height          =   1335
         Left            =   120
         TabIndex        =   9
         Top             =   2280
         Width           =   4095
         Begin VB.TextBox txtResell 
            DataField       =   "resell"
            DataSource      =   "axsItem"
            Height          =   285
            Left            =   720
            TabIndex        =   74
            Top             =   960
            Width           =   1215
         End
         Begin VB.TextBox txtHp 
            DataField       =   "Hp"
            DataSource      =   "axsItem"
            Height          =   285
            Left            =   720
            TabIndex        =   14
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox txtMana 
            DataField       =   "Mana"
            DataSource      =   "axsItem"
            Height          =   285
            Left            =   720
            TabIndex        =   13
            Top             =   600
            Width           =   1215
         End
         Begin VB.TextBox txtStr 
            DataField       =   "Str"
            DataSource      =   "axsItem"
            Height          =   285
            Left            =   2760
            TabIndex        =   12
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox txtDex 
            DataField       =   "Dex"
            DataSource      =   "axsItem"
            Height          =   285
            Left            =   2760
            TabIndex        =   11
            Top             =   600
            Width           =   1215
         End
         Begin VB.TextBox txtSta 
            DataField       =   "Sta"
            DataSource      =   "axsItem"
            Height          =   285
            Left            =   2760
            TabIndex        =   10
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Resell:"
            Height          =   255
            Index           =   12
            Left            =   120
            TabIndex        =   75
            Top             =   960
            Width           =   495
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "HP:"
            Height          =   255
            Index           =   5
            Left            =   240
            TabIndex        =   19
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Mana:"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   18
            Top             =   600
            Width           =   495
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Str:"
            Height          =   255
            Index           =   2
            Left            =   1440
            TabIndex        =   17
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Dex:"
            Height          =   255
            Index           =   3
            Left            =   1440
            TabIndex        =   16
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Sta:"
            Height          =   255
            Index           =   4
            Left            =   1440
            TabIndex        =   15
            Top             =   960
            Width           =   1215
         End
      End
      Begin VB.Frame frmWeapon 
         Caption         =   "Weapon Propertys"
         Height          =   975
         Left            =   1920
         TabIndex        =   4
         Top             =   3600
         Width           =   2295
         Begin VB.TextBox txtDmg 
            DataField       =   "Dmg"
            DataSource      =   "axsItem"
            Height          =   285
            Left            =   840
            TabIndex        =   6
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox txtDly 
            DataField       =   "Dly"
            DataSource      =   "axsItem"
            Height          =   285
            Left            =   840
            TabIndex        =   5
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Dmg:"
            Height          =   255
            Index           =   7
            Left            =   360
            TabIndex        =   8
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Dly:"
            Height          =   255
            Index           =   8
            Left            =   240
            TabIndex        =   7
            Top             =   600
            Width           =   495
         End
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "New Item"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   4680
         Width           =   1215
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save Item"
         Height          =   375
         Left            =   1560
         TabIndex        =   2
         Top             =   4680
         Width           =   1215
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete Item"
         Height          =   375
         Left            =   2880
         TabIndex        =   1
         Top             =   4680
         Width           =   1215
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   6135
      Left            =   0
      TabIndex        =   65
      Top             =   240
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   10821
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Item Editor"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Store Editor"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label Label6 
      DataField       =   "StoresID"
      DataSource      =   "axsStore"
      Height          =   255
      Left            =   3000
      TabIndex        =   70
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Store:"
      Height          =   255
      Left            =   2400
      TabIndex        =   69
      Top             =   0
      Width           =   495
   End
   Begin VB.Label Label4 
      DataField       =   "ItemID"
      DataSource      =   "axsItem"
      Height          =   255
      Left            =   1680
      TabIndex        =   68
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Current Item:"
      Height          =   255
      Left            =   600
      TabIndex        =   67
      Top             =   0
      Width           =   975
   End
End
Attribute VB_Name = "frmItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub txtImg_Change()

picshow.Picture = picItem(txtImg.Text).Picture

End Sub

Private Sub cmbImg_Click()

txtImg.Text = cmbImg.ListIndex

End Sub

Private Sub cmbType_Click()

frmWeapon.Visible = False

If cmbType.ListIndex = 0 Then
txtSlot.Text = 12
txtEquip.Text = 1
ElseIf cmbType.ListIndex = 1 Then
txtSlot.Text = 9
txtEquip.Text = 1
ElseIf cmbType.ListIndex = 2 Then
txtSlot.Text = 8
txtEquip.Text = 1
ElseIf cmbType.ListIndex = 3 Then
txtSlot.Text = 10
txtEquip.Text = 1
ElseIf cmbType.ListIndex = 4 Then
txtSlot.Text = 11
txtEquip.Text = 1
ElseIf cmbType.ListIndex = 5 Then
txtSlot.Text = 6
txtEquip.Text = 1
ElseIf cmbType.ListIndex = 6 Then
txtSlot.Text = 3
txtEquip.Text = 1
ElseIf cmbType.ListIndex = 7 Then
txtSlot.Text = 13
txtEquip.Text = 0
ElseIf cmbType.ListIndex = 8 Then
txtSlot.Text = 7
txtEquip.Text = 1
ElseIf cmbType.ListIndex = 9 Then
txtSlot.Text = 2
txtEquip.Text = 1
ElseIf cmbType.ListIndex = 10 Then
txtSlot.Text = 4
txtEquip.Text = 1
ElseIf cmbType.ListIndex = 11 Then
txtSlot.Text = 5
txtEquip.Text = 1
ElseIf cmbType.ListIndex = 12 Then
frmWeapon.Visible = True
txtSlot.Text = 1
txtEquip.Text = 1
End If

End Sub

Private Sub cmdAddToStore_Click()
added = False

For i = 1 To 15

If Store(i).Text = lblItemId.Caption Then
added = True
msg = MsgBox("Current store already has that item.", vbOKOnly)
Exit Sub
Else
If Store(i).Text = 0 Then
added = True
Store(i).Text = lblItemId.Caption
Exit Sub
End If
End If
Next i

If added = False Then
msg = MsgBox("Current store is full.", vbOKOnly)
End If
End Sub

Private Sub cmdDelete_Click()
dodelete = MsgBox("Are you sure you want to delete this item?", vbOKCancel)
If dodelete = vbOK Then
axsItem.Recordset.Delete
End If
End Sub

Private Sub cmdDeleteStore_Click()
dodelete = MsgBox("Are you sure you want to delete this store?", vbOKCancel)
If dodelete = vbOK Then
axsStore.Recordset.Delete
End If
End Sub

Private Sub cmdNew_Click()
axsItem.Recordset.AddNew
axsItem.Recordset.Update
End Sub

Private Sub cmdNewStore_Click()
axsStore.Recordset.AddNew
axsStore.Recordset.Update
End Sub

Private Sub cmdRemove_Click()

For i = 1 To 15

If optItem(i).Value = True Then

For j = i To 14
Store(j).Text = Store(j + 1).Text
Store(15).Text = 0
Next j

Exit Sub
End If
Next i



End Sub

Private Sub cmdSave_Click()
axsItem.Recordset.Edit
axsItem.Recordset.Update
End Sub

Private Sub cmdSaveStore_Click()
axsStore.Recordset.Edit
axsStore.Recordset.Update
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub

Private Sub TabStrip1_Click()

If TabStrip1.SelectedItem.Index = 1 Then
frmStore.Visible = False
frmItem.Visible = True
ElseIf TabStrip1.SelectedItem.Index = 2 Then
frmStore.Visible = True
frmItem.Visible = False
End If
End Sub

Private Sub txtSlot_Change()
'On Error Resume Next

If txtSlot.Text = 0 Then
cmbType.ListIndex = 0
ElseIf txtSlot.Text = 1 Then
cmbType.ListIndex = 12
ElseIf txtSlot.Text = 2 Then
cmbType.ListIndex = 9
ElseIf txtSlot.Text = 3 Then
cmbType.ListIndex = 6
ElseIf txtSlot.Text = 4 Then
cmbType.ListIndex = 10
ElseIf txtSlot.Text = 5 Then
cmbType.ListIndex = 11
ElseIf txtSlot.Text = 6 Then
cmbType.ListIndex = 5
ElseIf txtSlot.Text = 7 Then
cmbType.ListIndex = 8
ElseIf txtSlot.Text = 8 Then
cmbType.ListIndex = 2
ElseIf txtSlot.Text = 9 Then
cmbType.ListIndex = 1
ElseIf txtSlot.Text = 10 Then
cmbType.ListIndex = 3
ElseIf txtSlot.Text = 11 Then
cmbType.ListIndex = 4
ElseIf txtSlot.Text = 12 Then
cmbType.ListIndex = 0
ElseIf txtSlot.Text = 13 Then
cmbType.ListIndex = 7
End If

End Sub
