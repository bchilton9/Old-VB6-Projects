VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmServer 
   Caption         =   "LoT Server"
   ClientHeight    =   7290
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7830
   Icon            =   "frmServer.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7290
   ScaleWidth      =   7830
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtItemResell 
      DataField       =   "resell"
      DataSource      =   "axsItem"
      Height          =   285
      Left            =   2400
      TabIndex        =   123
      Top             =   3600
      Width           =   735
   End
   Begin VB.CommandButton cmdDebugger 
      Caption         =   "Debugger"
      Height          =   255
      Left            =   1680
      TabIndex        =   122
      Top             =   1800
      Width           =   1455
   End
   Begin VB.TextBox txtItemImage 
      DataField       =   "Image"
      DataSource      =   "axsItem"
      Height          =   285
      Left            =   3840
      TabIndex        =   121
      Top             =   3360
      Width           =   735
   End
   Begin VB.TextBox txtGoldBank 
      DataField       =   "GoldBank"
      DataSource      =   "axsMember"
      Height          =   285
      Left            =   120
      TabIndex        =   120
      Top             =   6000
      Width           =   735
   End
   Begin VB.TextBox txtPrice 
      DataField       =   "Price15"
      DataSource      =   "axsStore"
      Height          =   285
      Index           =   15
      Left            =   2400
      TabIndex        =   119
      Top             =   6840
      Width           =   735
   End
   Begin VB.TextBox txtPrice 
      DataField       =   "Price14"
      DataSource      =   "axsStore"
      Height          =   285
      Index           =   14
      Left            =   3840
      TabIndex        =   118
      Top             =   6600
      Width           =   735
   End
   Begin VB.TextBox txtPrice 
      DataField       =   "Price13"
      DataSource      =   "axsStore"
      Height          =   285
      Index           =   13
      Left            =   3120
      TabIndex        =   117
      Top             =   6600
      Width           =   735
   End
   Begin VB.TextBox txtPrice 
      DataField       =   "Price12"
      DataSource      =   "axsStore"
      Height          =   285
      Index           =   12
      Left            =   2400
      TabIndex        =   116
      Top             =   6600
      Width           =   735
   End
   Begin VB.TextBox txtPrice 
      DataField       =   "Price11"
      DataSource      =   "axsStore"
      Height          =   285
      Index           =   11
      Left            =   3840
      TabIndex        =   115
      Top             =   6360
      Width           =   735
   End
   Begin VB.TextBox txtPrice 
      DataField       =   "Price10"
      DataSource      =   "axsStore"
      Height          =   285
      Index           =   10
      Left            =   3120
      TabIndex        =   114
      Top             =   6360
      Width           =   735
   End
   Begin VB.TextBox txtPrice 
      DataField       =   "Price9"
      DataSource      =   "axsStore"
      Height          =   285
      Index           =   9
      Left            =   2400
      TabIndex        =   113
      Top             =   6360
      Width           =   735
   End
   Begin VB.TextBox txtPrice 
      DataField       =   "Price8"
      DataSource      =   "axsStore"
      Height          =   285
      Index           =   8
      Left            =   3840
      TabIndex        =   112
      Top             =   6120
      Width           =   735
   End
   Begin VB.TextBox txtPrice 
      DataField       =   "Price7"
      DataSource      =   "axsStore"
      Height          =   285
      Index           =   7
      Left            =   3120
      TabIndex        =   111
      Top             =   6120
      Width           =   735
   End
   Begin VB.TextBox txtPrice 
      DataField       =   "Price6"
      DataSource      =   "axsStore"
      Height          =   285
      Index           =   6
      Left            =   2400
      TabIndex        =   110
      Top             =   6120
      Width           =   735
   End
   Begin VB.TextBox txtPrice 
      DataField       =   "Price5"
      DataSource      =   "axsStore"
      Height          =   285
      Index           =   5
      Left            =   3840
      TabIndex        =   109
      Top             =   5880
      Width           =   735
   End
   Begin VB.TextBox txtPrice 
      DataField       =   "Price4"
      DataSource      =   "axsStore"
      Height          =   285
      Index           =   4
      Left            =   3120
      TabIndex        =   108
      Top             =   5880
      Width           =   735
   End
   Begin VB.TextBox txtPrice 
      DataField       =   "Price3"
      DataSource      =   "axsStore"
      Height          =   285
      Index           =   3
      Left            =   2400
      TabIndex        =   107
      Top             =   5880
      Width           =   735
   End
   Begin VB.TextBox txtPrice 
      DataField       =   "Price2"
      DataSource      =   "axsStore"
      Height          =   285
      Index           =   2
      Left            =   3840
      TabIndex        =   106
      Top             =   5640
      Width           =   735
   End
   Begin VB.TextBox txtPrice 
      DataField       =   "Price1"
      DataSource      =   "axsStore"
      Height          =   285
      Index           =   1
      Left            =   3120
      TabIndex        =   105
      Top             =   5640
      Width           =   735
   End
   Begin VB.CommandButton cmdRestart 
      Caption         =   "Restart Server"
      Height          =   255
      Left            =   240
      TabIndex        =   104
      Top             =   1800
      Width           =   1455
   End
   Begin VB.TextBox txtAdminMsg 
      DataField       =   "meessage"
      DataSource      =   "axsMember"
      Height          =   285
      Left            =   1560
      TabIndex        =   103
      Top             =   5760
      Width           =   735
   End
   Begin VB.TextBox txtAdmin 
      DataField       =   "admin"
      DataSource      =   "axsMember"
      Height          =   285
      Left            =   840
      TabIndex        =   102
      Top             =   5760
      Width           =   735
   End
   Begin VB.TextBox txtLvl 
      DataField       =   "level"
      DataSource      =   "axsMember"
      Height          =   285
      Left            =   120
      TabIndex        =   101
      Top             =   5760
      Width           =   735
   End
   Begin VB.TextBox txtStore 
      DataField       =   "Store15"
      DataSource      =   "axsStore"
      Height          =   285
      Index           =   15
      Left            =   2400
      TabIndex        =   100
      Top             =   5640
      Width           =   735
   End
   Begin VB.TextBox txtStore 
      DataField       =   "Store14"
      DataSource      =   "axsStore"
      Height          =   285
      Index           =   14
      Left            =   3840
      TabIndex        =   99
      Top             =   5400
      Width           =   735
   End
   Begin VB.TextBox txtStore 
      DataField       =   "Store13"
      DataSource      =   "axsStore"
      Height          =   285
      Index           =   13
      Left            =   3120
      TabIndex        =   98
      Top             =   5400
      Width           =   735
   End
   Begin VB.TextBox txtStore 
      DataField       =   "Store12"
      DataSource      =   "axsStore"
      Height          =   285
      Index           =   12
      Left            =   2400
      TabIndex        =   97
      Top             =   5400
      Width           =   735
   End
   Begin VB.TextBox txtStore 
      DataField       =   "Store11"
      DataSource      =   "axsStore"
      Height          =   285
      Index           =   11
      Left            =   3840
      TabIndex        =   96
      Top             =   5160
      Width           =   735
   End
   Begin VB.TextBox txtStore 
      DataField       =   "Store10"
      DataSource      =   "axsStore"
      Height          =   285
      Index           =   10
      Left            =   3120
      TabIndex        =   95
      Top             =   5160
      Width           =   735
   End
   Begin VB.TextBox txtStore 
      DataField       =   "Store9"
      DataSource      =   "axsStore"
      Height          =   285
      Index           =   9
      Left            =   2400
      TabIndex        =   94
      Top             =   5160
      Width           =   735
   End
   Begin VB.TextBox txtStore 
      DataField       =   "Store8"
      DataSource      =   "axsStore"
      Height          =   285
      Index           =   8
      Left            =   3840
      TabIndex        =   93
      Top             =   4920
      Width           =   735
   End
   Begin VB.TextBox txtStore 
      DataField       =   "Store7"
      DataSource      =   "axsStore"
      Height          =   285
      Index           =   7
      Left            =   3120
      TabIndex        =   92
      Top             =   4920
      Width           =   735
   End
   Begin VB.TextBox txtStore 
      DataField       =   "Store6"
      DataSource      =   "axsStore"
      Height          =   285
      Index           =   6
      Left            =   2400
      TabIndex        =   91
      Top             =   4920
      Width           =   735
   End
   Begin VB.TextBox txtStore 
      DataField       =   "Store5"
      DataSource      =   "axsStore"
      Height          =   285
      Index           =   5
      Left            =   3840
      TabIndex        =   90
      Top             =   4680
      Width           =   735
   End
   Begin VB.TextBox txtStore 
      DataField       =   "Store4"
      DataSource      =   "axsStore"
      Height          =   285
      Index           =   4
      Left            =   3120
      TabIndex        =   89
      Top             =   4680
      Width           =   735
   End
   Begin VB.TextBox txtStore 
      DataField       =   "Store3"
      DataSource      =   "axsStore"
      Height          =   285
      Index           =   3
      Left            =   2400
      TabIndex        =   88
      Top             =   4680
      Width           =   735
   End
   Begin VB.TextBox txtStore 
      DataField       =   "Store2"
      DataSource      =   "axsStore"
      Height          =   285
      Index           =   2
      Left            =   3840
      TabIndex        =   87
      Top             =   4440
      Width           =   735
   End
   Begin VB.TextBox txtStore 
      DataField       =   "Store1"
      DataSource      =   "axsStore"
      Height          =   285
      Index           =   1
      Left            =   3120
      TabIndex        =   86
      Top             =   4440
      Width           =   735
   End
   Begin VB.TextBox txtStoreId 
      DataField       =   "StoresID"
      DataSource      =   "axsStore"
      Height          =   285
      Left            =   2400
      TabIndex        =   85
      Top             =   4440
      Width           =   735
   End
   Begin VB.Data axsStore 
      Caption         =   "Store"
      Connect         =   "Access 2000;"
      DatabaseName    =   "events\Item.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Stores"
      Top             =   4080
      Width           =   2175
   End
   Begin VB.TextBox txtItemEquip 
      DataField       =   "equip"
      DataSource      =   "axsItem"
      Height          =   285
      Left            =   3120
      TabIndex        =   84
      Top             =   3360
      Width           =   735
   End
   Begin VB.TextBox txtItemSlot 
      DataField       =   "slot"
      DataSource      =   "axsItem"
      Height          =   285
      Left            =   2400
      TabIndex        =   83
      Top             =   3360
      Width           =   735
   End
   Begin VB.TextBox txtItemDly 
      DataField       =   "Dly"
      DataSource      =   "axsItem"
      Height          =   285
      Left            =   3840
      TabIndex        =   82
      Top             =   3120
      Width           =   735
   End
   Begin VB.TextBox txtItemDmg 
      DataField       =   "Dmg"
      DataSource      =   "axsItem"
      Height          =   285
      Left            =   3120
      TabIndex        =   81
      Top             =   3120
      Width           =   735
   End
   Begin VB.TextBox txtItemMana 
      DataField       =   "Mana"
      DataSource      =   "axsItem"
      Height          =   285
      Left            =   2400
      TabIndex        =   80
      Top             =   3120
      Width           =   735
   End
   Begin VB.TextBox txtItemHp 
      DataField       =   "Hp"
      DataSource      =   "axsItem"
      Height          =   285
      Left            =   3840
      TabIndex        =   79
      Top             =   2880
      Width           =   735
   End
   Begin VB.TextBox txtItemSta 
      DataField       =   "Sta"
      DataSource      =   "axsItem"
      Height          =   285
      Left            =   3120
      TabIndex        =   78
      Top             =   2880
      Width           =   735
   End
   Begin VB.TextBox txtItemDex 
      DataField       =   "Dex"
      DataSource      =   "axsItem"
      Height          =   285
      Left            =   2400
      TabIndex        =   77
      Top             =   2880
      Width           =   735
   End
   Begin VB.TextBox txtItemStr 
      DataField       =   "Str"
      DataSource      =   "axsItem"
      Height          =   285
      Left            =   3840
      TabIndex        =   76
      Top             =   2640
      Width           =   735
   End
   Begin VB.TextBox txtItemName 
      DataField       =   "Name"
      DataSource      =   "axsItem"
      Height          =   285
      Left            =   3120
      TabIndex        =   75
      Top             =   2640
      Width           =   735
   End
   Begin VB.TextBox txtMana 
      DataField       =   "Mana"
      DataSource      =   "axsMember"
      Height          =   285
      Left            =   1560
      TabIndex        =   74
      Top             =   5520
      Width           =   735
   End
   Begin VB.TextBox txtHp 
      DataField       =   "HP"
      DataSource      =   "axsMember"
      Height          =   285
      Left            =   840
      TabIndex        =   73
      Top             =   5520
      Width           =   855
   End
   Begin VB.TextBox txtStr 
      DataField       =   "Str"
      DataSource      =   "axsMember"
      Height          =   285
      Left            =   120
      TabIndex        =   72
      Top             =   5520
      Width           =   735
   End
   Begin VB.TextBox txtSta 
      DataField       =   "Sta"
      DataSource      =   "axsMember"
      Height          =   285
      Left            =   1560
      TabIndex        =   71
      Top             =   5280
      Width           =   735
   End
   Begin VB.TextBox txtDex 
      DataField       =   "Dex"
      DataSource      =   "axsMember"
      Height          =   285
      Left            =   840
      TabIndex        =   70
      Top             =   5280
      Width           =   735
   End
   Begin VB.TextBox txtClass 
      DataField       =   "Class"
      DataSource      =   "axsMember"
      Height          =   285
      Left            =   120
      TabIndex        =   69
      Top             =   5280
      Width           =   735
   End
   Begin VB.TextBox txtItemId 
      DataField       =   "ItemId"
      DataSource      =   "axsItem"
      Height          =   285
      Left            =   2400
      TabIndex        =   68
      Top             =   2640
      Width           =   735
   End
   Begin VB.Data axsItem 
      Caption         =   "Items"
      Connect         =   "Access 2000;"
      DatabaseName    =   "events\Item.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Item"
      Top             =   2280
      Width           =   2175
   End
   Begin VB.TextBox txtEquip 
      DataField       =   "Equip44"
      DataSource      =   "axsMember"
      Height          =   285
      Index           =   44
      Left            =   1800
      TabIndex        =   67
      Top             =   5040
      Width           =   255
   End
   Begin VB.TextBox txtEquip 
      DataField       =   "Equip43"
      DataSource      =   "axsMember"
      Height          =   285
      Index           =   43
      Left            =   1560
      TabIndex        =   66
      Top             =   5040
      Width           =   255
   End
   Begin VB.TextBox txtEquip 
      DataField       =   "Equip42"
      DataSource      =   "axsMember"
      Height          =   285
      Index           =   42
      Left            =   1320
      TabIndex        =   65
      Top             =   5040
      Width           =   255
   End
   Begin VB.TextBox txtEquip 
      DataField       =   "Equip41"
      DataSource      =   "axsMember"
      Height          =   285
      Index           =   41
      Left            =   1080
      TabIndex        =   64
      Top             =   5040
      Width           =   255
   End
   Begin VB.TextBox txtEquip 
      DataField       =   "Equip40"
      DataSource      =   "axsMember"
      Height          =   285
      Index           =   40
      Left            =   840
      TabIndex        =   63
      Top             =   5040
      Width           =   255
   End
   Begin VB.TextBox txtEquip 
      DataField       =   "Equip39"
      DataSource      =   "axsMember"
      Height          =   285
      Index           =   39
      Left            =   600
      TabIndex        =   62
      Top             =   5040
      Width           =   255
   End
   Begin VB.TextBox txtEquip 
      DataField       =   "Equip38"
      DataSource      =   "axsMember"
      Height          =   285
      Index           =   38
      Left            =   360
      TabIndex        =   61
      Top             =   5040
      Width           =   255
   End
   Begin VB.TextBox txtEquip 
      DataField       =   "Equip37"
      DataSource      =   "axsMember"
      Height          =   285
      Index           =   37
      Left            =   120
      TabIndex        =   60
      Top             =   5040
      Width           =   255
   End
   Begin VB.TextBox txtEquip 
      DataField       =   "Equip36"
      DataSource      =   "axsMember"
      Height          =   285
      Index           =   36
      Left            =   2040
      TabIndex        =   59
      Top             =   4800
      Width           =   255
   End
   Begin VB.TextBox txtEquip 
      DataField       =   "Equip35"
      DataSource      =   "axsMember"
      Height          =   285
      Index           =   35
      Left            =   1800
      TabIndex        =   58
      Top             =   4800
      Width           =   255
   End
   Begin VB.TextBox txtEquip 
      DataField       =   "Equip34"
      DataSource      =   "axsMember"
      Height          =   285
      Index           =   34
      Left            =   1560
      TabIndex        =   57
      Top             =   4800
      Width           =   255
   End
   Begin VB.TextBox txtEquip 
      DataField       =   "Equip33"
      DataSource      =   "axsMember"
      Height          =   285
      Index           =   33
      Left            =   1320
      TabIndex        =   56
      Top             =   4800
      Width           =   255
   End
   Begin VB.TextBox txtEquip 
      DataField       =   "Equip32"
      DataSource      =   "axsMember"
      Height          =   285
      Index           =   32
      Left            =   1080
      TabIndex        =   55
      Top             =   4800
      Width           =   255
   End
   Begin VB.TextBox txtEquip 
      DataField       =   "Equip31"
      DataSource      =   "axsMember"
      Height          =   285
      Index           =   31
      Left            =   840
      TabIndex        =   54
      Top             =   4800
      Width           =   255
   End
   Begin VB.TextBox txtEquip 
      DataField       =   "Equip30"
      DataSource      =   "axsMember"
      Height          =   285
      Index           =   30
      Left            =   600
      TabIndex        =   53
      Top             =   4800
      Width           =   255
   End
   Begin VB.TextBox txtEquip 
      DataField       =   "Equip29"
      DataSource      =   "axsMember"
      Height          =   285
      Index           =   29
      Left            =   360
      TabIndex        =   52
      Top             =   4800
      Width           =   255
   End
   Begin VB.TextBox txtEquip 
      DataField       =   "Equip28"
      DataSource      =   "axsMember"
      Height          =   285
      Index           =   28
      Left            =   120
      TabIndex        =   51
      Top             =   4800
      Width           =   255
   End
   Begin VB.TextBox txtEquip 
      DataField       =   "Equip27"
      DataSource      =   "axsMember"
      Height          =   285
      Index           =   27
      Left            =   2040
      TabIndex        =   50
      Top             =   4560
      Width           =   255
   End
   Begin VB.TextBox txtEquip 
      DataField       =   "Equip26"
      DataSource      =   "axsMember"
      Height          =   285
      Index           =   26
      Left            =   1800
      TabIndex        =   49
      Top             =   4560
      Width           =   255
   End
   Begin VB.TextBox txtEquip 
      DataField       =   "Equip25"
      DataSource      =   "axsMember"
      Height          =   285
      Index           =   25
      Left            =   1560
      TabIndex        =   48
      Top             =   4560
      Width           =   255
   End
   Begin VB.TextBox txtEquip 
      DataField       =   "Equip24"
      DataSource      =   "axsMember"
      Height          =   285
      Index           =   24
      Left            =   1320
      TabIndex        =   47
      Top             =   4560
      Width           =   255
   End
   Begin VB.TextBox txtEquip 
      DataField       =   "Equip23"
      DataSource      =   "axsMember"
      Height          =   285
      Index           =   23
      Left            =   1080
      TabIndex        =   46
      Top             =   4560
      Width           =   255
   End
   Begin VB.TextBox txtEquip 
      DataField       =   "Equip22"
      DataSource      =   "axsMember"
      Height          =   285
      Index           =   22
      Left            =   840
      TabIndex        =   45
      Top             =   4560
      Width           =   255
   End
   Begin VB.TextBox txtEquip 
      DataField       =   "Equip21"
      DataSource      =   "axsMember"
      Height          =   285
      Index           =   21
      Left            =   600
      TabIndex        =   44
      Top             =   4560
      Width           =   255
   End
   Begin VB.TextBox txtEquip 
      DataField       =   "Equip20"
      DataSource      =   "axsMember"
      Height          =   285
      Index           =   20
      Left            =   360
      TabIndex        =   43
      Top             =   4560
      Width           =   255
   End
   Begin VB.TextBox txtEquip 
      DataField       =   "Equip19"
      DataSource      =   "axsMember"
      Height          =   285
      Index           =   19
      Left            =   120
      TabIndex        =   42
      Top             =   4560
      Width           =   255
   End
   Begin VB.TextBox txtEquip 
      DataField       =   "Equip18"
      DataSource      =   "axsMember"
      Height          =   285
      Index           =   18
      Left            =   2040
      TabIndex        =   41
      Top             =   4320
      Width           =   255
   End
   Begin VB.TextBox txtEquip 
      DataField       =   "Equip17"
      DataSource      =   "axsMember"
      Height          =   285
      Index           =   17
      Left            =   1800
      TabIndex        =   40
      Top             =   4320
      Width           =   255
   End
   Begin VB.TextBox txtEquip 
      DataField       =   "Equip16"
      DataSource      =   "axsMember"
      Height          =   285
      Index           =   16
      Left            =   1560
      TabIndex        =   39
      Top             =   4320
      Width           =   255
   End
   Begin VB.TextBox txtEquip 
      DataField       =   "Equip15"
      DataSource      =   "axsMember"
      Height          =   285
      Index           =   15
      Left            =   1320
      TabIndex        =   38
      Top             =   4320
      Width           =   255
   End
   Begin VB.TextBox txtEquip 
      DataField       =   "Equip14"
      DataSource      =   "axsMember"
      Height          =   285
      Index           =   14
      Left            =   1080
      TabIndex        =   37
      Top             =   4320
      Width           =   255
   End
   Begin VB.TextBox txtEquip 
      DataField       =   "Equip13"
      DataSource      =   "axsMember"
      Height          =   285
      Index           =   13
      Left            =   840
      TabIndex        =   36
      Top             =   4320
      Width           =   255
   End
   Begin VB.TextBox txtEquip 
      DataField       =   "Equip12"
      DataSource      =   "axsMember"
      Height          =   285
      Index           =   12
      Left            =   600
      TabIndex        =   35
      Top             =   4320
      Width           =   255
   End
   Begin VB.TextBox txtEquip 
      DataField       =   "Equip11"
      DataSource      =   "axsMember"
      Height          =   285
      Index           =   11
      Left            =   360
      TabIndex        =   34
      Top             =   4320
      Width           =   255
   End
   Begin VB.TextBox txtEquip 
      DataField       =   "Equip10"
      DataSource      =   "axsMember"
      Height          =   285
      Index           =   10
      Left            =   120
      TabIndex        =   33
      Top             =   4320
      Width           =   255
   End
   Begin VB.TextBox txtEquip 
      DataField       =   "Equip9"
      DataSource      =   "axsMember"
      Height          =   285
      Index           =   9
      Left            =   2040
      TabIndex        =   32
      Top             =   4080
      Width           =   255
   End
   Begin VB.TextBox txtEquip 
      DataField       =   "Equip8"
      DataSource      =   "axsMember"
      Height          =   285
      Index           =   8
      Left            =   1800
      TabIndex        =   31
      Top             =   4080
      Width           =   255
   End
   Begin VB.TextBox txtEquip 
      DataField       =   "Equip7"
      DataSource      =   "axsMember"
      Height          =   285
      Index           =   7
      Left            =   1560
      TabIndex        =   30
      Top             =   4080
      Width           =   255
   End
   Begin VB.TextBox txtEquip 
      DataField       =   "Equip6"
      DataSource      =   "axsMember"
      Height          =   285
      Index           =   6
      Left            =   1320
      TabIndex        =   29
      Top             =   4080
      Width           =   255
   End
   Begin VB.TextBox txtEquip 
      DataField       =   "Equip5"
      DataSource      =   "axsMember"
      Height          =   285
      Index           =   5
      Left            =   1080
      TabIndex        =   28
      Top             =   4080
      Width           =   255
   End
   Begin VB.TextBox txtEquip 
      DataField       =   "Equip4"
      DataSource      =   "axsMember"
      Height          =   285
      Index           =   4
      Left            =   840
      TabIndex        =   27
      Top             =   4080
      Width           =   255
   End
   Begin VB.TextBox txtEquip 
      DataField       =   "Equip3"
      DataSource      =   "axsMember"
      Height          =   285
      Index           =   3
      Left            =   600
      TabIndex        =   26
      Top             =   4080
      Width           =   255
   End
   Begin VB.TextBox txtEquip 
      DataField       =   "Equip2"
      DataSource      =   "axsMember"
      Height          =   285
      Index           =   2
      Left            =   360
      TabIndex        =   25
      Top             =   4080
      Width           =   255
   End
   Begin VB.TextBox txtEquip 
      DataField       =   "Equip1"
      DataSource      =   "axsMember"
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   24
      Top             =   4080
      Width           =   255
   End
   Begin VB.TextBox txtGold 
      DataField       =   "Gold"
      DataSource      =   "axsMember"
      Height          =   285
      Left            =   1560
      TabIndex        =   23
      Top             =   3840
      Width           =   735
   End
   Begin VB.TextBox txtOffy 
      DataField       =   "offy"
      DataSource      =   "axsMember"
      Height          =   285
      Left            =   840
      TabIndex        =   22
      Top             =   3840
      Width           =   735
   End
   Begin VB.TextBox txtOffx 
      DataField       =   "offx"
      DataSource      =   "axsMember"
      Height          =   285
      Left            =   120
      TabIndex        =   21
      Top             =   3840
      Width           =   735
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit Members"
      Height          =   255
      Left            =   1680
      TabIndex        =   20
      Top             =   1560
      Width           =   1455
   End
   Begin VB.TextBox txtBan 
      DataField       =   "ban"
      DataSource      =   "axsMember"
      Height          =   285
      Left            =   1560
      TabIndex        =   19
      Top             =   3600
      Width           =   735
   End
   Begin VB.CommandButton cmbKick 
      Caption         =   "Kick"
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "System"
      Height          =   1335
      Left            =   2400
      TabIndex        =   9
      Top             =   120
      Width           =   2295
      Begin VB.Label txtTime 
         Height          =   255
         Left            =   720
         TabIndex        =   18
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Online:"
         Height          =   255
         Left            =   -120
         TabIndex        =   17
         Top             =   960
         Width           =   735
      End
      Begin VB.Label txtPortB 
         Height          =   255
         Left            =   720
         TabIndex        =   15
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label txtPortA 
         Height          =   255
         Left            =   720
         TabIndex        =   14
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "PortB:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "PortA:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   495
      End
      Begin VB.Label txtIP 
         Height          =   255
         Left            =   720
         TabIndex        =   11
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Add:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.TextBox txtHeight 
      DataField       =   "Height"
      DataSource      =   "axsMember"
      Height          =   285
      Left            =   840
      TabIndex        =   8
      Top             =   3600
      Width           =   735
   End
   Begin VB.TextBox txtMask 
      DataField       =   "Mask"
      DataSource      =   "axsMember"
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Top             =   3600
      Width           =   855
   End
   Begin VB.TextBox txtY 
      DataField       =   "Y"
      DataSource      =   "axsMember"
      Height          =   285
      Left            =   1560
      TabIndex        =   6
      Top             =   3360
      Width           =   735
   End
   Begin VB.TextBox txtX 
      DataField       =   "X"
      DataSource      =   "axsMember"
      Height          =   285
      Left            =   840
      TabIndex        =   5
      Top             =   3360
      Width           =   735
   End
   Begin VB.TextBox txtLoc 
      DataField       =   "Location"
      DataSource      =   "axsMember"
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   3360
      Width           =   735
   End
   Begin VB.TextBox txtPassword 
      DataField       =   "password"
      DataSource      =   "axsMember"
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   2880
      Width           =   2175
   End
   Begin VB.TextBox txtEmail 
      DataField       =   "EMail"
      DataSource      =   "axsMember"
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   3120
      Width           =   2175
   End
   Begin VB.TextBox txtName 
      DataField       =   "name"
      DataSource      =   "axsMember"
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   2640
      Width           =   2175
   End
   Begin VB.Data axsMember 
      Caption         =   "Members"
      Connect         =   "Access 2000;"
      DatabaseName    =   "members.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "members"
      Top             =   2280
      Width           =   2175
   End
   Begin VB.ListBox lstPlayers 
      Height          =   1425
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   2175
   End
   Begin MSWinsockLib.Winsock wskServer 
      Index           =   0
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock wskSignup 
      Index           =   0
      Left            =   360
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Timer 
      Interval        =   60000
      Left            =   720
      Top             =   0
   End
   Begin VB.Menu mnuSystray 
      Caption         =   "Systray"
      Visible         =   0   'False
      Begin VB.Menu mnuRestore 
         Caption         =   "Restore / Open"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Edit Members"
      End
      Begin VB.Menu cmdMnuRestart 
         Caption         =   "Restart Server"
      End
      Begin VB.Menu Spacer1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Public Minute, Hour, Day As Integer
Public myMinute

Private Sub cmdDebugger_Click()
Load frmDebugger
frmDebugger.Show
End Sub

Private Sub cmdMnuRestart_Click()
cmdRestart_Click
End Sub

Private Sub cmdRestart_Click()
run = Shell(App.Path & "\reload.exe", vbNormalFocus)
End
End Sub

Sub loadItems()
Dim itemnum As Integer

axsItem.Recordset.MoveFirst
Do Until axsItem.Recordset.EOF

With Item(txtItemId.Text)
   .itemdex = Val(txtItemDex)
   .itemdly = Val(txtItemDly)
   .itemdmg = Val(txtItemDmg)
   .itemequip = txtItemEquip
   .itemhp = Val(txtItemHp)
   .itemImage = Val(txtItemImage)
   .itemmana = Val(txtItemMana)
   .itemName = txtItemName
   .itemslot = Val(txtItemSlot)
   .itemsta = Val(txtItemSta)
   .itemStr = Val(txtItemStr)
   .itemResell = Val(txtItemResell)
End With
axsItem.Recordset.MoveNext
Loop

With Item(0)
   .itemdex = 0
   .itemdly = 0
   .itemdmg = 0
   .itemequip = 0
   .itemhp = 0
   .itemImage = 0
   .itemmana = 0
   .itemName = 0
   .itemslot = 0
   .itemsta = 0
   .itemStr = 0
   .itemResell = 0
End With

itemsLoaded = True

End Sub

Private Sub Form_Load()

Me.WindowState = vbMinimized

Dim i, j As Long
Dim myevent, eventdata As String
itemsLoaded = False


SetParce

txtIP.Caption = serv
txtPortA.Caption = servPort

  wskSignup(0).LocalPort = servPort
  wskSignup(0).Listen
  
  For i = 1 To MaxPlayers
    Load wskServer(i)
    lstPlayers.AddItem i & ":"
    Load wskSignup(i)
  Next i
    
Dim mapcount As Integer

On Error GoTo donewithmaps
  For i = 1 To MaxMaps
j = 0
Open App.Path & "\events\event" & i & ".dat" For Input As #1
Do Until (EOF(1))

j = j + 1
Line Input #1, eventdata

If eventdata <> "" Then
myevent = Split(eventdata, pChar)
With events(i).myevent(j)
    .X = myevent(0)
    .Y = myevent(1)
    .type = myevent(2)
    .Tirgger = myevent(3)
    .location = myevent(4)
    .toX = myevent(5)
    .toY = myevent(6)
    .Action = myevent(7)
End With

End If
Loop

Close #1

Next i

donewithmaps:

Minute = 0
Day = 0
Hour = 0

txtTime.Caption = "0:0:00"


End Sub


Private Sub Timer_Timer()

    Minute = Minute + 1
    If Minute >= 60 Then
        Hour = Hour + 1
        Minute = 0
    End If
    myMinute = Minute
    If Hour >= 24 Then
        Day = Day + 1
        Hour = 0
    End If
    If Minute <= 9 Then
        myMinute = "0" & Minute
    End If
    
txtTime.Caption = Day & ":" & Hour & ":" & myMinute

End Sub

Private Sub wskServer_ConnectionRequest(Index As Integer, ByVal requestID As Long)
Dim i As Long
Dim Connected As Boolean

  i = 1
  Connected = False
  
  Do While Not Connected And (i < MaxPlayers)
    If (wskServer(i).State = sckClosed) Then
      wskServer(i).Accept requestID
      
      lstPlayers.List(i - 1) = i & ": " & "Connecting..."
      
      Connected = True
      wskServer(i).SendData "login" & pEnd
    End If
    i = i + 1
  Loop

End Sub

Private Sub wskServer_DataArrival(Index As Integer, ByVal bytesTotal As Long)
'On Error Resume Next
Dim s As String
Dim Packet() As String
Dim i As Long


  wskServer(Index).GetData s ', vbString, bytesTotal
  Packet = Split(s, pEnd)
  For i = 0 To UBound(Packet) - 1
    realtext Packet(i), Index
  Next i

End Sub

Sub realtext(txt As String, Index As Integer)

If frmDebugger.chkEnable.Value = 1 Then
frmDebugger.Text1.Text = frmDebugger.Text1.Text & txt & vbCrLf
End If

If itemsLoaded = False Then
loadItems
End If

'On Error Resume Next
Dim Parce() As String
Dim i, j, k As Integer
'Dim connected
Parce = Split(txt, pChar)

   'password is Parce(2)
 
If Parce(0) = "login" Then
'Connected = False
Player(Index).Connected = False

frmServer.axsMember.Recordset.MoveFirst
Do Until LCase(frmServer.txtName.Text) = LCase(Parce(1)) Or frmServer.axsMember.Recordset.EOF
frmServer.axsMember.Recordset.MoveNext
Loop

If LCase(frmServer.txtName.Text) = LCase(Parce(1)) Then
If LCase(frmServer.txtPassword.Text) = LCase(Parce(2)) Then

    With Player(Index)
        .mask = frmServer.txtMask
        .maskstep = 1
        .Name = frmServer.txtName.Text
        .X = frmServer.txtX
        .Y = frmServer.txtY
        .Container = .mask + 7
        .location = frmServer.txtLoc
        .Height = frmServer.txtHeight
        .Connected = True
        .offx = frmServer.txtOffx
        .offy = frmServer.txtOffy
        .gold = frmServer.txtGold
        .Admin = frmServer.txtAdmin
        .Adminmsg = frmServer.txtAdminMsg
        For k = 1 To 44
        .inv.invitem(k) = txtEquip(k)
        Next k
    End With
    
For i = 1 To MaxPlayers
If i <> Index Then
    If (wskServer(i).State <> sckClosed) Then
    wskServer(i).SendData "player" & pChar & Val(Index) & pChar & Player(Index).Name & pChar & Player(Index).location & pChar & Player(Index).X & pChar & Player(Index).Y & pChar & Player(Index).mask & pChar & Player(Index).Height & pChar & Player(Index).offx & pChar & Player(Index).offy & pChar & Player(Index).mask & pEnd
    End If
End If
Next i

For i = 1 To MaxPlayers
    If (wskServer(i).State <> sckClosed) Then
    wskServer(Index).SendData "player" & pChar & Val(i) & pChar & Player(i).Name & pChar & Player(i).location & pChar & Player(i).X & pChar & Player(i).Y & pChar & Player(i).mask & pChar & Player(i).Height & pChar & Player(i).offx & pChar & Player(i).offy & pChar & Player(i).mask & pEnd
    End If
Next i

    lstPlayers.List(Index - 1) = Index & ": " & Parce(1)
    wskServer(Index).SendData "logedin" & pChar & Val(Index) & pEnd
    If Player(Index).Adminmsg = "" Then
    'do nuthing
    Else
    wskServer(Index).SendData "adminmsg" & pChar & Player(Index).Adminmsg & pEnd
    
    End If

'Connected = True
End If
End If

If Player(Index).Connected = False Then
wskServer(Index).SendData "nouser" & pEnd
lstPlayers.List(Index - 1) = Index & ": "
End If

End If

If Parce(0) = "gotadminmsg" Then
frmServer.axsMember.Recordset.MoveFirst
Do Until LCase(frmServer.txtName.Text) = LCase(Player(Index).Name) Or frmServer.axsMember.Recordset.EOF
frmServer.axsMember.Recordset.MoveNext
Loop

If LCase(frmServer.txtName.Text) = LCase(Player(Index).Name) Then
txtAdminMsg.Text = ""
axsMember.Recordset.Edit
axsMember.Recordset.Update
End If
    
Player(Index).Adminmsg = ""
End If

If Parce(0) = "admin" Then
If Player(Index).Admin = "yes" Then
wskServer(Index).SendData "admin yes you are" & pEnd
End If
End If

If Parce(0) = "adminmsg" Then
If Player(Index).Admin = "yes" Then

frmServer.axsMember.Recordset.MoveFirst
Do Until LCase(frmServer.txtName.Text) = LCase(Parce(1)) Or frmServer.axsMember.Recordset.EOF
frmServer.axsMember.Recordset.MoveNext
Loop

If LCase(frmServer.txtName.Text) = LCase(Parce(1)) Then
frmServer.axsMember.Recordset.Edit
frmServer.txtAdminMsg.Text = frmServer.txtAdminMsg.Text & "VBCRLF" & Parce(2)
frmServer.axsMember.Recordset.Update
wskServer(Index).SendData "system" & pChar & "Message Sent!" & pEnd
Else
wskServer(Index).SendData "system" & pChar & "Message not Sent!" & pEnd
End If
End If
End If

If Parce(0) = "reload" Then
If Player(Index).Admin = "yes" Then
wskServer(Index).SendData "system" & pChar & "Reloading Server!" & pEnd
run = Shell(App.Path & "\reload.exe", vbNormalFocus)
Shell_NotifyIcon NIM_DELETE, nid
End
End If
End If

If Parce(0) = "logout" Then

frmServer.axsMember.Recordset.MoveFirst
Do Until LCase(frmServer.txtName.Text) = LCase(Player(Index).Name) Or frmServer.axsMember.Recordset.EOF
frmServer.axsMember.Recordset.MoveNext
Loop

frmServer.txtX = Player(Index).X
frmServer.txtY = Player(Index).Y
frmServer.txtLoc = Player(Index).location
frmServer.txtHeight = Player(Index).Height
frmServer.txtGold = Player(Index).gold
'frmServer.txtGoldBank = player(Index).bankgold

frmServer.axsMember.UpdateRecord
    
    With Player(Index)
        .mask = 0
        .maskstep = 0
        .Name = ""
        .X = 0
        .Y = 0
        .Container = 0
        .location = 1
        .Height = 0
        .offx = 0
        .offy = 0
        .gold = 0
        .Admin = ""
        .Adminmsg = ""
    End With

For i = 1 To MaxPlayers
If i <> Index Then
    If (wskServer(i).State <> sckClosed) Then
    wskServer(i).SendData "close" & pChar & Val(Index) & pEnd
    End If
End If
Next i

wskServer(Index).SendData "logedout" & pEnd

lstPlayers.List(Index - 1) = Index & ": "
End If

If Parce(0) = "getinventory" Then
sendinventory Index
End If

If Parce(0) = "moveitem" Then
Dim save As Integer
save = Player(Index).inv.invitem(Val(Parce(1)))

If Val(Parce(1)) > 12 And Val(Parce(2)) > 12 Then
'move with in inv
ElseIf Val(Parce(1)) < 13 And Val(Parce(2)) > 12 And Item(Player(Index).inv.invitem(Val(Parce(2)))).itemslot = Val(Parce(1)) Then
'move from inv to equiped
ElseIf Val(Parce(1)) < 13 And Val(Parce(2)) > 12 And Player(Index).inv.invitem(Val(Parce(2))) = 0 Then
'move from inv to equiped item 0
ElseIf Val(Parce(1)) > 12 And Val(Parce(2)) < 13 And Item(Player(Index).inv.invitem(Val(Parce(1)))).itemslot = Val(Parce(2)) Then
'move from equiped to inv
ElseIf Val(Parce(1)) > 12 And Val(Parce(2)) < 13 And Item(Player(Index).inv.invitem(Val(Parce(1)))).itemslot = 0 Then
'move from equiped to inv item 0
Else
wskServer(Index).SendData "system" & pChar & "That item can not go in that slot." & pEnd
GoTo nomove
End If


frmServer.axsMember.Recordset.MoveFirst
Do Until LCase(frmServer.txtName.Text) = LCase(Player(Index).Name) Or frmServer.axsMember.Recordset.EOF
frmServer.axsMember.Recordset.MoveNext
Loop

If LCase(frmServer.txtName.Text) = LCase(Player(Index).Name) Then

frmServer.axsMember.Recordset.Edit

txtEquip(Val(Parce(1))).Text = txtEquip(Val(Parce(2))).Text
txtEquip(Val(Parce(2))).Text = save

frmServer.axsMember.Recordset.Update

Player(Index).inv.invitem(Val(Parce(1))) = Player(Index).inv.invitem(Val(Parce(2)))
Player(Index).inv.invitem(Val(Parce(2))) = save

sendinventory Index

Else
wskServer(Index).SendData "system" & pChar & "Error. Unable to move item please contact an admin if this problem happens agine." & pEnd
End If


nomove:
End If


If Parce(0) = "isevent" Then

For j = 0 To MaxEvents
If Val(events(Player(Index).location).myevent(j).Tirgger) = 0 Then
If Val(events(Player(Index).location).myevent(j).X) * PicX = Val(Parce(1)) And Val(events(Player(Index).location).myevent(j).Y) * PicY = Val(Parce(2)) Then

    If events(Player(Index).location).myevent(j).type = "warp" Then
        With Player(Index)
            .X = events(Parce(4)).myevent(j).toX * PicX
            .Y = events(Parce(4)).myevent(j).toY * PicY
            .location = events(Parce(4)).myevent(j).location
        End With
    
    ElseIf events(Player(Index).location).myevent(j).type = "sign" Then
        wskServer(Index).SendData "sign" & pChar & events(Player(Index).location).myevent(j).Action & pEnd

    ElseIf events(Player(Index).location).myevent(j).type = "store" Then
    sendinventory Index
        wskServer(Index).SendData "store" & pChar & events(Player(Index).location).myevent(j).Action & pEnd

    ElseIf events(Player(Index).location).myevent(j).type = "bank" Then
    sendinventory Index
        wskServer(Index).SendData "bank" & pEnd

    End If

End If
End If
Next j

End If

If Parce(0) = "isclickevent" Then

For j = 0 To MaxEvents
If Val(events(Player(Index).location).myevent(j).Tirgger) = 2 Then
If events(Player(Index).location).myevent(j).X = Val(Parce(1)) And events(Player(Index).location).myevent(j).Y = Val(Parce(2)) Then

    If events(Player(Index).location).myevent(j).type = "warp" Then
        With Player(Index)
            .X = events(Parce(4)).myevent(j).toX * PicX
            .Y = events(Parce(4)).myevent(j).toY * PicY
            .location = events(Parce(4)).myevent(j).location
        End With
    
    ElseIf events(Player(Index).location).myevent(j).type = "sign" Then
        wskServer(Index).SendData "sign" & pChar & events(Player(Index).location).myevent(j).Action & pEnd

    ElseIf events(Player(Index).location).myevent(j).type = "store" Then
    sendinventory Index
        wskServer(Index).SendData "store" & pChar & events(Player(Index).location).myevent(j).Action & pEnd

    ElseIf events(Player(Index).location).myevent(j).type = "bank" Then
    sendinventory Index
        wskServer(Index).SendData "bank" & pEnd

    End If

End If
End If
Next j

End If

If Parce(0) = "move" Then

With Player(Index)
.Container = Val(Parce(1))
.X = Parce(2)
.Y = Parce(3)
.location = Parce(4)
.maskstep = Parce(5)
.Height = Parce(6)
.toX = Parce(7)
.toY = Parce(8)
End With

For j = 0 To MaxEvents
If Val(events(Parce(4)).myevent(j).Tirgger) = 1 Then
If Val(events(Parce(4)).myevent(j).X) * PicX = Val(Player(Index).toX) And Val(events(Parce(4)).myevent(j).Y) * PicY = Val(Player(Index).toY) Then

    If events(Parce(4)).myevent(j).type = "warp" Then
        With Player(Index)
            .X = events(Parce(4)).myevent(j).toX * PicX
            .Y = events(Parce(4)).myevent(j).toY * PicY
            .location = events(Parce(4)).myevent(j).location
        End With

    ElseIf events(Parce(4)).myevent(j).type = "sign" Then
        wskServer(Index).SendData "sign" & pChar & events(Parce(4)).myevent(j).Action & pEnd

    ElseIf events(Parce(4)).myevent(j).type = "store" Then
    sendinventory Index
        wskServer(Index).SendData "store" & pChar & events(Parce(4)).myevent(j).Action & pEnd

    ElseIf events(Parce(4)).myevent(j).type = "bank" Then
    sendinventory Index
        wskServer(Index).SendData "bank" & pEnd

    End If

End If
End If
Next j

        For i = 1 To MaxPlayers
        If (wskServer(i).State <> sckClosed) Then
            wskServer(i).SendData "move" & pChar & Val(Index) & pChar & Player(Index).Container & pChar & Player(Index).X & pChar & Player(Index).Y & pChar & Player(Index).location & pChar & Player(Index).maskstep & pChar & Player(Index).Height & pChar & Player(Index).offx & pChar & Player(Index).offy & pChar & Player(Index).mask & pEnd
        End If
        Next i

End If

If Parce(0) = "chat" Then
    For i = 1 To MaxPlayers
        If Player(i).location = Player(Index).location Then
            wskServer(i).SendData "chat" & pChar & Player(Index).Name & pChar & Parce(1) & pEnd
        End If
    Next i
End If

If Parce(0) = "who" Then
    wskServer(Index).SendData "who" & pChar
        For i = 1 To MaxPlayers
            If (wskServer(i).State <> sckClosed) Then
                wskServer(Index).SendData Player(i).Name & ", "
            End If
        Next i
    wskServer(Index).SendData pEnd
End If

If Parce(0) = "tell" Then
Dim done
done = False

If LCase(Player(Index).Name) = LCase(Parce(1)) Then
    wskServer(Index).SendData "system" & pChar & "You can't send a tell to yourself!</FONT>" & pEnd
Else
    For i = 1 To MaxPlayers
        If LCase(Player(i).Name) = LCase(Parce(1)) Then
            wskServer(Index).SendData "youtell" & pChar & Player(i).Name & pChar & Parce(2) & pEnd
            wskServer(i).SendData "tellyou" & pChar & Player(Index).Name & pChar & Parce(2) & pEnd
            done = True
        End If
    Next i
    If done = False Then
        wskServer(Index).SendData "system" & pChar & Parce(1) & " is not online!" & pEnd
    End If
End If

End If

'End If

End Sub

Sub sendinventory(Index As Integer)
Dim invstr, invImg As String

        For k = 1 To 44
        itemnum = Player(Index).inv.invitem(k)
        
        invstr = invstr & pChar & Player(Index).inv.invitem(k)
        invImg = invImg & pChar & Item(Player(Index).inv.invitem(k)).itemImage
        Next k

wskServer(Index).SendData "inventory" & invstr & invImg & pChar & Player(Index).gold & pEnd

End Sub

Private Sub wskServer_Close(Index As Integer)
Dim i As Integer
    With Player(Index)
        .mask = 0
        .maskstep = 0
        .Name = ""
        .X = 0
        .Y = 0
        .Container = 0
        .location = 1
        .Height = 0
    End With

For i = 1 To MaxPlayers
If i <> Index Then
    If (wskServer(i).State <> sckClosed) Then
    wskServer(i).SendData "close" & pChar & Val(Index) & pEnd
    End If
End If
Next i

lstPlayers.List(Index - 1) = Index & ": "
wskServer(Index).Close


End Sub

Private Sub wskSignup_ConnectionRequest(Index As Integer, ByVal requestID As Long)
Dim i As Long
Dim Connected As Boolean

  i = 1
  Connected = False
  
  Do While Not Connected And (i < MaxPlayers)
    If (wskSignup(i).State = sckClosed) Then
      wskSignup(i).Accept requestID
      Connected = True
      wskSignup(i).SendData "sendinfo" & pEnd
    End If
    i = i + 1
  Loop

End Sub

Private Sub wskSignup_DataArrival(Index As Integer, ByVal bytesTotal As Long)
'On Error Resume Next
Dim s As String
Dim Packet() As String
Dim i As Long


  wskSignup(Index).GetData s ', vbString, bytesTotal
  Packet = Split(s, pEnd)
  For i = 0 To UBound(Packet) - 1
    realtextsignup Packet(i), Index
  Next i

End Sub

Sub realtextsignup(txt As String, Index As Integer)
Dim Parce() As String
Dim i As Integer
Dim Connected
Parce = Split(txt, pChar)

If Parce(0) = "newuser" Then

'User Parce(1)
'Pass Parce(2)
'Email Parce(3)
'mask Parce(4)

frmServer.axsMember.Recordset.AddNew

frmServer.txtEmail.Text = Parce(3)
frmServer.txtHeight.Text = 30
frmServer.txtLoc.Text = 1
frmServer.txtMask.Text = Parce(4)
frmServer.txtName.Text = Parce(1)
frmServer.txtPassword.Text = Parce(2)
frmServer.txtX.Text = 320
frmServer.txtY.Text = 320
frmServer.txtOffx.Text = Parce(5)
frmServer.txtOffy.Text = Parce(6)
frmServer.axsMember.Recordset.Update

wskSignup(Index).SendData "usercreated" & pEnd
End If

End Sub

'THIS MAKES THE MENU POPUP WHEN THE FORM IS HIDDEN IN THE SYSTRAY'
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Sys As Long
Sys = X / Screen.TwipsPerPixelX
Select Case Sys
Case WM_LBUTTONDOWN:
Me.PopupMenu mnuSystray
End Select
End Sub

'THIS MAKES THE FOR DISSAPEAR/MINIMIZE TO THE SYSTRAY'
Private Sub Form_Resize()
If WindowState = vbMinimized Then
Me.Hide
Me.Refresh
With nid
.cbSize = Len(nid)
.hwnd = Me.hwnd
.uId = vbNull
.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
.uCallBackMessage = WM_MOUSEMOVE
.hIcon = Me.Icon
.szTip = Me.Caption & vbNullChar
End With
Shell_NotifyIcon NIM_ADD, nid
Else
Shell_NotifyIcon NIM_DELETE, nid
End If


If Me.WindowState = 2 Then
Me.WindowState = vbNormal
End If

  If Me.WindowState <> 1 And Me.WindowState <> 2 Then
    Me.Width = 4920
    Me.Height = 2645
  End If

End Sub

Private Sub mnuRestore_Click()
WindowState = vbNormal
Me.Show
End Sub

'THIS WILL KILL THE SYSTRAY ICON IF THE FORM IS UNLOADED'
Private Sub Form_Unload(Cancel As Integer)
Shell_NotifyIcon NIM_DELETE, nid
End
End Sub

'THIS UNLOADS THE FORM FROM THE MENU'
Private Sub mnuexit_Click()
Shell_NotifyIcon NIM_DELETE, nid
End
End Sub

Private Sub cmdEdit_Click()
Load frmEdit
frmEdit.Show
End Sub

Private Sub mnuEdit_Click()
Load frmEdit
frmEdit.Show
End Sub
