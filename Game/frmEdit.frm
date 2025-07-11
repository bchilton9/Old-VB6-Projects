VERSION 5.00
Begin VB.Form frmEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Members"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   4710
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtAdminMsg 
      DataField       =   "meessage"
      DataSource      =   "axsMember"
      Height          =   285
      Left            =   960
      TabIndex        =   79
      Top             =   4080
      Width           =   2535
   End
   Begin VB.TextBox txtIsadmin 
      DataField       =   "admin"
      DataSource      =   "axsMember"
      Height          =   285
      Left            =   960
      TabIndex        =   78
      Top             =   3720
      Width           =   735
   End
   Begin VB.TextBox txtLevel 
      DataField       =   "level"
      DataSource      =   "axsMember"
      Height          =   285
      Left            =   2400
      TabIndex        =   67
      Top             =   2280
      Width           =   735
   End
   Begin VB.TextBox txtEquip 
      DataField       =   "Equip1"
      DataSource      =   "axsMember"
      Height          =   285
      Index           =   1
      Left            =   3600
      TabIndex        =   66
      Top             =   840
      Width           =   255
   End
   Begin VB.TextBox txtEquip 
      DataField       =   "Equip2"
      DataSource      =   "axsMember"
      Height          =   285
      Index           =   2
      Left            =   3840
      TabIndex        =   65
      Top             =   840
      Width           =   255
   End
   Begin VB.TextBox txtEquip 
      DataField       =   "Equip3"
      DataSource      =   "axsMember"
      Height          =   285
      Index           =   3
      Left            =   4080
      TabIndex        =   64
      Top             =   840
      Width           =   255
   End
   Begin VB.TextBox txtEquip 
      DataField       =   "Equip4"
      DataSource      =   "axsMember"
      Height          =   285
      Index           =   4
      Left            =   4320
      TabIndex        =   63
      Top             =   840
      Width           =   255
   End
   Begin VB.TextBox txtEquip 
      DataField       =   "Equip5"
      DataSource      =   "axsMember"
      Height          =   285
      Index           =   5
      Left            =   3600
      TabIndex        =   62
      Top             =   1080
      Width           =   255
   End
   Begin VB.TextBox txtEquip 
      DataField       =   "Equip6"
      DataSource      =   "axsMember"
      Height          =   285
      Index           =   6
      Left            =   3840
      TabIndex        =   61
      Top             =   1080
      Width           =   255
   End
   Begin VB.TextBox txtEquip 
      DataField       =   "Equip7"
      DataSource      =   "axsMember"
      Height          =   285
      Index           =   7
      Left            =   4080
      TabIndex        =   60
      Top             =   1080
      Width           =   255
   End
   Begin VB.TextBox txtEquip 
      DataField       =   "Equip8"
      DataSource      =   "axsMember"
      Height          =   285
      Index           =   8
      Left            =   4320
      TabIndex        =   59
      Top             =   1080
      Width           =   255
   End
   Begin VB.TextBox txtEquip 
      DataField       =   "Equip9"
      DataSource      =   "axsMember"
      Height          =   285
      Index           =   9
      Left            =   3600
      TabIndex        =   58
      Top             =   1320
      Width           =   255
   End
   Begin VB.TextBox txtEquip 
      DataField       =   "Equip10"
      DataSource      =   "axsMember"
      Height          =   285
      Index           =   10
      Left            =   3840
      TabIndex        =   57
      Top             =   1320
      Width           =   255
   End
   Begin VB.TextBox txtEquip 
      DataField       =   "Equip11"
      DataSource      =   "axsMember"
      Height          =   285
      Index           =   11
      Left            =   4080
      TabIndex        =   56
      Top             =   1320
      Width           =   255
   End
   Begin VB.TextBox txtEquip 
      DataField       =   "Equip12"
      DataSource      =   "axsMember"
      Height          =   285
      Index           =   12
      Left            =   4320
      TabIndex        =   55
      Top             =   1320
      Width           =   255
   End
   Begin VB.TextBox txtEquip 
      DataField       =   "Equip13"
      DataSource      =   "axsMember"
      Height          =   285
      Index           =   13
      Left            =   3600
      TabIndex        =   54
      Top             =   1920
      Width           =   255
   End
   Begin VB.TextBox txtEquip 
      DataField       =   "Equip14"
      DataSource      =   "axsMember"
      Height          =   285
      Index           =   14
      Left            =   3840
      TabIndex        =   53
      Top             =   1920
      Width           =   255
   End
   Begin VB.TextBox txtEquip 
      DataField       =   "Equip15"
      DataSource      =   "axsMember"
      Height          =   285
      Index           =   15
      Left            =   4080
      TabIndex        =   52
      Top             =   1920
      Width           =   255
   End
   Begin VB.TextBox txtEquip 
      DataField       =   "Equip16"
      DataSource      =   "axsMember"
      Height          =   285
      Index           =   16
      Left            =   4320
      TabIndex        =   51
      Top             =   1920
      Width           =   255
   End
   Begin VB.TextBox txtEquip 
      DataField       =   "Equip17"
      DataSource      =   "axsMember"
      Height          =   285
      Index           =   17
      Left            =   3600
      TabIndex        =   50
      Top             =   2160
      Width           =   255
   End
   Begin VB.TextBox txtEquip 
      DataField       =   "Equip18"
      DataSource      =   "axsMember"
      Height          =   285
      Index           =   18
      Left            =   3840
      TabIndex        =   49
      Top             =   2160
      Width           =   255
   End
   Begin VB.TextBox txtEquip 
      DataField       =   "Equip19"
      DataSource      =   "axsMember"
      Height          =   285
      Index           =   19
      Left            =   4080
      TabIndex        =   48
      Top             =   2160
      Width           =   255
   End
   Begin VB.TextBox txtEquip 
      DataField       =   "Equip20"
      DataSource      =   "axsMember"
      Height          =   285
      Index           =   20
      Left            =   4320
      TabIndex        =   47
      Top             =   2160
      Width           =   255
   End
   Begin VB.TextBox txtEquip 
      DataField       =   "Equip21"
      DataSource      =   "axsMember"
      Height          =   285
      Index           =   21
      Left            =   3600
      TabIndex        =   46
      Top             =   2400
      Width           =   255
   End
   Begin VB.TextBox txtEquip 
      DataField       =   "Equip22"
      DataSource      =   "axsMember"
      Height          =   285
      Index           =   22
      Left            =   3840
      TabIndex        =   45
      Top             =   2400
      Width           =   255
   End
   Begin VB.TextBox txtEquip 
      DataField       =   "Equip23"
      DataSource      =   "axsMember"
      Height          =   285
      Index           =   23
      Left            =   4080
      TabIndex        =   44
      Top             =   2400
      Width           =   255
   End
   Begin VB.TextBox txtEquip 
      DataField       =   "Equip24"
      DataSource      =   "axsMember"
      Height          =   285
      Index           =   24
      Left            =   4320
      TabIndex        =   43
      Top             =   2400
      Width           =   255
   End
   Begin VB.TextBox txtEquip 
      DataField       =   "Equip25"
      DataSource      =   "axsMember"
      Height          =   285
      Index           =   25
      Left            =   3600
      TabIndex        =   42
      Top             =   2640
      Width           =   255
   End
   Begin VB.TextBox txtEquip 
      DataField       =   "Equip26"
      DataSource      =   "axsMember"
      Height          =   285
      Index           =   26
      Left            =   3840
      TabIndex        =   41
      Top             =   2640
      Width           =   255
   End
   Begin VB.TextBox txtEquip 
      DataField       =   "Equip27"
      DataSource      =   "axsMember"
      Height          =   285
      Index           =   27
      Left            =   4080
      TabIndex        =   40
      Top             =   2640
      Width           =   255
   End
   Begin VB.TextBox txtEquip 
      DataField       =   "Equip28"
      DataSource      =   "axsMember"
      Height          =   285
      Index           =   28
      Left            =   4320
      TabIndex        =   39
      Top             =   2640
      Width           =   255
   End
   Begin VB.TextBox txtEquip 
      DataField       =   "Equip29"
      DataSource      =   "axsMember"
      Height          =   285
      Index           =   29
      Left            =   3600
      TabIndex        =   38
      Top             =   3240
      Width           =   255
   End
   Begin VB.TextBox txtEquip 
      DataField       =   "Equip30"
      DataSource      =   "axsMember"
      Height          =   285
      Index           =   30
      Left            =   3840
      TabIndex        =   37
      Top             =   3240
      Width           =   255
   End
   Begin VB.TextBox txtEquip 
      DataField       =   "Equip31"
      DataSource      =   "axsMember"
      Height          =   285
      Index           =   31
      Left            =   4080
      TabIndex        =   36
      Top             =   3240
      Width           =   255
   End
   Begin VB.TextBox txtEquip 
      DataField       =   "Equip32"
      DataSource      =   "axsMember"
      Height          =   285
      Index           =   32
      Left            =   4320
      TabIndex        =   35
      Top             =   3240
      Width           =   255
   End
   Begin VB.TextBox txtEquip 
      DataField       =   "Equip33"
      DataSource      =   "axsMember"
      Height          =   285
      Index           =   33
      Left            =   3600
      TabIndex        =   34
      Top             =   3480
      Width           =   255
   End
   Begin VB.TextBox txtEquip 
      DataField       =   "Equip34"
      DataSource      =   "axsMember"
      Height          =   285
      Index           =   34
      Left            =   3840
      TabIndex        =   33
      Top             =   3480
      Width           =   255
   End
   Begin VB.TextBox txtEquip 
      DataField       =   "Equip35"
      DataSource      =   "axsMember"
      Height          =   285
      Index           =   35
      Left            =   4080
      TabIndex        =   32
      Top             =   3480
      Width           =   255
   End
   Begin VB.TextBox txtEquip 
      DataField       =   "Equip36"
      DataSource      =   "axsMember"
      Height          =   285
      Index           =   36
      Left            =   4320
      TabIndex        =   31
      Top             =   3480
      Width           =   255
   End
   Begin VB.TextBox txtEquip 
      DataField       =   "Equip37"
      DataSource      =   "axsMember"
      Height          =   285
      Index           =   37
      Left            =   3600
      TabIndex        =   30
      Top             =   3720
      Width           =   255
   End
   Begin VB.TextBox txtEquip 
      DataField       =   "Equip38"
      DataSource      =   "axsMember"
      Height          =   285
      Index           =   38
      Left            =   3840
      TabIndex        =   29
      Top             =   3720
      Width           =   255
   End
   Begin VB.TextBox txtEquip 
      DataField       =   "Equip39"
      DataSource      =   "axsMember"
      Height          =   285
      Index           =   39
      Left            =   4080
      TabIndex        =   28
      Top             =   3720
      Width           =   255
   End
   Begin VB.TextBox txtEquip 
      DataField       =   "Equip40"
      DataSource      =   "axsMember"
      Height          =   285
      Index           =   40
      Left            =   4320
      TabIndex        =   27
      Top             =   3720
      Width           =   255
   End
   Begin VB.TextBox txtEquip 
      DataField       =   "Equip41"
      DataSource      =   "axsMember"
      Height          =   285
      Index           =   41
      Left            =   3600
      TabIndex        =   26
      Top             =   3960
      Width           =   255
   End
   Begin VB.TextBox txtEquip 
      DataField       =   "Equip42"
      DataSource      =   "axsMember"
      Height          =   285
      Index           =   42
      Left            =   3840
      TabIndex        =   25
      Top             =   3960
      Width           =   255
   End
   Begin VB.TextBox txtEquip 
      DataField       =   "Equip43"
      DataSource      =   "axsMember"
      Height          =   285
      Index           =   43
      Left            =   4080
      TabIndex        =   24
      Top             =   3960
      Width           =   255
   End
   Begin VB.TextBox txtEquip 
      DataField       =   "Equip44"
      DataSource      =   "axsMember"
      Height          =   285
      Index           =   44
      Left            =   4320
      TabIndex        =   23
      Top             =   3960
      Width           =   255
   End
   Begin VB.TextBox txtClass 
      DataField       =   "Class"
      DataSource      =   "axsMember"
      Height          =   285
      Left            =   960
      TabIndex        =   22
      Top             =   2640
      Width           =   735
   End
   Begin VB.TextBox txtDex 
      DataField       =   "Dex"
      DataSource      =   "axsMember"
      Height          =   285
      Left            =   2400
      TabIndex        =   21
      Top             =   3000
      Width           =   735
   End
   Begin VB.TextBox txtSta 
      DataField       =   "Sta"
      DataSource      =   "axsMember"
      Height          =   285
      Left            =   2400
      TabIndex        =   20
      Top             =   3360
      Width           =   735
   End
   Begin VB.TextBox txtStr 
      DataField       =   "Str"
      DataSource      =   "axsMember"
      Height          =   285
      Left            =   2400
      TabIndex        =   19
      Top             =   2640
      Width           =   735
   End
   Begin VB.TextBox txtHp 
      DataField       =   "HP"
      DataSource      =   "axsMember"
      Height          =   285
      Left            =   960
      TabIndex        =   18
      Top             =   3000
      Width           =   735
   End
   Begin VB.TextBox txtMana 
      DataField       =   "Mana"
      DataSource      =   "axsMember"
      Height          =   285
      Left            =   960
      TabIndex        =   17
      Top             =   3360
      Width           =   735
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   2280
      TabIndex        =   16
      Top             =   4440
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   960
      TabIndex        =   15
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Data axsMember 
      Caption         =   "Members"
      Connect         =   "Access 2000;"
      DatabaseName    =   "members.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "members"
      Top             =   0
      Width           =   4695
   End
   Begin VB.TextBox txtName 
      DataField       =   "name"
      DataSource      =   "axsMember"
      Height          =   285
      Left            =   960
      TabIndex        =   6
      Top             =   480
      Width           =   2175
   End
   Begin VB.TextBox txtEmail 
      DataField       =   "EMail"
      DataSource      =   "axsMember"
      Height          =   285
      Left            =   960
      TabIndex        =   5
      Top             =   1200
      Width           =   2175
   End
   Begin VB.TextBox txtPassword 
      DataField       =   "password"
      DataSource      =   "axsMember"
      Height          =   285
      Left            =   960
      TabIndex        =   4
      Top             =   840
      Width           =   2175
   End
   Begin VB.TextBox txtBan 
      DataField       =   "ban"
      DataSource      =   "axsMember"
      Height          =   285
      Left            =   960
      TabIndex        =   3
      Top             =   2280
      Width           =   735
   End
   Begin VB.TextBox txtOffx 
      DataField       =   "offx"
      DataSource      =   "axsMember"
      Height          =   285
      Left            =   1440
      TabIndex        =   2
      Top             =   1560
      Width           =   615
   End
   Begin VB.TextBox txtOffy 
      DataField       =   "offy"
      DataSource      =   "axsMember"
      Height          =   285
      Left            =   2520
      TabIndex        =   1
      Top             =   1560
      Width           =   615
   End
   Begin VB.TextBox txtGold 
      DataField       =   "Gold"
      DataSource      =   "axsMember"
      Height          =   285
      Left            =   960
      TabIndex        =   0
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Message:"
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   81
      Top             =   4080
      Width           =   735
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Admin:"
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   80
      Top             =   3720
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "Bank:"
      Height          =   255
      Index           =   8
      Left            =   3480
      TabIndex        =   77
      Top             =   3000
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "Inventory:"
      Height          =   255
      Index           =   7
      Left            =   3360
      TabIndex        =   76
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label10 
      Caption         =   "Equiped:"
      Height          =   255
      Left            =   3360
      TabIndex        =   75
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Sta:"
      Height          =   255
      Index           =   6
      Left            =   1560
      TabIndex        =   74
      Top             =   3360
      Width           =   735
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Dex:"
      Height          =   255
      Index           =   5
      Left            =   1560
      TabIndex        =   73
      Top             =   3000
      Width           =   735
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Str:"
      Height          =   255
      Index           =   4
      Left            =   1560
      TabIndex        =   72
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Level:"
      Height          =   255
      Index           =   3
      Left            =   1560
      TabIndex        =   71
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Mana:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   70
      Top             =   3360
      Width           =   735
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "HP:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   69
      Top             =   3000
      Width           =   735
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "Class:"
      Height          =   255
      Left            =   120
      TabIndex        =   68
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label Label8 
      Caption         =   "Y"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2280
      TabIndex        =   14
      Top             =   1560
      Width           =   135
   End
   Begin VB.Label Label7 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   13
      Top             =   1560
      Width           =   135
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Ban:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   12
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Gold:"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Offset:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "E-Mail:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Password:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   480
      Width           =   735
   End
End
Attribute VB_Name = "frmEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDelete_Click()
axsMember.Recordset.Delete
End Sub

Private Sub cmdSave_Click()
axsMember.Recordset.Edit
axsMember.Recordset.Update

End Sub
