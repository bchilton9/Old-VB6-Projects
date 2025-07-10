VERSION 5.00
Begin VB.Form frmItems 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Items"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Types"
      Height          =   1815
      Left            =   3720
      TabIndex        =   23
      Top             =   120
      Width           =   1455
      Begin VB.Label Label16 
         Caption         =   "5: Boots"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label15 
         Caption         =   "4: Gloves"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label14 
         Caption         =   "3: Pants"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label11 
         Caption         =   "0: Misc"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label12 
         Caption         =   "1: Weapon"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label13 
         Caption         =   "2: Chest"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   720
         Width           =   1215
      End
   End
   Begin VB.TextBox txtFormItemType 
      DataField       =   "type"
      DataSource      =   "dataItems"
      Height          =   285
      Left            =   1320
      TabIndex        =   21
      Top             =   2880
      Width           =   495
   End
   Begin VB.TextBox txtItemDmg 
      DataField       =   "Dmg"
      DataSource      =   "dataItems"
      Height          =   285
      Left            =   2760
      TabIndex        =   17
      Top             =   2520
      Width           =   495
   End
   Begin VB.TextBox txtFormItemEquip 
      DataField       =   "Equip"
      DataSource      =   "dataItems"
      Height          =   285
      Left            =   1320
      TabIndex        =   15
      Top             =   2520
      Width           =   495
   End
   Begin VB.TextBox txtFormItemHP 
      DataField       =   "HP"
      DataSource      =   "dataItems"
      Height          =   285
      Left            =   1320
      TabIndex        =   12
      Top             =   2160
      Width           =   495
   End
   Begin VB.TextBox txtFormItemMana 
      DataField       =   "Mana"
      DataSource      =   "dataItems"
      Height          =   285
      Left            =   2760
      TabIndex        =   11
      Top             =   2160
      Width           =   495
   End
   Begin VB.TextBox txtFormItemSell 
      DataField       =   "sell"
      DataSource      =   "dataItems"
      Height          =   285
      Left            =   2760
      TabIndex        =   10
      Top             =   1800
      Width           =   495
   End
   Begin VB.TextBox txtFormItemBuy 
      DataField       =   "buy"
      DataSource      =   "dataItems"
      Height          =   285
      Left            =   1320
      TabIndex        =   9
      Top             =   1800
      Width           =   495
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   3240
      TabIndex        =   5
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      Top             =   3360
      Width           =   975
   End
   Begin VB.TextBox txtFormItemDesc 
      DataField       =   "Description"
      DataSource      =   "dataItems"
      Height          =   285
      Left            =   1320
      TabIndex        =   3
      Top             =   1320
      Width           =   2175
   End
   Begin VB.TextBox txtFormItemName 
      DataField       =   "Name"
      DataSource      =   "dataItems"
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Top             =   960
      Width           =   2175
   End
   Begin VB.Data dataItems 
      Caption         =   "Items"
      Connect         =   "Access"
      DatabaseName    =   "F:\vbprojects\Mesto\mesto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Items"
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "Type:"
      Height          =   255
      Left            =   360
      TabIndex        =   22
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label lblIDnum 
      DataField       =   "ID"
      DataSource      =   "dataItems"
      Height          =   255
      Left            =   1920
      TabIndex        =   20
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "Item ID Number:"
      Height          =   255
      Left            =   480
      TabIndex        =   19
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "Damage:"
      Height          =   255
      Left            =   1800
      TabIndex        =   18
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Equipable:"
      Height          =   255
      Left            =   360
      TabIndex        =   16
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "HP:"
      Height          =   255
      Left            =   360
      TabIndex        =   14
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Mana:"
      Height          =   255
      Left            =   1920
      TabIndex        =   13
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Sell Price:"
      Height          =   255
      Left            =   1920
      TabIndex        =   8
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Buy Price:"
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Item Desc:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Item Name:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   855
   End
End
Attribute VB_Name = "frmItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdEdit_Click()
dataItems.Recordset.Edit
End Sub

Private Sub cmdNew_Click()
dataItems.Recordset.AddNew
End Sub

Private Sub cmdSave_Click()
On Error GoTo error
dataItems.Recordset.Update
Exit Sub
error:
Msg = MsgBox("Unable to add/edit item!", vbCritical)
End Sub
