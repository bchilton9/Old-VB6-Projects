VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBuilder 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Deck Builder"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   10875
   Icon            =   "frmBuilder.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   455
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   725
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   2790
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuilder.frx":1272
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuilder.frx":1356
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuilder.frx":142A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuilder.frx":1503
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuilder.frx":15D1
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuilder.frx":16B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuilder.frx":1794
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuilder.frx":1873
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuilder.frx":194E
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuilder.frx":1A3D
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuilder.frx":1B4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuilder.frx":1C5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuilder.frx":1D63
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuilder.frx":1E7A
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuilder.frx":1F7F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlFrames 
      Left            =   2160
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   198
      ImageHeight     =   297
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuilder.frx":208D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuilder.frx":3C29
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuilder.frx":575F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cdlShow 
      Left            =   0
      Top             =   4320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer tmrArrow 
      Enabled         =   0   'False
      Interval        =   90
      Left            =   0
      Top             =   6240
   End
   Begin VB.Frame Frame2 
      Caption         =   "Deck"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   90
      TabIndex        =   3
      Top             =   0
      Width           =   10695
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1485
         Index           =   39
         Left            =   4125
         Top             =   3270
         Width           =   990
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1485
         Index           =   38
         Left            =   4620
         Top             =   3270
         Width           =   990
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1485
         Index           =   37
         Left            =   5115
         Top             =   3270
         Width           =   990
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1485
         Index           =   36
         Left            =   5610
         Top             =   3270
         Width           =   990
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1485
         Index           =   35
         Left            =   6105
         Top             =   3270
         Width           =   990
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1485
         Index           =   34
         Left            =   6600
         Top             =   3270
         Width           =   990
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1485
         Index           =   33
         Left            =   7095
         Top             =   3270
         Width           =   990
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1485
         Index           =   32
         Left            =   7590
         Top             =   3270
         Width           =   990
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1485
         Index           =   31
         Left            =   8085
         Top             =   3270
         Width           =   990
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1485
         Index           =   30
         Left            =   8580
         Top             =   3270
         Width           =   990
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1485
         Index           =   29
         Left            =   9075
         Top             =   3270
         Width           =   990
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1485
         Index           =   27
         Left            =   3135
         Top             =   1755
         Width           =   990
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1485
         Index           =   26
         Left            =   3630
         Top             =   1755
         Width           =   990
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1485
         Index           =   25
         Left            =   4125
         Top             =   1755
         Width           =   990
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1485
         Index           =   24
         Left            =   4620
         Top             =   1755
         Width           =   990
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1485
         Index           =   23
         Left            =   5115
         Top             =   1755
         Width           =   990
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1485
         Index           =   22
         Left            =   5610
         Top             =   1755
         Width           =   990
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1485
         Index           =   21
         Left            =   6105
         Top             =   1755
         Width           =   990
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1485
         Index           =   20
         Left            =   6600
         Top             =   1755
         Width           =   990
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1485
         Index           =   19
         Left            =   7095
         Top             =   1755
         Width           =   990
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1485
         Index           =   18
         Left            =   7590
         Top             =   1755
         Width           =   990
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1485
         Index           =   17
         Left            =   8085
         Top             =   1755
         Width           =   990
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1485
         Index           =   16
         Left            =   8580
         Top             =   1755
         Width           =   990
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1485
         Index           =   15
         Left            =   9075
         Top             =   1755
         Width           =   990
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1485
         Index           =   13
         Left            =   3135
         Top             =   240
         Width           =   990
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1485
         Index           =   12
         Left            =   3630
         Top             =   240
         Width           =   990
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1485
         Index           =   11
         Left            =   4125
         Top             =   240
         Width           =   990
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1485
         Index           =   10
         Left            =   4620
         Top             =   240
         Width           =   990
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1485
         Index           =   9
         Left            =   5115
         Top             =   240
         Width           =   990
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1485
         Index           =   8
         Left            =   5610
         Top             =   240
         Width           =   990
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1485
         Index           =   7
         Left            =   6105
         Top             =   240
         Width           =   990
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1485
         Index           =   6
         Left            =   6600
         Top             =   240
         Width           =   990
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1485
         Index           =   5
         Left            =   7095
         Top             =   240
         Width           =   990
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1485
         Index           =   4
         Left            =   7590
         Top             =   240
         Width           =   990
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1485
         Index           =   3
         Left            =   8085
         Top             =   240
         Width           =   990
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1485
         Index           =   2
         Left            =   8580
         Top             =   240
         Width           =   990
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1485
         Index           =   1
         Left            =   9075
         Top             =   240
         Width           =   990
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1485
         Index           =   0
         Left            =   9570
         Top             =   240
         Width           =   990
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1485
         Index           =   28
         Left            =   9570
         Top             =   3270
         Width           =   990
      End
      Begin VB.Image imgCard 
         Appearance      =   0  'Flat
         Height          =   1485
         Index           =   14
         Left            =   9570
         Top             =   1755
         Width           =   990
      End
      Begin VB.Label lblData 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   5
         Left            =   270
         TabIndex        =   9
         Top             =   3390
         Width           =   1350
      End
      Begin VB.Label lblData 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   4
         Left            =   1965
         TabIndex        =   8
         Top             =   4380
         Width           =   840
      End
      Begin VB.Image imgStar 
         Height          =   210
         Index           =   9
         Left            =   450
         Top             =   750
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.Image imgStar 
         Height          =   210
         Index           =   8
         Left            =   675
         Top             =   750
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.Image imgStar 
         Height          =   210
         Index           =   7
         Left            =   900
         Top             =   750
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.Image imgStar 
         Height          =   210
         Index           =   6
         Left            =   1125
         Top             =   750
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.Image imgStar 
         Height          =   210
         Index           =   5
         Left            =   1350
         Top             =   750
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.Image imgStar 
         Height          =   210
         Index           =   4
         Left            =   1575
         Top             =   750
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.Image imgStar 
         Height          =   210
         Index           =   3
         Left            =   1800
         Top             =   750
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.Image imgStar 
         Height          =   210
         Index           =   2
         Left            =   2025
         Top             =   750
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.Image imgStar 
         Height          =   210
         Index           =   1
         Left            =   2250
         Top             =   750
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.Image imgStar 
         Height          =   210
         Index           =   0
         Left            =   2475
         Top             =   750
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.Image imgIcon 
         Height          =   240
         Left            =   2610
         Top             =   450
         Width           =   240
      End
      Begin VB.Label lblData 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   3
         Left            =   1275
         TabIndex        =   7
         Top             =   4380
         Width           =   840
      End
      Begin VB.Label lblData 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   1
         Left            =   1410
         TabIndex        =   6
         Top             =   735
         Width           =   1350
      End
      Begin VB.Label lblData 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   825
         Index           =   2
         Left            =   270
         TabIndex        =   5
         Top             =   3525
         Width           =   2610
      End
      Begin VB.Label lblData 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   0
         Left            =   315
         TabIndex        =   4
         Top             =   480
         Width           =   2340
      End
      Begin VB.Image imgMain 
         Height          =   2130
         Left            =   510
         Top             =   1080
         Width           =   2130
      End
      Begin VB.Image imgFrame 
         Appearance      =   0  'Flat
         Height          =   4455
         Left            =   90
         Stretch         =   -1  'True
         Top             =   255
         Width           =   2970
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Card List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   90
      TabIndex        =   0
      Top             =   4950
      Width           =   10695
      Begin VB.Label lblArrow 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "       >"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1200
         Index           =   1
         Left            =   10440
         TabIndex        =   2
         Top             =   360
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Label lblArrow 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "       <"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1200
         Index           =   0
         Left            =   0
         TabIndex        =   1
         Top             =   360
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Image img 
         Appearance      =   0  'Flat
         Height          =   1485
         Index           =   1
         Left            =   255
         Stretch         =   -1  'True
         Top             =   225
         Width           =   990
      End
      Begin VB.Image img 
         Appearance      =   0  'Flat
         Height          =   1485
         Index           =   2
         Left            =   1275
         Stretch         =   -1  'True
         Top             =   225
         Width           =   990
      End
      Begin VB.Image img 
         Appearance      =   0  'Flat
         Height          =   1485
         Index           =   3
         Left            =   2295
         Stretch         =   -1  'True
         Top             =   225
         Width           =   990
      End
      Begin VB.Image img 
         Appearance      =   0  'Flat
         Height          =   1485
         Index           =   4
         Left            =   3315
         Stretch         =   -1  'True
         Top             =   225
         Width           =   990
      End
      Begin VB.Image img 
         Appearance      =   0  'Flat
         Height          =   1485
         Index           =   5
         Left            =   4335
         Stretch         =   -1  'True
         Top             =   225
         Width           =   990
      End
      Begin VB.Image img 
         Appearance      =   0  'Flat
         Height          =   1485
         Index           =   6
         Left            =   5355
         Stretch         =   -1  'True
         Top             =   225
         Width           =   990
      End
      Begin VB.Image img 
         Appearance      =   0  'Flat
         Height          =   1485
         Index           =   7
         Left            =   6375
         Stretch         =   -1  'True
         Top             =   225
         Width           =   990
      End
      Begin VB.Image img 
         Appearance      =   0  'Flat
         Height          =   1485
         Index           =   8
         Left            =   7395
         Stretch         =   -1  'True
         Top             =   225
         Width           =   990
      End
      Begin VB.Image img 
         Appearance      =   0  'Flat
         Height          =   1485
         Index           =   9
         Left            =   8415
         Stretch         =   -1  'True
         Top             =   225
         Width           =   990
      End
      Begin VB.Image img 
         Appearance      =   0  'Flat
         Height          =   1485
         Index           =   10
         Left            =   9435
         Stretch         =   -1  'True
         Top             =   225
         Width           =   990
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&File"
      Begin VB.Menu mnuMFile 
         Caption         =   "&Open Deck"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuMSave 
         Caption         =   "&Save Deck"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMExit 
         Caption         =   "&Quit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHAbout 
         Caption         =   "&About"
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "frmBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private iInc As Integer
Private iList As Integer
Private iStart As Integer
Private strFile As String
Private Type Card
  Frame As Integer
  Name As String
  Attribute As Integer
  Icon As String
  Type As String
  Description As String
  Level As Integer
  Cost As Integer
  Attack As Integer
  Defence As Integer
  Phase As String
  Spell As String
  Value As String
  Value2 As String
End Type

Private Sub Copy_Card(objCard As Object, objWith As Object)
  objWith.Picture = objCard.Picture
  objWith.Tag = objCard.Tag
  objWith.Visible = True
End Sub

Private Sub img_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  Show_Preview img(Index).Tag
End Sub

Private Sub img_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim iZum As Integer
  
  For i% = 0 To 39
    If imgCard(i%).Tag = img(Index).Tag Then iZum = iZum + 1
    If iZum > 2 Then MsgBox "You already have 3 duplicates of " & Get_Card(img(Index).Tag).Name & ".": Exit For
    If imgCard(i%).Tag = "" Then Copy_Card img(Index), imgCard(i%): Exit For
    If i% = 39 Then MsgBox "Deck complete."
  Next i%
End Sub

Private Sub lblArrow_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 lblArrow(Index).ForeColor = vbMagenta
 iInc = IIf(Index = 0, -1, 1)
 tmrArrow.Enabled = True
End Sub

Private Sub lblArrow_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 lblArrow(Index).ForeColor = vbBlack
 tmrArrow.Enabled = False
End Sub

Private Sub mnuHAbout_Click()
  MsgBox "Monsters Duel Deck Builder" & vbCrLf & "by Abel Antonio Ricaurte J"
End Sub

Private Sub mnuMSave_Click()
Dim strCards As String

  If imgCard(38).Tag = "" Or imgCard(39).Tag = "" Then MsgBox "A Deck must contain 40 cards.": Exit Sub
  For i% = 0 To imgCard.Count - 1
    strCards = strCards & "|" & imgCard(i%).Tag
  Next i%
  WriteINI "Deck", "Cards", iList & strCards, strFile
End Sub

Private Sub tmrArrow_Timer()
  iStart = iStart + iInc
  Load_List
  lblArrow(1).Visible = IIf(iStart > iList, False, True)
  lblArrow(0).Visible = IIf(iStart < 1, False, True)
  If iStart > iList Or iStart < 1 Then tmrArrow.Enabled = False
End Sub

Private Sub imgCard_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  imgCard(Index).Visible = False
  imgCard(Index).Tag = ""
  Rearrange_Cards
End Sub

Private Sub Rearrange_Cards()
  For i% = imgCard.Count - 2 To 0 Step -1
    If imgCard(i%).Tag = "" And imgCard(i% + 1).Tag <> "" Then Replace_Card imgCard(i%), imgCard(i% + 1): i% = imgCard.Count - 2
  Next i%
End Sub

Private Sub Replace_Card(objCard As Object, objWith As Object)
  objCard.Picture = objWith.Picture
  objCard.Tag = objWith.Tag
  objWith.Visible = False
  objCard.Visible = True
  objWith.Tag = ""
End Sub

Private Sub mnuMFile_Click()
On Error Resume Next
  cdlShow.Filter = "Deck Files (*.dek)|*.dek"
  cdlShow.InitDir = App.Path
  cdlShow.ShowOpen
  If Err = 0 Then strFile = cdlShow.filename: iStart = 0: Me.Caption = "Deck Builder - " & cdlShow.FileTitle: Load_Deck: Load_List
End Sub

Private Sub imgCard_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  Show_Preview imgCard(Index).Tag
End Sub

Private Function Get_Card(strCard As String) As Card
Dim varData As Variant
  
  varData = Split(ReadINI("Cards", strCard, App.Path & "\images\set.mds"), "|")
  
  With Get_Card
    .Frame = varData(0)
    .Name = varData(1)
    .Attribute = varData(2)
    .Icon = varData(3)
    .Type = varData(4)
    .Description = varData(5)
    If varData(6) = "" Then .Level = 0 Else .Level = varData(6)
    If varData(7) = "" Then .Cost = 0 Else .Cost = varData(7)
    If varData(8) = "" Then .Attack = 0 Else .Attack = varData(8)
    If varData(9) = "" Then .Defence = 0 Else .Defence = varData(9)
    .Phase = varData(10)
    .Spell = varData(11)
    .Value = varData(12)
    .Value2 = varData(13)
  End With
End Function

Private Sub Show_Preview(strCard As String, Optional strCase As String)
Dim crdCard As Card

  If strCard = imgMain.Tag Or strCard = "" Then Exit Sub
  crdCard = Get_Card(strCard)
  imgMain.Tag = strCard
  lblData(0) = crdCard.Name
  lblData(1) = IIf(crdCard.Frame > 2, "", "[ " & crdCard.Type & " Card" & IIf(crdCard.Icon <> "", "      ]", " ]"))
  lblData(2) = crdCard.Description
  lblData(3) = IIf(crdCard.Frame > 2, "ATK/ " & crdCard.Attack, "")
  lblData(4) = IIf(crdCard.Frame > 2, "DEF/ " & crdCard.Defence, "")
  lblData(2).Height = IIf(crdCard.Frame > 2, 55 * 15, 72 * 15)
  lblData(2).Top = IIf(crdCard.Frame > 2, 234 * 15, 224 * 15) + 15
  lblData(5) = IIf(crdCard.Frame > 2, "[" & crdCard.Type & "]", "")
  
  For i% = 0 To 9: imgStar(i%).Visible = IIf(crdCard.Level > i%, True, False): imgStar(i%).Picture = imlIcons.ListImages(9).Picture: Next i%
  If crdCard.Icon <> "" Then imgStar(0).Picture = imlIcons.ListImages(Switch(crdCard.Icon = "Continuous", 10, crdCard.Icon = "Counter", 11, crdCard.Icon = "Equip", 12, crdCard.Icon = "Field", 13, crdCard.Icon = "Quick", 14, crdCard.Icon = "Ritual", 15)).Picture: imgStar(0).Visible = True
  imgIcon.Picture = imlIcons.ListImages(crdCard.Attribute).Picture
  imgFrame.Picture = imlFrames.ListImages(crdCard.Frame).Picture
  imgMain.Picture = LoadPicture(App.Path & "\images\" & crdCard.Name & ".jpg")
End Sub

Private Sub Load_Deck()
Dim varData As Variant

  varData = Split(ReadINI("Deck", "Cards", strFile), "|")
  
  iList = varData(0)
  For i% = 0 To 39
    imgCard(i%).Tag = varData(i% + 1)
    imgCard(i%).Picture = LoadPicture(App.Path & "\images\Atk_" & Get_Card(imgCard(i%).Tag).Name & ".jpg")
    imgCard(i%).Visible = True
  Next i%
End Sub

Private Sub Load_List()
  If iStart = 0 Then lblArrow(0).Visible = False: lblArrow(1).Visible = True
  For i% = 1 To 10
    img(i%).Tag = iStart + i%
    img(i%).Picture = LoadPicture(App.Path & "\images\Atk_" & Get_Card(img(i%).Tag).Name & ".jpg")
  Next i%
End Sub

Private Sub mnuMExit_Click()
  End
End Sub
