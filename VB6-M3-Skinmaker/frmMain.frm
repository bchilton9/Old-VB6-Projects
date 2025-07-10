VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "M3 DS Simply - Skin Maker"
   ClientHeight    =   9345
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   21570
   LinkTopic       =   "Form1"
   ScaleHeight     =   9345
   ScaleWidth      =   21570
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Caption         =   "bckgrd_2.bmp"
      Height          =   4095
      Left            =   4560
      TabIndex        =   16
      Top             =   4320
      Width           =   4215
      Begin VB.PictureBox picBG2 
         Height          =   2940
         Left            =   120
         ScaleHeight     =   2880
         ScaleWidth      =   3840
         TabIndex        =   18
         Top             =   240
         Width           =   3900
      End
      Begin VB.CommandButton cmdBG2 
         Caption         =   "Background 2"
         Height          =   495
         Left            =   120
         TabIndex        =   17
         Top             =   3240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "icons.bmp"
      Height          =   4095
      Left            =   120
      TabIndex        =   13
      Top             =   4320
      Width           =   4215
      Begin VB.PictureBox picIcon 
         Height          =   2940
         Left            =   120
         ScaleHeight     =   2880
         ScaleWidth      =   3840
         TabIndex        =   15
         Top             =   240
         Width           =   3900
      End
      Begin VB.CommandButton cmdIconBG 
         Caption         =   "Icon Background"
         Height          =   495
         Left            =   120
         TabIndex        =   14
         Top             =   3360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "bckgrd_1.bmp"
      Height          =   3975
      Left            =   4560
      TabIndex        =   10
      Top             =   120
      Width           =   4215
      Begin VB.CheckBox chkSGbg1 
         Caption         =   "Show Grid"
         Height          =   255
         Left            =   1920
         TabIndex        =   19
         Top             =   3240
         Width           =   1575
      End
      Begin VB.CommandButton cmbBG1 
         Caption         =   "Background 1"
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   3240
         Width           =   1695
      End
      Begin VB.PictureBox picBG1 
         Height          =   2940
         Left            =   120
         ScaleHeight     =   2880
         ScaleWidth      =   3840
         TabIndex        =   11
         Top             =   240
         Width           =   3900
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "logo.bmp"
      Height          =   3975
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   4215
      Begin VB.CommandButton cmdLogo 
         Caption         =   "Logo Image"
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   3240
         Width           =   1695
      End
      Begin VB.PictureBox picLogo 
         Height          =   2940
         Left            =   120
         ScaleHeight     =   2880
         ScaleWidth      =   3840
         TabIndex        =   8
         Top             =   240
         Width           =   3900
      End
   End
   Begin MSComDlg.CommonDialog bg2img 
      Left            =   9000
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog iconbg 
      Left            =   9000
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picBG2temp 
      Height          =   2940
      Left            =   13560
      ScaleHeight     =   2880
      ScaleWidth      =   3840
      TabIndex        =   6
      Top             =   3240
      Width           =   3900
   End
   Begin VB.PictureBox picIconTemp 
      Height          =   2940
      Left            =   9600
      ScaleHeight     =   2880
      ScaleWidth      =   3840
      TabIndex        =   5
      Top             =   3240
      Width           =   3900
   End
   Begin VB.PictureBox picBG1Temp 
      Height          =   2940
      Left            =   13560
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   2880
      ScaleWidth      =   3840
      TabIndex        =   4
      Top             =   240
      Width           =   3900
   End
   Begin MSComDlg.CommonDialog bg1img 
      Left            =   9000
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picBG1GridMask 
      Height          =   2940
      Left            =   13560
      Picture         =   "frmMain.frx":24044
      ScaleHeight     =   2880
      ScaleWidth      =   3840
      TabIndex        =   3
      Top             =   6240
      Width           =   3900
   End
   Begin VB.PictureBox picBG1Grid 
      Height          =   2940
      Left            =   9600
      Picture         =   "frmMain.frx":48088
      ScaleHeight     =   2880
      ScaleWidth      =   3840
      TabIndex        =   2
      Top             =   6240
      Width           =   3900
   End
   Begin VB.PictureBox PicNoMask 
      Height          =   2940
      Left            =   17520
      Picture         =   "frmMain.frx":6C0CC
      ScaleHeight     =   2880
      ScaleWidth      =   3840
      TabIndex        =   1
      Top             =   6240
      Width           =   3900
   End
   Begin VB.PictureBox picLogoTemp 
      Height          =   2940
      Left            =   9600
      Picture         =   "frmMain.frx":90110
      ScaleHeight     =   2880
      ScaleWidth      =   3840
      TabIndex        =   0
      Top             =   240
      Width           =   3900
   End
   Begin MSComDlg.CommonDialog logoimg 
      Left            =   9000
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkSGbg1_Click()
  BitBlt picBG1.hdc, 0, 0, 256, 192, PicNoMask.hdc, 0, 0, vbMergePaint
  BitBlt picBG1.hdc, 0, 0, 256, 192, picBG1Temp.hdc, 0, 0, vbSrcAnd

If chkSGbg1.Value = Checked Then
  BitBlt picBG1.hdc, 0, 0, 256, 192, picBG1GridMask.hdc, 0, 0, vbMergePaint
  BitBlt picBG1.hdc, 0, 0, 256, 192, picBG1Grid.hdc, 0, 0, vbSrcAnd
End If
End Sub

Private Sub cmbBG1_Click()

bg2img.Filter = "BMP files (*.bmp)|*.bmp"
bg2img.ShowOpen

picBG1Temp.Picture = LoadPicture(bg2img.FileName)

  BitBlt picBG1.hdc, 0, 0, 256, 192, PicNoMask.hdc, 0, 0, vbMergePaint
  BitBlt picBG1.hdc, 0, 0, 256, 192, picBG1Temp.hdc, 0, 0, vbSrcAnd

If chkSGbg1.Value = Checked Then
  BitBlt picBG1.hdc, 0, 0, 256, 192, picBG1GridMask.hdc, 0, 0, vbMergePaint
  BitBlt picBG1.hdc, 0, 0, 256, 192, picBG1Grid.hdc, 0, 0, vbSrcAnd
End If

End Sub

Private Sub cmdBG2_Click()

bg1img.Filter = "BMP files (*.bmp)|*.bmp"
bg1img.ShowOpen

picBG2temp.Picture = LoadPicture(bg1img.FileName)

  BitBlt picBG2.hdc, 0, 0, 256, 192, PicNoMask.hdc, 0, 0, vbMergePaint
  BitBlt picBG2.hdc, 0, 0, 256, 192, picBG2temp.hdc, 0, 0, vbSrcAnd

End Sub

Private Sub cmdIconBG_Click()

iconbg.Filter = "BMP files (*.bmp)|*.bmp"
iconbg.ShowOpen

picIconTemp.Picture = LoadPicture(iconbg.FileName)


  BitBlt picIcon.hdc, 0, 0, 256, 192, PicNoMask.hdc, 0, 0, vbMergePaint
  BitBlt picIcon.hdc, 0, 0, 256, 192, picIconTemp.hdc, 0, 0, vbSrcAnd


End Sub

Private Sub cmdLogo_Click()

logoimg.Filter = "BMP files (*.bmp)|*.bmp"
logoimg.ShowOpen

picLogoTemp.Picture = LoadPicture(logoimg.FileName)


  BitBlt picLogo.hdc, 0, 0, 256, 192, PicNoMask.hdc, 0, 0, vbMergePaint
  BitBlt picLogo.hdc, 0, 0, 256, 192, picLogoTemp.hdc, 0, 0, vbSrcAnd


End Sub
