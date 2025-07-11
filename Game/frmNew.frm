VERSION 5.00
Begin VB.Form frmNew 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "The Land's of Tidel - New Character"
   ClientHeight    =   4860
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7020
   Icon            =   "frmNew.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmNew.frx":08CA
   ScaleHeight     =   4860
   ScaleWidth      =   7020
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox cmdCancel3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      Picture         =   "frmNew.frx":82CD
      ScaleHeight     =   555
      ScaleWidth      =   1620
      TabIndex        =   16
      Top             =   2880
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.PictureBox cmdCreate3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      Picture         =   "frmNew.frx":AF01
      ScaleHeight     =   555
      ScaleWidth      =   1620
      TabIndex        =   15
      Top             =   1440
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.PictureBox cmdCancel2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      Picture         =   "frmNew.frx":DC19
      ScaleHeight     =   555
      ScaleWidth      =   1620
      TabIndex        =   14
      Top             =   2160
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.PictureBox cmdCreate2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      Picture         =   "frmNew.frx":105A9
      ScaleHeight     =   555
      ScaleWidth      =   1620
      TabIndex        =   13
      Top             =   840
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.PictureBox cmdCancel 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   4200
      Picture         =   "frmNew.frx":12F95
      ScaleHeight     =   555
      ScaleWidth      =   1620
      TabIndex        =   12
      Top             =   3840
      Width           =   1620
   End
   Begin VB.PictureBox cmdCreate 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   2400
      Picture         =   "frmNew.frx":15BC9
      ScaleHeight     =   555
      ScaleWidth      =   1620
      TabIndex        =   11
      Top             =   3840
      Width           =   1620
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   495
      Left            =   5880
      TabIndex        =   9
      Top             =   3000
      Width           =   615
      Begin VB.PictureBox picSpite 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   120
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   10
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.Timer tmrSpite 
      Interval        =   500
      Left            =   120
      Top             =   0
   End
   Begin VB.TextBox txtSet 
      Height          =   375
      Left            =   2400
      TabIndex        =   8
      Top             =   6000
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtOffy 
      Height          =   375
      Left            =   1800
      TabIndex        =   7
      Top             =   6000
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtOffx 
      Height          =   375
      Left            =   1200
      TabIndex        =   6
      Top             =   6000
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtMask 
      Height          =   375
      Left            =   600
      TabIndex        =   5
      Text            =   "0"
      Top             =   6000
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtPassB 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   3000
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1920
      Width           =   2415
   End
   Begin VB.ComboBox cmbMask 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmNew.frx":188E1
      Left            =   4080
      List            =   "frmNew.frx":1891B
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   3120
      Width           =   1575
   End
   Begin VB.TextBox txtEmail 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Top             =   2520
      Width           =   2415
   End
   Begin VB.TextBox txtPass 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2400
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1320
      Width           =   2415
   End
   Begin VB.TextBox txtUser 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   840
      Width           =   2415
   End
End
Attribute VB_Name = "frmNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private ontop As New clsOnTop
Dim step As Integer


Private Sub cmbMask_Click()
If cmbMask.ListIndex = 1 Then
txtOffx.text = 0
txtOffy.text = 0
txtSet.text = 0
ElseIf cmbMask.ListIndex = 2 Then
txtOffx.text = 0
txtOffy.text = 128
txtSet.text = 0
ElseIf cmbMask.ListIndex = 0 Then
txtOffx.text = 73
txtOffy.text = 0
txtSet.text = 0
ElseIf cmbMask.ListIndex = 3 Then
txtOffx.text = 73
txtOffy.text = 128
txtSet.text = 0
ElseIf cmbMask.ListIndex = 4 Then
txtOffx.text = 146
txtOffy.text = 0
txtSet.text = 0
ElseIf cmbMask.ListIndex = 5 Then
txtOffx.text = 146
txtOffy.text = 128
txtSet.text = 0
ElseIf cmbMask.ListIndex = 6 Then
txtOffx.text = 219
txtOffy.text = 0
txtSet.text = 0
ElseIf cmbMask.ListIndex = 7 Then
txtOffx.text = 219
txtOffy.text = 128
txtSet.text = 0

'next set
ElseIf cmbMask.ListIndex = 8 Then
txtOffx.text = 0
txtOffy.text = 250
txtSet.text = 0
ElseIf cmbMask.ListIndex = 9 Then
txtOffx.text = 1
txtOffy.text = 378
txtSet.text = 0
ElseIf cmbMask.ListIndex = 10 Then
txtOffx.text = 73
txtOffy.text = 250
txtSet.text = 0
ElseIf cmbMask.ListIndex = 11 Then
txtOffx.text = 73
txtOffy.text = 378
txtSet.text = 0
ElseIf cmbMask.ListIndex = 12 Then
txtOffx.text = 146
txtOffy.text = 251
txtSet.text = 0
ElseIf cmbMask.ListIndex = 13 Then
txtOffx.text = 146
txtOffy.text = 378
txtSet.text = 0
ElseIf cmbMask.ListIndex = 14 Then
txtOffx.text = 219
txtOffy.text = 251
txtSet.text = 0
ElseIf cmbMask.ListIndex = 15 Then
txtOffx.text = 219
txtOffy.text = 378
txtSet.text = 0
End If

txtMask.text = 1

Call BitBlt(picSpite.hdc, 0, 0, 22, 32, tiles.picChar.hdc, txtOffx.text, txtOffy.text, SRCCOPY)

End Sub

Private Sub cmdCancel_Click()

    frmLogin.Show
    Unload frmTCP
    Unload frmNew

End Sub

Private Sub cmdCreate_Click()

'Open (App.Path & "\host.dat") For Input As #1
'Input #1, serv
'Input #1, porta
'Input #1, portb
'Close #1

If txtUser.text = "" Then GoTo nouser
If txtPass.text = "" Then GoTo nopass
If txtPassB.text = "" Then GoTo nopassb
If txtEmail.text = "" Then GoTo noemail
If txtMask.text = 0 Then GoTo nospite
If txtPass.text <> txtPassB.text Then GoTo nomatch

isConnectedTryed = True
Load frmTCP
frmTCP.Show
Me.Hide
frmTCP.wskServer.RemotePort = portb
frmTCP.wskServer.Connect
Exit Sub

nouser:
msg = MsgBox("Create User Error: Please Enter a username!", vbCritical, "Create User Error.")
Exit Sub
nopass:
msg = MsgBox("Create User Error: Please Enter a password!", vbCritical, "Create User Error.")
Exit Sub
nopassb:
msg = MsgBox("Create User Error: Please Re-Type Password!", vbCritical, "Create User Error.")
Exit Sub
noemail:
msg = MsgBox("Create User Error: Please Enter a E-Mail Address!", vbCritical, "Create User Error.")
Exit Sub
nospite:
msg = MsgBox("Create User Error: Please chose a charator!", vbCritical, "Create User Error.")
Exit Sub
nomatch:
msg = MsgBox("Create User Error: Passwords do not match!", vbCritical, "Create User Error.")
Exit Sub
End Sub

Private Sub Form_Load()

  Set ontop = New clsOnTop
  ontop.MakeTopMost hWnd

cmbMask.ListIndex = 0
Load tiles
'tiles.Show
step = 7
End Sub

Private Sub Form_Unload(Cancel As Integer)

    frmLogin.Show
    Unload frmTCP
    Unload frmNew

End Sub

Private Sub tmrSpite_Timer()
If txtMask.text <> 0 Then
step = step + 1

If step = 1 Then
pushx = 0
pushy = 0
ElseIf step = 2 Then
pushx = 23
pushy = 0
ElseIf step = 3 Then
pushx = 46
pushy = 0
ElseIf step = 4 Then
pushx = 0
pushy = 32
ElseIf step = 5 Then
pushx = 23
pushy = 32
ElseIf step = 6 Then
pushx = 46
pushy = 32
ElseIf step = 7 Then
pushx = 0
pushy = 64
ElseIf step = 8 Then
pushx = 23
pushy = 64
ElseIf step = 9 Then
pushx = 46
pushy = 64
ElseIf step = 10 Then
pushx = 0
pushy = 96
ElseIf step = 11 Then
pushx = 23
pushy = 96
ElseIf step = 12 Then
pushx = 46
pushy = 96

step = 0
End If

Call BitBlt(picSpite.hdc, 0, 0, 23, 32, tiles.picChar.hdc, txtOffx.text + pushx, txtOffy.text + pushy, SRCCOPY)


End If
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdCreate = cmdCreate3
cmdCancel = cmdCancel3
End Sub

Private Sub cmdCreate_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdCreate = cmdCreate2
End Sub

Private Sub cmdCancel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdCancel = cmdCancel2
End Sub
