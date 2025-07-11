VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "The Land's of Tidel - Login"
   ClientHeight    =   4860
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7020
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLogin.frx":08CA
   ScaleHeight     =   4860
   ScaleWidth      =   7020
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkLocal 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Connect to LocalHost"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4200
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   15
      Top             =   4560
      Width           =   2295
   End
   Begin VB.PictureBox cmdDelete3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   3960
      Picture         =   "frmLogin.frx":79EB
      ScaleHeight     =   555
      ScaleWidth      =   2505
      TabIndex        =   14
      Top             =   600
      Visible         =   0   'False
      Width           =   2505
   End
   Begin VB.PictureBox cmdEdit3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   2640
      Picture         =   "frmLogin.frx":AC8D
      ScaleHeight     =   555
      ScaleWidth      =   2370
      TabIndex        =   13
      Top             =   600
      Visible         =   0   'False
      Width           =   2370
   End
   Begin VB.PictureBox cmdNew3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   1320
      Picture         =   "frmLogin.frx":DF92
      ScaleHeight     =   555
      ScaleWidth      =   2325
      TabIndex        =   12
      Top             =   600
      Visible         =   0   'False
      Width           =   2325
   End
   Begin VB.PictureBox cmdConnect3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   120
      Picture         =   "frmLogin.frx":1124D
      ScaleHeight     =   555
      ScaleWidth      =   1620
      TabIndex        =   11
      Top             =   600
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.PictureBox cmdDelete2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   3960
      Picture         =   "frmLogin.frx":13FFF
      ScaleHeight     =   555
      ScaleWidth      =   2505
      TabIndex        =   10
      Top             =   120
      Visible         =   0   'False
      Width           =   2505
   End
   Begin VB.PictureBox cmdEdit2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   2640
      Picture         =   "frmLogin.frx":16CA6
      ScaleHeight     =   555
      ScaleWidth      =   2370
      TabIndex        =   9
      Top             =   120
      Visible         =   0   'False
      Width           =   2370
   End
   Begin VB.PictureBox cmdNew2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   1320
      Picture         =   "frmLogin.frx":199C3
      ScaleHeight     =   555
      ScaleWidth      =   2325
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   2325
   End
   Begin VB.PictureBox cmdConnect2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   120
      Picture         =   "frmLogin.frx":1C6B9
      ScaleHeight     =   555
      ScaleWidth      =   1620
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.PictureBox cmdDelete 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   1440
      Picture         =   "frmLogin.frx":1F11F
      ScaleHeight     =   555
      ScaleWidth      =   2505
      TabIndex        =   6
      Top             =   3600
      Width           =   2505
   End
   Begin VB.PictureBox cmdEdit 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   1080
      Picture         =   "frmLogin.frx":223C1
      ScaleHeight     =   555
      ScaleWidth      =   2370
      TabIndex        =   5
      Top             =   3000
      Width           =   2370
   End
   Begin VB.PictureBox cmdNew 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   720
      Picture         =   "frmLogin.frx":256C6
      ScaleHeight     =   555
      ScaleWidth      =   2325
      TabIndex        =   4
      Top             =   2400
      Width           =   2325
   End
   Begin VB.PictureBox cmdConnect 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   360
      Picture         =   "frmLogin.frx":28981
      ScaleHeight     =   555
      ScaleWidth      =   1620
      TabIndex        =   3
      Top             =   1800
      Width           =   1620
   End
   Begin VB.TextBox txtPass 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   4200
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   4080
      Width           =   2055
   End
   Begin VB.TextBox txtUser 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   4200
      TabIndex        =   0
      Top             =   3480
      Width           =   2055
   End
   Begin VB.ListBox lstPlayer 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   2010
      Left            =   4200
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   990
      Width           =   2055
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private ontop As New clsOnTop
Public addnew As Boolean

Private Sub chkLocal_Click()
If frmLogin.chkLocal.Value = 1 Then
frmTCP.wskServer.RemoteHost = "localhost"
Else
frmTCP.wskServer.RemoteHost = serv
End If
End Sub

Private Sub cmdConnect_Click()

If txtUser.Text = "" Then GoTo nouser
If txtPass.Text = "" Then GoTo nopass

If addnew = True Then
Open (App.Path & "\player.dat") For Append As #1
Write #1, txtUser.Text, txtPass.Text
Close #1
lstPlayer.AddItem txtUser.Text
End If

isConnectedTryed = True
Load frmTCP
frmTCP.Show
Me.Hide
frmTCP.wskServer.RemotePort = porta
frmTCP.wskServer.Connect
Exit Sub

nouser:
msg = MsgBox("Please Enter a username!", vbOKOnly)
Exit Sub
nopass:
msg = MsgBox("Please Enter a password!", vbOKOnly)
Exit Sub
End Sub

Private Sub cmdDelete_Click()

dodelete = MsgBox("Are you sure you want to delete this character?", vbOKCancel, "Delete Character.")

If dodelete = vbOK Then

Open App.Path & "\player.dat" For Input As #1
Open App.Path & "\player2.dat" For Output As #2
        Do Until (EOF(1))
            Input #1, pName, pPass
       
            If lstPlayer = pName Then
               'Write #2, txtUser.text, txtPass.text
               addnew = False
               
            Else
            Write #2, pName, pPass
            addnew = False
            End If
            
        Loop
Close #1
Close #2

Open App.Path & "\player2.dat" For Input As #1
Open App.Path & "\player.dat" For Output As #2
        Do Until (EOF(1))
            Input #1, pName, pPass
            Write #2, pName, pPass
            
        Loop
Close #1
Close #2

Kill App.Path & "\player2.dat"

lstPlayer.Clear

On Error GoTo skip
Open App.Path & "\player.dat" For Input As #1
            'Input #1, pName, pPass
        Do Until (EOF(1))
            Input #1, pName, pPass
            
            lstPlayer.AddItem pName
            
        Loop
Close #1

skip:

End If
End Sub

Private Sub cmdEdit_Click()

Open App.Path & "\player.dat" For Input As #1
Open App.Path & "\player2.dat" For Output As #2
        Do Until (EOF(1))
            Input #1, pName, pPass
       
            If lstPlayer = pName Then
               Write #2, txtUser.Text, txtPass.Text
               addnew = False
               
            Else
            Write #2, pName, pPass
            addnew = False
            End If
            
        Loop
Close #1
Close #2

Open App.Path & "\player2.dat" For Input As #1
Open App.Path & "\player.dat" For Output As #2
        Do Until (EOF(1))
            Input #1, pName, pPass
            Write #2, pName, pPass
            
        Loop
Close #1
Close #2

Kill App.Path & "\player2.dat"

lstPlayer.Clear

On Error GoTo skip
Open App.Path & "\player.dat" For Input As #1
            'Input #1, pName, pPass
        Do Until (EOF(1))
            Input #1, pName, pPass
            
            lstPlayer.AddItem pName
            
        Loop
Close #1

skip:

End Sub

Private Sub cmdNew_Click()
Load frmNew
frmNew.Show
Me.Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Form_Load()

Load frmDebugger
frmDebugger.Show

  Set ontop = New clsOnTop
  ontop.MakeTopMost hWnd

addnew = True

On Error GoTo pass
    Kill App.Path & "\TidelUpdate.exe"
pass:

On Error GoTo skip
Open App.Path & "\player.dat" For Input As #1
            'Input #1, pName, pPass
        Do Until (EOF(1))
            Input #1, pName, pPass
            
            lstPlayer.AddItem pName
            
        Loop
Close #1

skip:
Exit Sub
End Sub

Private Sub Form_Activate()
frmTCP.selectedmidi = "login.mid"
Call MidiPlay
End Sub


Private Sub lstPlayer_Click()

On Error GoTo skip
Open App.Path & "\player.dat" For Input As #1
            'Input #1, pName, pPass
        Do Until lstPlayer = pName Or (EOF(1))
            Input #1, pName, pPass
            
            If lstPlayer = pName Then
               txtUser.Text = pName
               txtPass.Text = pPass
               addnew = False
            End If
            
        Loop
Close #1
skip:

'txtUser.text = lstPlayer
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdConnect = cmdConnect3
cmdNew = cmdNew3
cmdEdit = cmdEdit3
cmdDelete = cmdDelete3
End Sub

Private Sub cmdConnect_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdConnect = cmdConnect2
cmdNew = cmdNew3
cmdEdit = cmdEdit3
cmdDelete = cmdDelete3
End Sub

Private Sub cmdNew_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdNew = cmdNew2
cmdConnect = cmdConnect3
cmdEdit = cmdEdit3
cmdDelete = cmdDelete3
End Sub

Private Sub cmdEdit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdEdit = cmdEdit2
cmdConnect = cmdConnect3
cmdNew = cmdNew3
cmdDelete = cmdDelete3
End Sub

Private Sub cmdDelete_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdDelete = cmdDelete2
cmdConnect = cmdConnect3
cmdNew = cmdNew3
cmdEdit = cmdEdit3
End Sub
