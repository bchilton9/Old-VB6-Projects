VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000012&
   Caption         =   "Game"
   ClientHeight    =   8760
   ClientLeft      =   4860
   ClientTop       =   1440
   ClientWidth     =   5940
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   584
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   396
   Begin VB.PictureBox picCoin 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   7
      Left            =   5040
      Picture         =   "main.frx":0000
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   11
      Top             =   1080
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox picCoin 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   6
      Left            =   4560
      Picture         =   "main.frx":0B0C
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   10
      Top             =   1080
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox picCoin 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   5
      Left            =   4080
      Picture         =   "main.frx":1618
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   9
      Top             =   1080
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox picCoin 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   4
      Left            =   3600
      Picture         =   "main.frx":2124
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   8
      Top             =   1080
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox picCoin 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   3
      Left            =   3120
      Picture         =   "main.frx":2C30
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   7
      Top             =   1080
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox picCoin 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   2
      Left            =   2640
      Picture         =   "main.frx":373C
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   6
      Top             =   1080
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox picCoin 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   1
      Left            =   2160
      Picture         =   "main.frx":4248
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   5
      Top             =   1080
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox picCoin 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   0
      Left            =   1680
      Picture         =   "main.frx":4D54
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   4
      Top             =   1080
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   8040
      Visible         =   0   'False
      Width           =   5415
   End
   Begin VB.TextBox txtMsg 
      Appearance      =   0  'Flat
      Height          =   975
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   7080
      Width           =   5415
   End
   Begin VB.PictureBox picMap 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   5385
      Left            =   240
      ScaleHeight     =   358.989
      ScaleMode       =   0  'User
      ScaleWidth      =   360
      TabIndex        =   0
      Top             =   1680
      Width           =   5400
   End
   Begin VB.Label lblCoins 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000012&
      Caption         =   "Coins:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000007&
      ForeColor       =   &H8000000E&
      Height          =   1095
      Left            =   720
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdNext_Click()
Select Case Guy.Location
    Case 12
    getCoin Kiki
End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim step As Integer
step = 30
If Guy.Location = Olga.Location Then obs(Olga.X / 30, Olga.Y / 30) = "X"
If Guy.Location = Merlon.Location Then obs(Merlon.X / 30, Merlon.Y / 30) = "X"
If KeyCode = vbKeyDown Then
    If obs(Int(Guy.X / 30), Int(Guy.Y + 30) / 30) <> "X" Then
        If obs(Int(Guy.X / 30), Int(Guy.Y + 30) / 30) = "W" Then
            If Guy.Flippers = True Then
                Guy.Container = 0
                Guy.MskContainer = 1
                Guy.Y = Guy.Y + step
                Guy.Height = 15
            End If
        ElseIf obs(Int(Guy.X / 30), Int(Guy.Y + 30) / 30) = "F" Then
            Guy.Height = 30
            Guy.Container = 0
            Guy.MskContainer = 1
            Guy.Y = Guy.Y + step
        ElseIf obs(Int(Guy.X / 30), Int(Guy.Y + 30) / 30) = "S" Then
            Guy.Height = 30
            Guy.Container = 0
            Guy.MskContainer = 1
            txtMsg.Text = sign(Guy.Location)
        End If
    End If
ElseIf KeyCode = vbKeyUp Then
    If obs(Int(Guy.X / 30), Int(Guy.Y - 30) / 30) <> "X" Then
        If obs(Int(Guy.X / 30), Int(Guy.Y - 30) / 30) = "W" Then
            If Guy.Flippers = True Then
            Guy.Container = 2
            Guy.MskContainer = 3
            Guy.Y = Guy.Y - step
            Guy.Height = 15
            End If
        ElseIf obs(Int(Guy.X / 30), Int(Guy.Y - 30) / 30) = "F" Then
        Guy.Height = 30
        Guy.Container = 2
        Guy.MskContainer = 3
        Guy.Y = Guy.Y - step
         ElseIf obs(Int(Guy.X / 30), Int(Guy.Y - 30) / 30) = "S" Then
        Guy.Height = 30
        Guy.Container = 2
        Guy.MskContainer = 3
        txtMsg.Text = sign(Guy.Location)
        End If
    End If
ElseIf KeyCode = vbKeyLeft Then
    If obs(Int((Guy.X - 30) / 30), Int(Guy.Y / 30)) <> "X" Then
        If obs(Int((Guy.X - 30) / 30), Int(Guy.Y / 30)) = "W" Then
            If Guy.Flippers = True Then
            Guy.Container = 6
            Guy.MskContainer = 7
            Guy.X = Guy.X - step
            Guy.Height = 15
            End If
        ElseIf obs(Int((Guy.X - 30) / 30), Int(Guy.Y / 30)) = "F" Then
        Guy.Height = 30
        Guy.Container = 6
        Guy.MskContainer = 7
        Guy.X = Guy.X - step
         ElseIf obs(Int((Guy.X - 30) / 30), Int(Guy.Y) / 30) = "S" Then
         Guy.Height = 30
        Guy.Container = 6
        Guy.MskContainer = 7
        txtMsg.Text = sign(Guy.Location)
        End If
    End If
ElseIf KeyCode = vbKeyRight Then
    If obs(Int((Guy.X + 30) / 30), Int(Guy.Y / 30)) <> "X" Then
        If obs(Int((Guy.X + 30) / 30), Int(Guy.Y / 30)) = "W" Then
            If Guy.Flippers = True Then
            Guy.Container = 4
            Guy.MskContainer = 5
            Guy.X = Guy.X + step
            Guy.Height = 15
            End If
        ElseIf obs(Int((Guy.X + 30) / 30), Int(Guy.Y / 30)) = "F" Then
        Guy.Height = 30
        Guy.Container = 4
        Guy.MskContainer = 5
        Guy.X = Guy.X + step
         ElseIf obs(Int((Guy.X + 30) / 30), Int(Guy.Y) / 30) = "S" Then
        Guy.Height = 30
        Guy.Container = 4
        Guy.MskContainer = 5
        txtMsg.Text = sign(Guy.Location)
        End If
    End If
End If
If Guy.X = 330 Then
Guy.Location = Guy.Location + 1
Guy.X = 30
txtMsg.Text = ""
cmdNext.Visible = False
End If
If Guy.Y = 0 Then
Guy.Location = Guy.Location + 5
Guy.Y = 300
txtMsg.Text = ""
cmdNext.Visible = False
End If
If Guy.X = 0 Then
Guy.Location = Guy.Location - 1
Guy.X = 300
txtMsg.Text = ""
cmdNext.Visible = False
End If
If Guy.Y = 330 Then
Guy.Location = Guy.Location - 5
Guy.Y = 60
txtMsg.Text = ""
cmdNext.Visible = False
End If

For i = 2 To 3
If Guy.X = items(i).X And Guy.Y = items(i).Y And Guy.Location = items(i).loc Then
    Select Case items(i).type
        Case 0 ' teleporters
        Guy.Location = items(i).field1
        Guy.X = items(i).field2
        Guy.Y = items(i).field3
        Case 1 ' coinholders
        If Kiki.coins = 1 Then
        txtMsg.Text = items(i).field1 & Guy.name & "," & items(i).field2
        cmdNext.Visible = True
        Else
        txtMsg.Text = "Hurry and retrieve the rest of the coins! You only have " & Guy.coins _
        & " coins! Hurry!"
        End If
        Case 2 ' people
        txtMsg.Text = items(i).field1
    End Select
End If
Next i


Label1.Caption = Guy.X & "," & Guy.Y
drawMap (Guy.Location)
For i = 1 To 4 ' paint items
    If items(i).loc = Guy.Location Then
    Select Case items(i).type
    Case 0
    picMap.PaintPicture frmtex.tex(items(i).pic), items(i).X, items(i).Y
    End Select
    End If
Next i
PaintChar Guy
PaintChar Kiki
PaintChar Merlon
PaintChar Olga
For i = 1 To 1
    If coin(i).coins = 1 Then PaintChar coin(i)
    If coin(i).coins = 1 And Guy.X = coin(i).X And Guy.Y = coin(i).Y And coin(i).Location = Guy.Location Then
    Guy.coins = Guy.coins + 1
    coin(i).coins = coin(i).coins - 1
    drawMap (Guy.Location)
    End If
Next i
If Guy.coins > -1 Then
    For i = 0 To Guy.coins
        picCoin(i).Visible = True
    Next i
End If
End Sub



Private Sub Form_Load()
'Guy.name = InputBox("Please enter your name.", "Darrins Game")
If Guy.name = "" Then Guy.name = "Guest"
InputChar
inputMaps
InitChar
Open App.Path & "\Misc\signs.txt" For Input As #1
For i = 1 To 9
Input #1, sign(i)
Next i
Close #1
Load frmtex
frmtex.Visible = True
drawMap (Guy.Location)
PaintChar Guy
For i = 1 To 4 ' paint items
    If items(i).loc = Guy.Location Then
    Select Case items(i).type
    Case 0
    picMap.PaintPicture frmtex.tex(items(i).pic), items(i).X, items(i).Y
    
    End Select
    End If
Next i
End Sub


Private Sub Form_Terminate()
End
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub
