VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "MapEditor"
   ClientHeight    =   6060
   ClientLeft      =   1035
   ClientTop       =   1320
   ClientWidth     =   7950
   LinkTopic       =   "Form1"
   ScaleHeight     =   404
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   530
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   495
      Left            =   6120
      TabIndex        =   3
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "&Load"
      Height          =   495
      Left            =   6120
      TabIndex        =   2
      Top             =   1080
      Width           =   1695
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   495
      Left            =   6120
      TabIndex        =   1
      Top             =   360
      Width           =   1695
   End
   Begin VB.PictureBox picMap 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   5400
      Left            =   480
      ScaleHeight     =   360
      ScaleMode       =   0  'User
      ScaleWidth      =   360
      TabIndex        =   0
      Top             =   360
      Width           =   5400
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   855
      Left            =   6120
      TabIndex        =   4
      Top             =   2880
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type POINTAPI
X As Long
Y As Long
End Type
Dim ppp As POINTAPI
Dim map(0 To 11, 0 To 11) As Integer



Private Sub cmdExit_Click()
End
End Sub

Private Sub cmdLoad_Click()
On Error Resume Next
Dim filedes As String
filedes = InputBox("Name of map to open.")

drawMap (filedes)

Dim strMap As String
Open App.Path & "\Maps\map" & filedes & ".txt" For Input As #1
For j = 1 To 12
    For i = 1 To 12
    Input #1, strMap
    map(j - 1, i - 1) = strMap
    Next i
Next j
Close #1
Form1.Caption = filedes
End Sub

Private Sub cmdSave_Click()
On Error Resume Next
Dim filedes As String
filedes = InputBox("Please enter a name for your map.")
Open App.Path & "\Maps\map" & filedes & ".txt" For Output As #1
For i = 0 To 11
    For j = 0 To 11
    Write #1, map(i, j);
    Next j
Next i
For i = 0 To 11
    For j = 0 To 11
    Select Case map(i, j)
                    Case 2
                    obs(i, j) = "F"
                    Case 52 To 54
                    obs(i, j) = "F"
                    Case 7
                    obs(i, j) = "F"
                    Case 16 To 23
                    obs(i, j) = "F"
                    Case 42
                    obs(i, j) = "F"
                    Case 44 To 45
                    obs(i, j) = "F"
                    Case 0 To 1
                    obs(i, j) = "W"
                    Case 36
                    obs(i, j) = "S"
                    Case Else
                    obs(i, j) = "X"
                    End Select
     If obs(i, j) = "" Then
    obs(i, j) = "W"
    End If
    Write #1, obs(i, j);
    Next j
Next i
Close #1
End Sub

Private Sub Form_Load()
On Error Resume Next
inputMaps

Load frmtex
frmtex.Visible = True

For i = 0 To 11
picMap.Line (i * 30, picMap.ScaleTop)-(i * 30, picMap.ScaleTop + picMap.ScaleHeight)
Next i
For i = 0 To 11
picMap.Line (0, i * 30)-(360, i * 30)
Next i
For i = 0 To 11
    For j = 0 To 11
    picMap.PaintPicture frmtex.tex(0).Picture, i * 30, j * 30
    Next j
Next i
End Sub

Private Sub picMap_Click()
On Error Resume Next
map(Int(ppp.X / 30), Int(ppp.Y / 30)) = cur

picMap.PaintPicture frmtex.tex(cur).Picture, Int(ppp.X / 30) * 30, Int(ppp.Y / 30) * 30
End Sub

Private Sub picMap_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
ppp.X = X
ppp.Y = Y
Label1.Caption = X & "," & Y
End Sub
