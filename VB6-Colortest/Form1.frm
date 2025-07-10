VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5055
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5055
   ScaleWidth      =   12000
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox CorPicB 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   540
      Left            =   2160
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   383
      TabIndex        =   27
      Top             =   5520
      Width           =   5805
   End
   Begin VB.OptionButton optPart 
      Caption         =   "Armor Dark"
      Height          =   255
      Index           =   8
      Left            =   240
      TabIndex        =   26
      Top             =   4680
      Width           =   1215
   End
   Begin VB.PictureBox picColor 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   420
      Index           =   8
      Left            =   1440
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   25
      Top             =   4560
      Width           =   405
   End
   Begin VB.OptionButton optPart 
      Caption         =   "Armor Med"
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   24
      Top             =   4200
      Width           =   1215
   End
   Begin VB.PictureBox picColor 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   420
      Index           =   7
      Left            =   1440
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   23
      Top             =   4080
      Width           =   405
   End
   Begin VB.OptionButton optPart 
      Caption         =   "Armor Light"
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   22
      Top             =   3720
      Width           =   1215
   End
   Begin VB.PictureBox picColor 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   420
      Index           =   6
      Left            =   1440
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   21
      Top             =   3600
      Width           =   405
   End
   Begin VB.OptionButton optPart 
      Caption         =   "Skin Dark"
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   20
      Top             =   3240
      Width           =   1215
   End
   Begin VB.PictureBox picColor 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   420
      Index           =   5
      Left            =   1440
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   19
      Top             =   3120
      Width           =   405
   End
   Begin VB.OptionButton optPart 
      Caption         =   "Skin Med"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   18
      Top             =   2760
      Width           =   1215
   End
   Begin VB.PictureBox picColor 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   420
      Index           =   4
      Left            =   1440
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   17
      Top             =   2640
      Width           =   405
   End
   Begin VB.OptionButton optPart 
      Caption         =   "Skin Light"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   16
      Top             =   2280
      Width           =   1215
   End
   Begin VB.PictureBox picColor 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   420
      Index           =   3
      Left            =   1440
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   15
      Top             =   2160
      Width           =   405
   End
   Begin VB.OptionButton optPart 
      Caption         =   "Hair Dark"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   14
      Top             =   1800
      Width           =   1215
   End
   Begin VB.PictureBox picColor 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   420
      Index           =   2
      Left            =   1440
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   13
      Top             =   1680
      Width           =   405
   End
   Begin VB.OptionButton optPart 
      Caption         =   "Hair Med"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   12
      Top             =   1320
      Width           =   1215
   End
   Begin VB.PictureBox picColor 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   420
      Index           =   1
      Left            =   1440
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   11
      Top             =   1200
      Width           =   405
   End
   Begin VB.PictureBox picColor 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   420
      Index           =   0
      Left            =   1440
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   2
      Top             =   720
      Width           =   405
   End
   Begin VB.OptionButton optPart 
      Caption         =   "Hair Light"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   10
      Top             =   840
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   3240
      Left            =   6120
      ScaleHeight     =   212
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   212
      TabIndex        =   9
      Top             =   720
      Width           =   3240
   End
   Begin VB.PictureBox Picture3 
      Height          =   375
      Left            =   1320
      ScaleHeight     =   315
      ScaleWidth      =   435
      TabIndex        =   8
      Top             =   5520
      Width           =   495
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   240
      Top             =   5520
   End
   Begin VB.CommandButton Command3 
      Caption         =   ">"
      Height          =   375
      Left            =   6120
      TabIndex        =   7
      Top             =   4080
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5520
      TabIndex        =   6
      Top             =   4080
      Width           =   375
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   3240
      Left            =   2640
      ScaleHeight     =   212
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   212
      TabIndex        =   5
      Top             =   720
      Width           =   3240
   End
   Begin VB.PictureBox CorPic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   540
      Left            =   6120
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   383
      TabIndex        =   4
      Top             =   120
      Width           =   5805
   End
   Begin VB.PictureBox picReplace 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   540
      Left            =   720
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   3
      Top             =   5520
      Width           =   540
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Process"
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   4440
      Width           =   1335
   End
   Begin VB.PictureBox OrPic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   540
      Left            =   120
      Picture         =   "Form1.frx":0068
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   383
      TabIndex        =   0
      Top             =   120
      Width           =   5805
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim offset As Integer

Private Sub Command1_Click()
Dim colorA, colorB As Long

Form1.MousePointer = 11

CorPic.Picture = OrPic.Picture
CorPicB.Picture = OrPic.Picture

For i = 0 To 8

colorA = picColor(i).Point(0, 0)

If i = 0 Then
colorB = picReplace.Point(0, 0)
ElseIf i = 1 Then
colorB = picReplace.Point(1, 0)
ElseIf i = 2 Then
colorB = picReplace.Point(2, 0)
ElseIf i = 3 Then
colorB = picReplace.Point(0, 1)
ElseIf i = 4 Then
colorB = picReplace.Point(1, 1)
ElseIf i = 5 Then
colorB = picReplace.Point(2, 1)
ElseIf i = 6 Then
colorB = picReplace.Point(0, 2)
ElseIf i = 7 Then
colorB = picReplace.Point(1, 2)
ElseIf i = 8 Then
colorB = picReplace.Point(2, 2)
End If

    For Y = 0 To CorPicB.ScaleHeight
        For X = 0 To CorPicB.ScaleWidth
                'if color of point is = to mask color then draw white point
                If CorPicB.Point(X, Y) = colorA Then
                CorPic.PSet (X, Y), colorB
                'Else
                'otherwise keep color as is
                'CorPic.PSet (X, Y), CorPicB.Point(X, Y)
                End If
        Next X
    Next Y
    
Call StretchBlt(CorPicB.hdc, 0, 0, CorPicB.ScaleWidth, CorPicB.ScaleHeight, CorPic.hdc, 0, 0, CorPic.ScaleWidth, CorPic.ScaleHeight, vbSrcAnd)

Next i

Form1.MousePointer = 0

End Sub

Private Sub Command2_Click()
offset = offset - 32
If offset = 0 Then
Command2.Enabled = False
End If
Command3.Enabled = True
End Sub

Private Sub Command3_Click()
offset = offset + 32
If offset = 352 Then
Command3.Enabled = False
End If
Command2.Enabled = True
End Sub

Private Sub Form_Load()
offset = 0
End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim color As Long
Dim i As Integer

color = Picture2.Point(X, Y)

If optPart(0).Value = True Then
i = 0
ElseIf optPart(1).Value = True Then
i = 1
ElseIf optPart(2).Value = True Then
i = 2
ElseIf optPart(3).Value = True Then
i = 3
ElseIf optPart(4).Value = True Then
i = 4
ElseIf optPart(5).Value = True Then
i = 5
ElseIf optPart(6).Value = True Then
i = 6
ElseIf optPart(7).Value = True Then
i = 7
ElseIf optPart(8).Value = True Then
i = 8
End If

    For iY = 0 To picColor(i).ScaleHeight
        For iX = 0 To picColor(i).ScaleWidth
                
            picColor(i).PSet (iX, iY), color
   
        Next
    Next


End Sub

Private Sub Timer1_Timer()
Picture2.Picture = Picture3.Picture
Picture1.Picture = Picture3.Picture
Call StretchBlt(Picture2.hdc, 0, 0, 192, 192, OrPic.hdc, offset, 0, 32, 32, vbSrcAnd)

Call StretchBlt(Picture1.hdc, 0, 0, 192, 192, CorPic.hdc, offset, 0, 32, 32, vbSrcAnd)

End Sub
