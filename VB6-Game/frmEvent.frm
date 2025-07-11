VERSION 5.00
Begin VB.Form frmAddMap 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Event Map"
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11535
   DrawStyle       =   5  'Transparent
   Icon            =   "frmEvent.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   11535
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picGrid 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   12000
      Picture         =   "frmEvent.frx":0E42
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   8
      Top             =   3720
      Width           =   540
   End
   Begin VB.PictureBox picGridMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   12600
      Picture         =   "frmEvent.frx":1A86
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   7
      Top             =   3720
      Width           =   540
   End
   Begin VB.PictureBox picObjectSelectMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   18780
      Left            =   11760
      Picture         =   "frmEvent.frx":26C8
      ScaleHeight     =   18720
      ScaleWidth      =   5760
      TabIndex        =   6
      Top             =   6360
      Width           =   5820
   End
   Begin VB.PictureBox picObjectSelectB 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   18780
      Left            =   5880
      Picture         =   "frmEvent.frx":16170A
      ScaleHeight     =   18720
      ScaleWidth      =   5760
      TabIndex        =   5
      Top             =   6360
      Width           =   5820
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   5880
      Width           =   1215
   End
   Begin VB.PictureBox picSelect 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   18720
      Left            =   0
      Picture         =   "frmEvent.frx":2C074C
      ScaleHeight     =   18720
      ScaleWidth      =   5760
      TabIndex        =   1
      Top             =   6360
      Width           =   5760
   End
   Begin VB.PictureBox picBufferMap 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5775
      Left            =   0
      ScaleHeight     =   383
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   767
      TabIndex        =   0
      Top             =   0
      Width           =   11535
   End
   Begin VB.Label txtLocation 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2160
      TabIndex        =   4
      Top             =   5880
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Location:"
      Height          =   255
      Left            =   1320
      TabIndex        =   3
      Top             =   5880
      Width           =   735
   End
End
Attribute VB_Name = "frmAddMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private ontop As New clsOnTop

Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, _
ByVal X As Long, _
ByVal Y As Long, _
ByVal nWidth As Long, _
ByVal nHeight As Long, _
ByVal hSrcDC As Long, _
ByVal xSrc As Long, _
ByVal ySrc As Long, _
ByVal nSrcWidth As Long, _
ByVal nSrcHeight As Long, _
ByVal dwRop As Long) As Long

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub Form_Load()

  Set ontop = New clsOnTop
  ontop.MakeTopMost hWnd
  
    For Y = 0 To MapY
      For X = 0 To MapX

     With LocalMapB.Tile(X, Y)
       .TileX = 0
       .TileY = 0
       .walk = 1
       .objectX = 0
       .objectY = 0
       .objectoverX = 0
       .objectoverY = 0
     End With

      Next X
    Next Y

Me.Show

Call cmdLoad_Click(frmEdit.txtWarpMap.Text)

End Sub

Sub cmdLoad_Click(map As Integer)
'On Error GoTo stopit

'MapNum = frmEventEdit.txtMapLoaded.Text
'= InputBox("Load map number?", "Load Map")

Open App.Path & "\maps\map" & map & ".dat" For Input As #1

  For l = 1 To 7
    For Y = 0 To MapY
      For X = 0 To MapX

        Input #1, mapsrt
        
     With LocalMapB.Tile(X, Y)
     
     If l = 1 Then
        .TileX = mapsrt
     ElseIf l = 2 Then
        .TileY = mapsrt
     ElseIf l = 3 Then
        .walk = mapsrt
     ElseIf l = 4 Then
        .objectX = mapsrt
     ElseIf l = 5 Then
        .objectY = mapsrt
     ElseIf l = 6 Then
        .objectoverX = mapsrt
     ElseIf l = 7 Then
        .objectoverY = mapsrt
     End If
     
     End With

      Next X
    Next Y
  Next l
  
Close #1

refreshmap

Exit Sub
stopit:
msg = MsgBox("Map number " & MapNum & " not found!", vbOKOnly, "Error")
Unload Me
End Sub

Private Sub refreshmap()

Dim X, Y As Long

    For Y = 0 To MapY
      For X = 0 To MapX
        With LocalMapB.Tile(X, Y)
                
StretchBlt picBufferMap.hdc, X * PicX, Y * PicY, 32, 32, picSelect.hdc, .TileX * 16, .TileY * 16, 16, 16, &HCC0020
            
 If .objectX = 0 And .objectY = 0 Then
  'donuthing
  Else
StretchBlt picBufferMap.hdc, X * PicX, Y * PicY, 32, 32, picObjectSelectMask.hdc, .objectX * 16, .objectY * 16, 16, 16, vbMergePaint
StretchBlt picBufferMap.hdc, X * PicX, Y * PicY, 32, 32, picObjectSelectB.hdc, .objectX * 16, .objectY * 16, 16, 16, vbSrcAnd
           Call SetTextColor(picBufferMap.hdc, vbYellow)
           Call TextOut(picBufferMap.hdc, X * PicX + 1, Y * PicY + 19, "U", 1)
   End If

If .objectoverX = 0 And .objectoverY = 0 Then
  'do nuthing
  Else
  StretchBlt picBufferMap.hdc, X * PicX, Y * PicY, 32, 32, picObjectSelectMask.hdc, .objectoverX * 16, .objectoverY * 16, 16, 16, vbMergePaint
  StretchBlt picBufferMap.hdc, X * PicX, Y * PicY, 32, 32, picObjectSelectB.hdc, .objectoverX * 16, .objectoverY * 16, 16, 16, vbSrcAnd
           Call SetTextColor(picBufferMap.hdc, vbYellow)
           Call TextOut(picBufferMap.hdc, X * PicX + 1, Y * PicY + 19, "O", 1)
   End If

BitBlt picBufferMap.hdc, X * PicX, Y * PicY, 32, 32, picGridMask.hdc, 0, 0, vbMergePaint
BitBlt picBufferMap.hdc, X * PicX, Y * PicY, 32, 32, picGrid.hdc, 0, 0, vbSrcAnd
                
           If .walk = 1 Then
           Call SetTextColor(picBufferMap.hdc, vbWhite)
             Call TextOut(picBufferMap.hdc, X * PicX + 8, Y * PicY + 8, "B", 1)
           End If
           
           If .walk = 2 Then
           Call SetTextColor(picBufferMap.hdc, vbWhite)
             Call TextOut(picBufferMap.hdc, X * PicX + 8, Y * PicY + 8, "W", 1)
           End If
           
        End With
      Next X
    Next Y

Open App.Path & "\events\event" & frmEdit.txtWarpMap.Text & ".dat" For Input As #1
Do Until (EOF(1))

Line Input #1, eventdata

If eventdata <> "" Then
myevent = Split(eventdata, pChar)

If myevent(2) = "warp" Then
leter = "Wa"
ElseIf myevent(2) = "store" Then
leter = "St"
ElseIf myevent(2) = "bank" Then
leter = "Ba"
ElseIf myevent(2) = "sign" Then
leter = "Si"
End If

Call SetTextColor(picBufferMap.hdc, vbGreen)
Call TextOut(picBufferMap.hdc, myevent(0) * PicX, myevent(1) * PicY, leter, 2)


End If
Loop

Close #1

End Sub

Private Sub picBufferMap_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim x1, y1 As Long

  x1 = Int(X / PicX)
  y1 = Int(Y / PicY)

'If frmEdit.chkMap.Value = 1 Then
'frmEdit.txtWarpX.Text = x1
'frmEdit.txtWarpY.Text = y1
'frmEdit.chkMap.Value = 0
'Else
frmEdit.txtWarpX.Text = x1
frmEdit.txtWarpY.Text = y1
'End If

Unload Me

End Sub

Private Sub picBufferMap_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim x1, y1 As Long

  x1 = Int(X / PicX)
  y1 = Int(Y / PicY)
  
  
txtLocation.Caption = x1 & "," & y1

End Sub
