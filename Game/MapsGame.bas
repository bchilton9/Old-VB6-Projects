Attribute VB_Name = "Maps"
Type Item
loc As Integer
X As Single
Y As Single
obName As String
type As Integer
field1 As Variant
field2 As Variant
field3 As Variant
pic As Integer
End Type

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

Sub drawMap(map As Integer)
'On Local Error Resume Next

If startmidi = True And frmTCP.selectedmidi <> "m" & map & ".mid" Then
On Error GoTo passmidi
Open App.Path & "\midis\m" & map & ".mid" For Input As #1
Input #1, bla
Close #1

frmTCP.selectedmidi = "m" & map & ".mid"
Call MidiPlay
passmidi:
End If

Open App.Path & "\maps\map" & map & ".dat" For Input As #1

'mape = Split(mymap, ",")

  For l = 1 To 9
    For Y = 0 To MapY
      For X = 0 To MapX

Input #1, mapstr
        
     With LocalMap.Tile(X, Y)
     
     If l = 1 Then
        .TileX = mapstr
     ElseIf l = 2 Then
        .TileY = mapstr
     ElseIf l = 3 Then
        .Walk = mapstr
     ElseIf l = 4 Then
        .objectX = mapstr
     ElseIf l = 5 Then
        .objectY = mapstr
     ElseIf l = 6 Then
        .objectoverX = mapstr
     ElseIf l = 7 Then
        .objectoverY = mapstr
     ElseIf l = 8 Then
        .npcX = mapstr
     ElseIf l = 9 Then
        .npcY = mapstr
     End If
     
     End With

      Next X
    Next Y
  Next l

Close #1
refreshmap

startmidi = False

End Sub

Private Sub refreshmap()

Dim X, Y As Long

    For Y = 0 To MapY
      For X = 0 To MapX
        With LocalMap.Tile(X, Y)
         
StretchBlt frmMain.picBuffer.hdc, X * PicX, Y * PicY, 32, 32, frmMain.picSelect.hdc, .TileX * 16, .TileY * 16, 16, 16, &HCC0020
         

StretchBlt frmMain.picBuffer.hdc, X * PicX, Y * PicY, 32, 32, frmMain.picObjectSelectMask.hdc, .objectX * 16, .objectY * 16, 16, 16, vbMergePaint
StretchBlt frmMain.picBuffer.hdc, X * PicX, Y * PicY, 32, 32, frmMain.picObjectSelectB.hdc, .objectX * 16, .objectY * 16, 16, 16, vbSrcAnd

StretchBlt frmMain.picBuffer.hdc, X * PicX, Y * PicY, 32, 32, frmMain.picNpcSelectMask.hdc, .npcX * 32, .npcY * 32, 32, 32, vbMergePaint
StretchBlt frmMain.picBuffer.hdc, X * PicX, Y * PicY, 32, 32, frmMain.picNpcSelectB.hdc, .npcX * 32, .npcY * 32, 32, 32, vbSrcAnd

        End With
      Next X
    Next Y

End Sub
