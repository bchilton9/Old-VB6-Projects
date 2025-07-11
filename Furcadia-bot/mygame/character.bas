Attribute VB_Name = "characterStuff"

Function PaintChar(chara As Character)
If chara.Location = Guy.Location Then
BitBlt Form1.picMap.hDC, chara.X, chara.Y, 30, chara.Height, frmtex.char(chara.MskContainer).hDC, 0, 0, vbMergePaint
BitBlt Form1.picMap.hDC, chara.X, chara.Y, 30, chara.Height, frmtex.char(chara.Container).hDC, 0, 0, vbSrcAnd
Form1.picMap.Refresh
End If
End Function



Sub InputChar()
For i = 0 To 15
frmtex.char(i) = LoadPicture(App.Path & "\characters\" & i & ".bmp")
Next i
End Sub

Sub InitChar()
With Guy
    .Bitmap = frmtex.char(0).Picture
    .Mask = frmtex.char(1).Picture
    .X = 180
    .Y = 210
    .Container = 0
    .MskContainer = 1
    .Location = 1
    .Flippers = True
    .Height = 30
    .coins = -1
    .Index = 0
End With
PaintChar Guy
With Kiki
    .Bitmap = frmtex.char(8).Picture
    .Mask = frmtex.char(9).Picture
    .X = 180
    .Y = 210
    .Container = 8
    .MskContainer = 9
    .Location = 12
    .Flippers = True
    .Height = 30
    .name = "Kiki"
    .coins = 1
    .Index = 1
End With

With Olga
    .Bitmap = frmtex.char(10).Picture
    .Mask = frmtex.char(11).Picture
    .X = 180
    .Y = 270
    .Container = 10
    .MskContainer = 11
    .Height = 30
    .name = "Olga"
    .Location = 9
    .Index = 2
End With

With Merlon
    .Bitmap = frmtex.char(12).Picture
    .Mask = frmtex.char(13).Picture
    .X = 240
    .Y = 120
    .Container = 12
    .MskContainer = 13
    .Height = 30
    .name = "Merlon"
    .Location = 8
    .Index = 3
End With

With coin(1)
    .Bitmap = frmtex.char(14).Picture
    .Mask = frmtex.char(15).Picture
    .X = 180
    .Y = 150
    .Container = 14
    .MskContainer = 15
    .Height = 30
    .name = "Coin1"
    .Location = 2
    .Index = 4
    .coins = 1
End With
    
End Sub

