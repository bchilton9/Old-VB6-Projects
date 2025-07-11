Attribute VB_Name = "char"
Function PaintChar(chara As Character)

'BitBlt frmMain.picBuffer.hdc, chara.X, chara.Y, 30, chara.Height, tiles.mask(chara.Container).hdc, 0, 0, vbMergePaint
'BitBlt frmMain.picBuffer.hdc, chara.X, chara.Y, 30, chara.Height, tiles.char(chara.Container).hdc, 0, 0, vbSrcAnd

If chara.Container = 1 Then
offsetx = 0
offsety = 0
ElseIf chara.Container = 2 Then
offsetx = 23
offsety = 0
ElseIf chara.Container = 3 Then
offsetx = 46
offsety = 0
ElseIf chara.Container = 4 Then
offsetx = 0
offsety = 32
ElseIf chara.Container = 5 Then
offsetx = 23
offsety = 32
ElseIf chara.Container = 6 Then
offsetx = 46
offsety = 32
ElseIf chara.Container = 7 Then
offsetx = 0
offsety = 64
ElseIf chara.Container = 8 Then
offsetx = 23
offsety = 64
ElseIf chara.Container = 9 Then
offsetx = 46
offsety = 64
ElseIf chara.Container = 10 Then
offsetx = 0
offsety = 96
ElseIf chara.Container = 11 Then
offsetx = 23
offsety = 96
ElseIf chara.Container = 12 Then
offsetx = 46
offsety = 96
End If


BitBlt frmMain.picBuffer.hdc, chara.X + 6, chara.Y, 20, chara.Height, tiles.picMask.hdc, offsetx + chara.offx, offsety + chara.offy, vbMergePaint
BitBlt frmMain.picBuffer.hdc, chara.X + 6, chara.Y, 20, chara.Height, tiles.picChar.hdc, offsetx + chara.offx, offsety + chara.offy, vbSrcAnd

Call SetTextColor(frmMain.picBuffer.hdc, vbBlue)
Call TextOut(frmMain.picBuffer.hdc, chara.X, chara.Y - 10, Trim(chara.Name), Len(Trim(chara.Name)))
Call SetTextColor(frmMain.picBuffer.hdc, vbWhite)
Call TextOut(frmMain.picBuffer.hdc, chara.X - 1, chara.Y - 11, Trim(chara.Name), Len(Trim(chara.Name)))

End Function

