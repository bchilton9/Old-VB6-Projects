Attribute VB_Name = "run"
Sub doseek(oldm, newm)
If frmBot.chkseek.Value = 1 Then
    xl = Mid(oldm, 2, 1)
    xl = Asc(xl) - 32
    x = xl * 2
    xl = Mid(oldm, 1, 1)
    xl = Asc(xl) - 32
    x = x + (xl * 2)
    
    yl = Mid(oldm, 4, 1)
    y = Asc(yl) - 32
    yl = Mid(oldm, 3, 1)
    yl = Asc(yl) - 32
    y = y + (yl * 2)

    lxl = Mid(newm, 2, 1)
    lxl = Asc(lxl) - 32
    lx = lxl * 2
    lxl = Mid(newm, 1, 1)
    lxl = Asc(lxl) - 32
    lx = lx + (lxl * 2)
    
    lyl = Mid(newm, 4, 1)
    ly = Asc(lyl) - 32
    lyl = Mid(newm, 3, 1)
    lyl = Asc(lyl) - 32
    ly = ly + (lyl * 2)


sx = 44
sy = 60


If sx > x Then ew = "e"
If sx < x Then ew = "w"
If sx = x Then ew = "l"

If sy > y Then ns = "s"
If sy < y Then ns = "n"
If sy = y Then ns = "l"

'If frmBot.Seek.Value = 1 Then
If lx = x And ly = y Then
frmBot.sckFurc.SendData "m 7" & vbLf

ElseIf ns = "n" And ew = "e" Then frmBot.sckFurc.SendData "m 9" & vbLf
ElseIf ns = "n" And ew = "w" Then frmBot.sckFurc.SendData "m 7" & vbLf
ElseIf ns = "s" And ew = "e" Then frmBot.sckFurc.SendData "m 3" & vbLf
ElseIf ns = "s" And ew = "w" Then frmBot.sckFurc.SendData "m 1" & vbLf

ElseIf ns = "l" And ew = "e" Then frmBot.sckFurc.SendData "m 9" & vbLf
ElseIf ns = "l" And ew = "w" Then frmBot.sckFurc.SendData "m 3" & vbLf

ElseIf ns = "n" And ew = "l" Then frmBot.sckFurc.SendData "m 9" & vbLf
ElseIf ns = "s" And ew = "l" Then frmBot.sckFurc.SendData "m 3" & vbLf

End If
End If

End Sub
