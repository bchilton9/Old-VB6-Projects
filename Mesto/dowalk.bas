Attribute VB_Name = "walk"
Sub dowalk(whatwalk, lastwalk)
On Error Resume Next
If (lastwalk = "j") Or (lastwalk = "k") Or (lastwalk = "l") Then ' moveing up
    If (whatwalk = "j") Or (whatwalk = "k") Or (whatwalk = "l") Then
        frmBot.sckFurc.SendData "m 9" & vbLf
        Else ' move left
            If (whatwalk = "f") Or (whatwalk = "g") Or (whatwalk = "h") Then
            frmBot.sckFurc.SendData "m 9" & vbLf
            frmBot.sckFurc.SendData "m 7" & vbLf
        Else 'move right
            If (whatwalk = "b") Or (whatwalk = "c") Or (whatwalk = "d") Then
            frmBot.sckFurc.SendData "m 9" & vbLf
            frmBot.sckFurc.SendData "m 3" & vbLf
        Else ' move down
            If (whatwalk = "`") Or (whatwalk = "_") Or (whatwalk = "^") Then
            frmBot.sckFurc.SendData "m 7" & vbLf
            frmBot.sckFurc.SendData "m 9" & vbLf
            frmBot.sckFurc.SendData "m 9" & vbLf
            frmBot.sckFurc.SendData "m 3" & vbLf
        End If
        End If
        End If
    End If
lastwalk = whatwalk
End If
If (lastwalk = "f") Or (lastwalk = "g") Or (lastwalk = "h") Then 'moveing left
        If (whatwalk = "f") Or (whatwalk = "g") Or (whatwalk = "h") Then
        frmBot.sckFurc.SendData "m 7" & vbLf
                Else 'move up
                    If (whatwalk = "j") Or (whatwalk = "k") Or (whatwalk = "l") Then
                    frmBot.sckFurc.SendData "m 7" & vbLf
                    frmBot.sckFurc.SendData "m 9" & vbLf
                Else 'move down
                    If (whatwalk = "`") Or (whatwalk = "_") Or (whatwalk = "^") Then
                    frmBot.sckFurc.SendData "m 7" & vbLf
                    frmBot.sckFurc.SendData "m 1" & vbLf
                Else 'move right
                If (whatwalk = "b") Or (whatwalk = "c") Or (whatwalk = "d") Then
                    frmBot.sckFurc.SendData "m 9" & vbLf
                    frmBot.sckFurc.SendData "m 7" & vbLf
                    frmBot.sckFurc.SendData "m 7" & vbLf
                    frmBot.sckFurc.SendData "m 1" & vbLf
                End If
                End If
                End If
        End If
lastwalk = whatwalk
End If
If (lastwalk = "b") Or (lastwalk = "c") Or (lastwalk = "d") Then 'moveing right
        If (whatwalk = "b") Or (whatwalk = "c") Or (whatwalk = "d") Then
        frmBot.sckFurc.SendData "m 3" & vbLf
                Else 'move up
                    If (whatwalk = "j") Or (whatwalk = "k") Or (whatwalk = "l") Then
                    frmBot.sckFurc.SendData "m 3" & vbLf
                    frmBot.sckFurc.SendData "m 9" & vbLf
                Else 'move down
                    If (whatwalk = "`") Or (whatwalk = "_") Or (whatwalk = "^") Then
                    frmBot.sckFurc.SendData "m 3" & vbLf
                    frmBot.sckFurc.SendData "m 1" & vbLf
                Else 'move left
                    If (whatwalk = "f") Or (whatwalk = "g") Or (whatwalk = "h") Then
                    frmBot.sckFurc.SendData "m 9" & vbLf
                    frmBot.sckFurc.SendData "m 3" & vbLf
                    frmBot.sckFurc.SendData "m 3" & vbLf
                    frmBot.sckFurc.SendData "m 1" & vbLf
                End If
                End If
                End If
        End If
lastwalk = whatwalk
End If
If (lastwalk = "`") Or (lastwalk = "_") Or (lastwalk = "^") Then 'moveing down
        If (whatwalk = "`") Or (whatwalk = "_") Or (whatwalk = "^") Then
        frmBot.sckFurc.SendData "m 1" & vbLf
                Else 'move left
                    If (whatwalk = "f") Or (whatwalk = "g") Or (whatwalk = "h") Then
                    frmBot.sckFurc.SendData "m 1" & vbLf
                    frmBot.sckFurc.SendData "m 7" & vbLf
                Else 'move right
                    If (whatwalk = "b") Or (whatwalk = "c") Or (whatwalk = "d") Then
                    frmBot.sckFurc.SendData "m 1" & vbLf
                    frmBot.sckFurc.SendData "m 3" & vbLf
                Else 'move up
                    If (whatwalk = "j") Or (whatwalk = "k") Or (whatwalk = "l") Then
                    frmBot.sckFurc.SendData "m 7" & vbLf
                    frmBot.sckFurc.SendData "m 1" & vbLf
                    frmBot.sckFurc.SendData "m 1" & vbLf
                    frmBot.sckFurc.SendData "m 3" & vbLf
                End If
                End If
                End If
        End If
lastwalk = whatwalk
End If
If (lastwalk = "none") Then
        If (whatwalk = "`") Or (whatwalk = "_") Or (whatwalk = "^") Then
            frmBot.sckFurc.SendData "m 1" & vbLf
        Else 'move left
            If (whatwalk = "f") Or (whatwalk = "g") Or (whatwalk = "h") Then
            frmBot.sckFurc.SendData "m 7" & vbLf
        Else 'move right
            If (whatwalk = "b") Or (whatwalk = "c") Or (whatwalk = "d") Then
            frmBot.sckFurc.SendData "m 3" & vbLf
        Else 'move left
            If (whatwalk = "j") Or (whatwalk = "k") Or (whatwalk = "l") Then
            frmBot.sckFurc.SendData "m 9" & vbLf
        End If
        End If
        End If
        End If
lastwalk = whatwalk
End If
End Sub
