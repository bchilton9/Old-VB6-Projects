Attribute VB_Name = "incomeing"
Sub incomeingtxt(Txt)

If Left(Txt, 15) Like "(Server going d" Then
    frmBot.sckFurc.Close
    frmBot.Connected = False
    frmBot.txtcnt = "False"
End If
If Left(Txt, 15) Like "(Someone else h" Then
    frmBot.sckFurc.Close
    frmBot.Connected = False
    frmBot.txtcnt = "False"
End If
If Left(Txt, 15) Like "(Disconnected f" Then
    frmBot.sckFurc.Close
    frmBot.Connected = False
    frmBot.txtcnt = "False"
End If

If frmBot.chkWhisp.Value = Checked Or frmBot.chkWhisp.Enabled = False Then
    If Left(Txt, 10) = "((You see " Then
        Furre = Mid(Txt, 11, Len(Txt) - 12)
        YouSee Furre, Sign
    End If

    If Txt Like "7*" Then
        Trg = Mid(Txt, 11, 1)
        Pos = Right(Txt, 4)
        TrgPos = Mid(Txt, 6, 4)
        SignTrg Trg, Pos, TrgPos
    End If
    
    If Left(Txt, 3) = "([ " And Right(Txt, 10) = " to you. ]" Then
        Tmsg = Split(Txt, " whispers, " & Chr(34))
        Furre = Right(Tmsg(0), Len(Tmsg(0)) - 3)
        Msg = Left(Tmsg(1), Len(Tmsg(1)) - 11)
        whispers.DoWhisper Furre, Msg
    End If
    


Else
    If Left(Txt, 3) = "([ " And Right(Txt, 10) = " to you. ]" Then
    Tmsg = Split(Txt, " whispers, " & Chr(34), 2)
    Furre = Right(Tmsg(0), Len(Tmsg(0)) - 3)
    frmBot.sckFurc.SendData "wh " & Furre & " Im currently offline please try again later." & vbLf
    End If
End If
Exit Sub
error:
    frmBot.sckFurc.SendData "wh " & Furre & " Im Sorry, " & Chr(34) & Msg & Chr(34) & " is not a valid command. Whisper me *help to learn how to use my service." & vbLf
    Resume stoptrying
stoptrying:
End Sub

