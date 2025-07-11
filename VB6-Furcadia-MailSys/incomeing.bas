Attribute VB_Name = "incomeing"
Sub incomeingtxt(Txt)
On Error Resume Next
If Left(Txt, 24) Like "(* You are not a phoenix" Then
frmBot.sckFurc.SendData "phoenix" & vbLf
frmBot.sckFurc.SendData "flame" & vbLf
End If
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

Dim sndrr As Integer
If frmBot.chkWhisp.Value = Checked Or frmBot.chkWhisp.Enabled = False Then
If Left(Txt, 3) = "([ " And Right(Txt, 10) = " to you. ]" Then
    tmsg = Split(Txt, " whispers, " & Chr(34), 2)
    Furre = Right(tmsg(0), Len(tmsg(0)) - 3)
    NMsg = Left(tmsg(1), Len(tmsg(1)) - 11)
    Msg = LCase(NMsg)
    Furre = LCase(Furre)
    If Left(Msg, 5) = "read " Then
    On Error GoTo error
        sndrr = Right(Msg, Len(Msg) - 5)
        read.remsg Furre, sndrr, Txt
    ElseIf Left(Msg, 7) = "delete " Then
    On Error GoTo error
        sndrr = Right(Msg, Len(Msg) - 7)
        delete.dodelete Furre, sndrr, Txt
    ElseIf Left(Msg, 5) = "send " Then
    On Error GoTo error
        aMsg = Split(Msg, " message ", 2)
        snd = Right(aMsg(0), Len(aMsg(0)) - 5)
        snd = Replace(snd, " ", "|")
        mssg = Left(aMsg(1), Len(aMsg(1)) - 0)
        send.sndmsg Furre, snd, mssg, Txt
    ElseIf Left(Msg, 6) = "check " Then
    On Error GoTo error
        snd = Right(Msg, Len(Msg) - 6)
        snd = Replace(snd, " ", "|")
        sugchk.chkfur Furre, snd
    ElseIf Left(Msg, 8) = "suggest " Then
    On Error GoTo error
        snd = Right(Msg, Len(Msg) - 8)
        snd = Replace(snd, " ", "|")
        sugchk.sugfur Furre, snd, Txt
    ElseIf Left(Msg, 5) = "card " Then
    On Error GoTo error
        aMsg = Split(Msg, " message ", 2)
        bMsg = Right(aMsg(0), Len(aMsg(0)) - 5)
        cMsg = Split(bMsg, " image ", 2)
        snd = Right(cMsg(0), Len(cMsg(0)) - 0)
        snd = Replace(snd, " ", "|")
        imag = Left(cMsg(1), Len(cMsg(1)) - 0)
        mssg = Left(aMsg(1), Len(aMsg(1)) - 0)
        mssg = Replace(mssg, " ", "%20")
        send.sndcard Furre, snd, mssg, imag, Txt
    ElseIf Left(Msg, 8) = "forward " Then
    On Error GoTo error
        snd = Right(Msg, Len(Msg) - 8)
        snd = Replace(snd, " ", "|")
        Forword.doforward Furre, snd, Txt
     ElseIf Left(Msg, 1) = "#" Then
     On Error GoTo error
        Msg = Right(Msg, Len(Msg) - 1)
        whispers.DoHelp Furre, Msg, Txt
     ElseIf Left(Msg, 1) = "@" Then
     On Error GoTo error
        Msg = Right(Msg, Len(Msg) - 1)
        whispers.DoAdmin Furre, Msg, Txt
    Else
        whispers.DoWhisper Furre, Msg, Txt
    End If

Else
    If Left(Txt, 3) = "([ " And Right(Txt, 10) = " to you. ]" Then
    tmsg = Split(Txt, " whispers, " & Chr(34), 2)
    Furre = Right(tmsg(0), Len(tmsg(0)) - 3)
    frmBot.sckFurc.SendData "wh " & Furre & " Im currently offline please try again later." & vbLf
    End If
End If
End If 'chkWhisp
Exit Sub
error:
    frmBot.sckFurc.SendData "wh " & Furre & " Im Sorry, " & Chr(34) & Msg & Chr(34) & " is not a valid command. Whisper me *help to learn how to use my service." & vbLf
    Resume stoptrying
stoptrying:
End Sub
