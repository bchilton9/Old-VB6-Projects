Attribute VB_Name = "read"
Dim messnum As Integer

Sub doread(Furre, snd)
    messnum = 0
    Open "messages.dat" For Input As #1
    Do
    Input #1, mto, mfrom, mess
    If mto = Furre Then
        messnum = messnum + 1
            If messnum = snd Then
                frmBot.sckFurc.SendData "wh " & Furre & " Message from " & mfrom & ": " & mess & vbLf
            Else
                frmBot.sckFurc.SendData "wh " & Furre & " You dont have a message numbered " & snd & vbLf
            End If
    End If
    Loop Until messnum = snd Or EOF(1)
    Close #1
End Sub
