Attribute VB_Name = "read"
Sub remsg(Furre, sndr, Txt)
Open "C:\mailsys\members.txt" For Input As #1
Do Until (fName = Furre) Or (EOF(1))
Input #1, fName, mnum, ban
Loop
Close #1
If fName = Furre Then
On Error GoTo reerr
Open "C:\mailsys\memfiles\" & mnum & ".txt" For Input As #1
Do Until (dnum = sndr) Or (EOF(1))
Input #1, dnum, ser, mssg
Loop
Close #1
If dnum = sndr Then
    frmBot.sckFurc.SendData "wh " & Furre & " Message from: " & ser & " - " & mssg & vbLf
Else
    frmBot.sckFurc.SendData "wh " & Furre & "  Im Sorry, " & Chr(34) & sndr & Chr(34) & " is not a valid message number. Whisper me #help to learn how to use my service." & vbLf
End If

Else
    frmBot.sckFurc.SendData "wh " & Furre & " You are not registered with me." & vbLf
End If
Exit Sub
reerr:
    Close #1
    frmBot.txterr = frmBot.txterr + 1
    Open "C:\mailsys\errorq.txt" For Input As #5
    Input #5, num
    Close #5
    num = num + 1
    Open "C:\mailsys\errorq.txt" For Output As #5
    Write #5, num
    Close #5
    Open "C:\mailsys\errorlog.txt" For Append As #5
    Write #5, "Read Message error (remsg), from " & Furre & " with " & Txt & ", [When: " & Date & " at " & Time & "]"
    Close #5
    frmBot.sckFurc.SendData "wh " & Furre & " There has been an error. Whisper me #help to learn how to use my service. If you keep recieveing this error try sending yourself a message. If that dose not work send a message to MailSys telling what happened to it can be repaired." & vbLf
    Resume stoptrying
stoptrying:
End Sub

Sub doproxy(Furre, Txt)
read:
Open "C:\mailsys\members.txt" For Input As #2
Do Until (fName = Furre) Or (EOF(2))
Input #2, fName, mnum, ban
Loop
Close #2
If fName = Furre Then
On Error GoTo reerr
        Open "C:\mailsys\memfiles\" & mnum & "q.txt" For Input As #3
        Input #3, qunt
        Close #3
        If qunt <> 0 Then
            On Error GoTo reerr
            frmBot.sckFurc.SendData "wh " & Furre & " You have messages from:"
            Open "C:\mailsys\memfiles\" & mnum & ".txt" For Input As #1
            Do Until (EOF(1))
            Input #1, dnum, ser, mssg
            frmBot.sckFurc.SendData " [" & ser & " - #" & dnum & "]"
            Loop
            Close #1
            frmBot.sckFurc.SendData vbLf
        End If 'end if qount = 0
Else
    frmBot.sckFurc.SendData "wh " & Furre & " You are not registered with me. Whisper me *help to learn how to use my service." & vbLf
End If

Exit Sub
reerr:
    Close #1, #2, #3
    frmBot.txterr = frmBot.txterr + 1
    Open "C:\mailsys\errorq.txt" For Input As #5
    Input #5, num
    Close #5
    num = num + 1
    Open "C:\mailsys\errorq.txt" For Output As #5
    Write #5, num
    Close #5
    Open "C:\mailsys\errorlog.txt" For Append As #5
    Write #5, "Read error (doread), from " & Furre & " with " & Txt & ", [When: " & Date & " at " & Time & "]"
    Close #5
    'frmBot.sckFurc.SendData "wh " & Furre & " There has been an error. Whisper me #help to learn how to use my service. If you keep recieveing this error try sending yourself a message. If that dose not work send a message to MailSys or email at MailSys@erenetwork.com telling what happened to it can be repaired." & vbLf
    Open "C:\mailsys\memfiles\" & mnum & ".txt" For Output As #1
    Write #1, "1", "MailSys", "Message file was courupted. File has been restored. Sorry for the inconvienice."
    Close #1
    Open "C:\mailsys\memfiles\" & mnum & "q.txt" For Output As #1
    Write #1, 1
    Close #1
    snd = 1
    Resume read
stoptrying:
End Sub

Sub doread(Furre, Txt)
read:
Open "C:\mailsys\members.txt" For Input As #2
Do Until (fName = Furre) Or (EOF(2))
Input #2, fName, mnum, ban
Loop
Close #2
If fName = Furre Then
On Error GoTo reerr
        Open "C:\mailsys\memfiles\" & mnum & "q.txt" For Input As #3
        Input #3, qunt
        Close #3
        If qunt = 0 Then
            frmBot.sckFurc.SendData "wh " & Furre & " You dont have any messages." & vbLf
        
        Else
            On Error GoTo reerr
            frmBot.sckFurc.SendData "wh " & Furre & " You have messages from:"
            Open "C:\mailsys\memfiles\" & mnum & ".txt" For Input As #1
            Do Until (EOF(1))
            Input #1, dnum, ser, mssg
            frmBot.sckFurc.SendData " [" & ser & " - #" & dnum & "]"
            Loop
            Close #1
            frmBot.sckFurc.SendData vbLf
        End If 'end if qount = 0
Else
    frmBot.sckFurc.SendData "wh " & Furre & " You are not registered with me. Whisper me *help to learn how to use my service." & vbLf
End If

Exit Sub
reerr:
    Close #1, #2, #3
    frmBot.txterr = frmBot.txterr + 1
    Open "C:\mailsys\errorq.txt" For Input As #5
    Input #5, num
    Close #5
    num = num + 1
    Open "C:\mailsys\errorq.txt" For Output As #5
    Write #5, num
    Close #5
    Open "C:\mailsys\errorlog.txt" For Append As #5
    Write #5, "Read error (doread), from " & Furre & " with " & Txt & ", [When: " & Date & " at " & Time & "]"
    Close #5
    'frmBot.sckFurc.SendData "wh " & Furre & " There has been an error. Whisper me #help to learn how to use my service. If you keep recieveing this error try sending yourself a message. If that dose not work send a message to MailSys or email at MailSys@erenetwork.com telling what happened to it can be repaired." & vbLf
    Open "C:\mailsys\memfiles\" & mnum & ".txt" For Output As #1
    Write #1, "1", "MailSys", "Message file was courupted. File has been restored. Sorry for the inconvienice."
    Close #1
    Open "C:\mailsys\memfiles\" & mnum & "q.txt" For Output As #1
    Write #1, 1
    Close #1
    snd = 1
    Resume read
stoptrying:
End Sub
