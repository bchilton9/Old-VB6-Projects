Attribute VB_Name = "send"
Sub sndcard(Furre, snd, mssg, imag, Txt)
mssg = Replace(mssg, Chr(34), "")
Open "C:\mailsys\members.txt" For Input As #1
Do Until (fName = Furre) Or (EOF(1))
Input #1, fName, mnum, ban
Loop
Close #1
If fName = Furre Then
On Error GoTo snerr

Open "C:\mailsys\forward.txt" For Input As #1
Do Until (fsfName = snd) Or (EOF(1))
Input #1, fsfName, fsmnum
Loop
Close #1
If fsfName = snd Then snd = fsmnum

Open "C:\mailsys\members.txt" For Input As #1
Do Until (sfName = snd) Or (EOF(1))
Input #1, sfName, smnum, ban
Loop
Close #1
If sfName = snd Then
Dim qun As String
    Open "C:\mailsys\memfiles\" & smnum & "q.txt" For Input As #3
    Input #3, qun
    Close #3
    qun = qun + 1
    Open "C:\mailsys\memfiles\" & smnum & "q.txt" For Output As #3
    Write #3, qun
    Close #3
    
    Open "C:\mailsys\sent.txt" For Input As #3
    Input #3, sent
    Close #3
    sent = sent + 1
    Open "C:\mailsys\sent.txt" For Output As #3
    Write #3, sent
    Close #3
    
    Open "C:\mailsys\memfiles\" & smnum & ".txt" For Append As #1
    Write #1, qun, "SysCard", "http://www.erenetwork.com/mailsys/images/card.cgi?mode=1&from=" & Furre & "&image=" & imag & "&mess=" & mssg & "                                                                  You have a SysCard wateing for you. Press F8 now to view.  " & " [Sent: " & Date & " at " & Time & " MST]"
    Close #1
    frmBot.sckFurc.SendData "wh " & Furre & " Message sent to: " & snd & vbLf
Else
    frmBot.sckFurc.SendData "wh " & Furre & " " & snd & " is not registered." & vbLf
End If
Else
frmBot.sckFurc.SendData "wh " & Furre & " You are not registered with me. Whisper me *help to learn how to use my service." & vbLf
End If
Exit Sub
snerr:
    Close #1, #3
    frmBot.txterr = frmBot.txterr + 1
    Open "C:\mailsys\errorq.txt" For Input As #5
    Input #5, num
    Close #5
    num = num + 1
    Open "C:\mailsys\errorq.txt" For Output As #5
    Write #5, num
    Close #5
    Open "C:\mailsys\errorlog.txt" For Append As #5
    Write #5, "Card error (sndcard), from " & Furre & " with " & Txt & ", [When: " & Date & " at " & Time & "]"
    Close #5
    frmBot.sckFurc.SendData "wh " & Furre & " Im Sorry, for some reason your message could not be delivered. This error has been loged and will be fixed as soon as possible." & vbLf
    Resume stoptrying
stoptrying:
End Sub

Sub sndmsg(Furre, snd, mssg, Txt)
mssg = Replace(mssg, Chr(34), "")
Open "C:\mailsys\members.txt" For Input As #1
Do Until (fName = Furre) Or (EOF(1))
Input #1, fName, mnum, ban
Loop
Close #1
If fName = Furre Then
On Error GoTo snerr

Open "C:\mailsys\forward.txt" For Input As #1
Do Until (fsfName = snd) Or (EOF(1))
Input #1, fsfName, fsmnum
Loop
Close #1
If fsfName = snd Then snd = fsmnum


Open "C:\mailsys\members.txt" For Input As #1
Input #1, sfName, smnum, ban
Do Until (sfName = snd) Or (EOF(1))
Input #1, sfName, smnum, ban
Loop
Close #1
If sfName = snd Then
Dim qun As String
    Open "C:\mailsys\memfiles\" & smnum & "q.txt" For Input As #3
    Input #3, qun
    Close #3
    qun = qun + 1
    Open "C:\mailsys\memfiles\" & smnum & "q.txt" For Output As #3
    Write #3, qun
    Close #3
    
    Open "C:\mailsys\sent.txt" For Input As #3
    Input #3, sent
    Close #3
    sent = sent + 1
    Open "C:\mailsys\sent.txt" For Output As #3
    Write #3, sent
    Close #3
    
    Open "C:\mailsys\memfiles\" & smnum & ".txt" For Append As #1
    Write #1, qun, Furre, mssg & " [Sent: " & Date & " at " & Time & " MST]"
    Close #1
    frmBot.sckFurc.SendData "wh " & Furre & " Message sent to: " & snd & vbLf
Else
    frmBot.sckFurc.SendData "wh " & Furre & " " & snd & " is not registered." & vbLf
End If
Else
frmBot.sckFurc.SendData "wh " & Furre & " You are not registered with me. Whisper me *help to learn how to use my service." & vbLf
End If
Exit Sub
snerr:
    Close #1, #3
    frmBot.txterr = frmBot.txterr + 1
    Open "C:\mailsys\errorq.txt" For Input As #5
    Input #5, num
    Close #5
    num = num + 1
    Open "C:\mailsys\errorq.txt" For Output As #5
    Write #5, num
    Close #5
    Open "C:\mailsys\errorlog.txt" For Append As #5
    Write #5, "Send error (sndmsg), from " & Furre & " with " & Txt & ", [When: " & Date & " at " & Time & "]"
    Close #5
    frmBot.sckFurc.SendData "wh " & Furre & " Im Sorry, for some reason your message could not be delivered. This error has been loged and will be fixed as soon as possible." & vbLf
    Resume stoptrying
stoptrying:
End Sub
