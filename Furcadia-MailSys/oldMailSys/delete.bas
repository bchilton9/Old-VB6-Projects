Attribute VB_Name = "delete"
Sub dodelete(Furre, snd, Txt)
delete:
Dim dnum As String

Open "C:\mailsys\members.txt" For Input As #1
Do Until (fName = Furre) Or (EOF(1))
Input #1, fName, mnum, ban
Loop
Close #1
If fName = Furre Then
On Error GoTo deerr
Open "C:\mailsys\memfiles\" & mnum & "q.txt" For Input As #1
Input #1, qun
Close #1

If qun = "0" Then
    frmBot.sckFurc.SendData "wh " & Furre & " No messages to delete." & vbLf
Else
del = 0
Open "C:\mailsys\memfiles\" & mnum & ".txt" For Input As #1
Open "C:\mailsys\memfiles\" & mnum & "a.txt" For Output As #2
Do Until EOF(1)
Input #1, dnum, frm, mssg
If dnum = snd And del = 0 Then
        frmBot.sckFurc.SendData "wh " & Furre & " Message #" & snd & " has been deleted." & vbLf
        nqun = qun - 1
        Open "C:\mailsys\memfiles\" & mnum & "q.txt" For Output As #3
        Write #3, nqun
        Close #3
        del = 1
Else
    If del = 1 Then dnum = dnum - 1
    Write #2, dnum, frm, mssg
End If
Loop
Close #1, #2

Open "C:\mailsys\memfiles\" & mnum & "a.txt" For Input As #1
Open "C:\mailsys\memfiles\" & mnum & ".txt" For Output As #2
Do Until EOF(1)
    Input #1, dnum, frm, mssg
    Write #2, dnum, frm, mssg
Loop
Close #1, #2
End If
Open "C:\mailsys\memfiles\" & mnum & "a.txt" For Output As #1
Write #1, ""
Close #1


If del = 0 Then frmBot.sckFurc.SendData "wh " & Furre & "  Im Sorry, " & Chr(34) & snd & Chr(34) & " is not a valid message number. Whisper me #help to learn how to use my service." & vbLf
Else
frmBot.sckFurc.SendData "wh " & Furre & " You are not registered with me. Whisper me *help to learn how to use my service." & vbLf
End If

Exit Sub
deerr:
    Close #1, #2
    frmBot.txterr = frmBot.txterr + 1
    Open "C:\mailsys\errorq.txt" For Input As #5
    Input #5, num
    Close #5
    num = num + 1
    Open "C:\mailsys\errorq.txt" For Output As #5
    Write #5, num
    Close #5
    Open "C:\mailsys\errorlog.txt" For Append As #5
    Write #5, "Delete error (dodelete), from " & Furre & " with " & Txt & ", [When: " & Date & " at " & Time & "]"
    Close #5
    'frmBot.sckFurc.SendData "wh " & Furre & " There has been an error. Whisper me *help to learn how to use my service. If you keep recieveing this error please leave a message for Mys' or e-mail him at mailsys@erenetwork.com." & vbLf
    Open "C:\mailsys\memfiles\" & mnum & ".txt" For Output As #1
    Write #1, "1", "MailSys", "Message file was courupted. File has been restored. Sorry for the inconvienice."
    Close #1
    Open "C:\mailsys\memfiles\" & mnum & "q.txt" For Output As #1
    Write #1, 1
    Close #1
    Resume delete
stoptrying:
End Sub
