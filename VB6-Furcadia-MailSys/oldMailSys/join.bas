Attribute VB_Name = "join"
Sub dojoin(Furre, Txt)
Open "C:\mailsys\members.txt" For Input As #1
Do Until (fName = Furre) Or (EOF(1))
Input #1, fName, mnum, ban
Loop
Close #1

If fName = Furre Then
On Error GoTo joerr
frmBot.sckFurc.SendData "wh " & Furre & " You are already registered. Whisper me #help to learn how to use my service." & vbLf
Else
Open "C:\mailsys\memnum.txt" For Input As #1
Input #1, nnum
Close #1
num = nnum + 1
Open "C:\mailsys\memnum.txt" For Output As #1
Write #1, num
Close #1


           Open "C:\mailsys\suggestq.txt" For Input As #3
            Input #3, qun
            Close #3
            If qun >= 1 Then
                Open "C:\mailsys\suggest.txt" For Input As #1
                Open "C:\mailsys\suggesta.txt" For Output As #2
                Do Until EOF(1)
                Input #1, frm, snd
                If frm = Furre Then
                    qun = qun - 1
                    Open "C:\mailsys\suggestq.txt" For Output As #3
                    Write #3, qun
                    Close #3
                Else
                    Write #2, frm, snd
                End If 'end if frm = Furre
                Loop
                
                Close #1, #2
            End If 'end if qun >= 1
            Open "C:\mailsys\suggesta.txt" For Input As #1
            Open "C:\mailsys\suggest.txt" For Output As #2
            Do Until EOF(1)
                Input #1, frm, snd
                Write #2, frm, snd
                Loop
            Close #1, #2
            Open "C:\mailsys\suggesta.txt" For Output As #1
            Write #1, ""
            Close #1



Open "C:\mailsys\memfiles\" & num & ".txt" For Output As #1
Write #1, "1", "mailsys", "Welcome to MailSys. The all new Messageing system for the Furries. Place it in your desc and tell your friends. Let them leave you a message when your off line. [Sent: " & Date & " at " & Time & " MST]"
Close #1
Open "C:\mailsys\memfiles\" & num & "q.txt" For Output As #3
Write #3, 1
Close #3
Open "C:\mailsys\members.txt" For Append As #1
Write #1, Furre, num, 0
Close #1
frmBot.sckFurc.SendData "wh " & Furre & " You are now regiastered with Mailsys. Whisper me #help to learn how to use my service." & vbLf
End If
Exit Sub
joerr:
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
    Write #5, "Join error (dojoin), from " & Furre & " with " & Txt & ", [When: " & Date & " at " & Time & "]"
    Close #5
    frmBot.sckFurc.SendData "wh " & Furre & " There has been an error. Whisper me #help to learn how to use my service. If you keep recieveing this error please leave a message for Mys' or e-mail him at mailsys@erenetwork.com." & vbLf
    Resume stoptrying
stoptrying:
End Sub
