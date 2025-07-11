Attribute VB_Name = "sugchk"
Sub remove(Furre)
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
    frmBot.sckFurc.SendData "wh " & Furre & " You have been removed from the suggest list. Thank you" & vbLf
End Sub

Sub sugfur(Furre, snd, Txt)
On Error GoTo sugfurerr
Open "C:\mailsys\members.txt" For Input As #1
Do Until (sfName = snd) Or (EOF(1))
Input #1, sfName, smnum, ban
Loop
Close #1
If snd = "" Then
frmBot.sckFurc.SendData "wh " & Furre & " I'm sorry, You must enter a Furry's name. Whisper me *help to learn how to use my service." & vbLf
Else
If sfName = snd Then
frmBot.sckFurc.SendData "wh " & Furre & " " & snd & " is allready a member." & vbLf

Else
Open "C:\mailsys\suggest.txt" For Append As #1
Write #1, snd, Furre
Close #1
        Open "C:\mailsys\suggestq.txt" For Input As #3
        Input #3, qunt
        Close #3
        qunt = qunt + 1
        Open "C:\mailsys\suggestq.txt" For Output As #3
        Write #3, qunt
        Close #3

frmBot.sckFurc.SendData "wh " & Furre & " I will let " & snd & " know that you suggested that he/she join's Mailsys. Thank you." & vbLf
End If
End If

Exit Sub
sugfurerr:
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
    Write #5, "Suggest Error (sugfur), from " & Furre & " with " & Txt & ", [When: " & Date & " at " & Time & "]"
    Close #5
    frmBot.sckFurc.SendData "wh " & Furre & " There has been an error. Whisper me #help to learn how to use my service. If you keep recieveing this error please leave a message for Mys' or e-mail him at mailsys@erenetwork.com." & vbLf
    Resume stoptrying
stoptrying:
End Sub

Sub chkfur(Furre, snd)
Open "C:\mailsys\members.txt" For Input As #1
Do Until (sfName = snd) Or (EOF(1))
Input #1, sfName, smnum, ban
Loop
Close #1
If sfName = snd Then
    frmBot.sckFurc.SendData "wh " & Furre & " " & snd & " is registered." & vbLf
Else
    frmBot.sckFurc.SendData "wh " & Furre & " " & snd & " is not registered." & vbLf
End If

End Sub
