Attribute VB_Name = "send"
Sub sndcard(Furre, snd, mssg, imag, Txt)
On Error Resume Next
mssg = Replace(mssg, Chr(34), "")
mssg = "http://www.erenetwork.com/mailsys/images/card.cgi?mode=1&from=" & Furre & "&image=" & imag & "&mess=" & mssg & "                                                                  You have a SysCard wateing for you. Press F8 now to view.  "

frmBot.members.Recordset.MoveFirst
Do Until frmBot.txtName.Text = Furre Or frmBot.members.Recordset.EOF
frmBot.members.Recordset.MoveNext
Loop

If frmBot.txtName.Text <> Furre Then
frmBot.sckFurc.SendData "wh " & Furre & " You are not registered with MailSys." & vbLf
Else

frmBot.members.Recordset.MoveFirst
Do Until frmBot.txtName.Text = snd Or frmBot.members.Recordset.EOF
frmBot.members.Recordset.MoveNext
Loop

If frmBot.txtName.Text <> snd Then
frmBot.sckFurc.SendData "wh " & Furre & " " & snd & " is not registered with MailSys." & vbLf
Else


frmBot.dbaMess.Recordset.AddNew
frmBot.txtTo.Text = snd
frmBot.txtFrom.Text = "SysCard"
frmBot.txtMess.Text = mssg
frmBot.txtSndDate = Now
frmBot.sckFurc.SendData "wh " & Furre & " Message sent to " & snd & "." & vbLf

Open "sent.txt" For Input As #3
Input #3, sent
Close #3
sent = sent + 1
frmBot.txtsent.Text = sent
Open "sent.txt" For Output As #3
Write #3, sent
Close #3


End If
End If

End Sub

Sub sndmsg(Furre, snd, mssg, Txt)
On Error Resume Next
mssg = Replace(mssg, Chr(34), "")
frmBot.members.Recordset.MoveFirst
Do Until frmBot.txtName.Text = Furre Or frmBot.members.Recordset.EOF
frmBot.members.Recordset.MoveNext
Loop

If frmBot.txtName.Text <> Furre Then
frmBot.sckFurc.SendData "wh " & Furre & " You are not registered with MailSys." & vbLf
Else

frmBot.members.Recordset.MoveFirst
Do Until frmBot.txtName.Text = snd Or frmBot.members.Recordset.EOF
frmBot.members.Recordset.MoveNext
Loop

If frmBot.txtName.Text <> snd Then
frmBot.sckFurc.SendData "wh " & Furre & " " & snd & " is not registered with MailSys." & vbLf
Else


frmBot.dbaMess.Recordset.AddNew
frmBot.txtTo.Text = snd
frmBot.txtFrom.Text = Furre
frmBot.txtMess.Text = mssg
frmBot.txtSndDate = Now
frmBot.sckFurc.SendData "wh " & Furre & " Message sent to " & snd & "." & vbLf

Open "sent.txt" For Input As #3
Input #3, sent
Close #3
sent = sent + 1
frmBot.txtsent.Text = sent
Open "sent.txt" For Output As #3
Write #3, sent
Close #3

End If
End If
End Sub
