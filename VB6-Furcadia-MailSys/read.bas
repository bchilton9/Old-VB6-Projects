Attribute VB_Name = "read"
Sub remsg(Furre, sndr As Integer, Txt)
On Error Resume Next
c = 1
frmBot.members.Recordset.MoveFirst
Do Until frmBot.txtName.Text = Furre Or frmBot.members.Recordset.EOF
frmBot.members.Recordset.MoveNext
Loop

If frmBot.txtName.Text <> Furre Then
frmBot.sckFurc.SendData "wh " & Furre & " You are not registered with MailSys." & vbLf

Else

frmBot.dbaMess.Recordset.MoveFirst
Do Until (c = sndr And frmBot.txtTo.Text = Furre) Or frmBot.dbaMess.Recordset.EOF
If frmBot.txtTo.Text = Furre Then
c = c + 1
End If
frmBot.dbaMess.Recordset.MoveNext
Loop

If c = sndr Then
frmBot.sckFurc.SendData "wh " & Furre & " From: " & frmBot.txtFrom.Text & " Message: " & frmBot.txtMess.Text & " Sent: " & frmBot.txtSndDate.Text & vbLf
Else
frmBot.sckFurc.SendData "wh " & Furre & " Invalid message number." & vbLf
End If


End If
End Sub

Sub doread(Furre, Txt)
On Error Resume Next
c = 0
frmBot.members.Recordset.MoveFirst
Do Until frmBot.txtName.Text = Furre Or frmBot.members.Recordset.EOF
frmBot.members.Recordset.MoveNext
Loop

If frmBot.txtName.Text <> Furre Then
frmBot.sckFurc.SendData "wh " & Furre & " You are not registered with MailSys." & vbLf

Else
frmBot.sckFurc.SendData "wh " & Furre & " You have messages from:"

frmBot.dbaMess.Recordset.MoveFirst
Do Until frmBot.dbaMess.Recordset.EOF
If frmBot.txtTo.Text = Furre Then
c = c + 1
frmBot.sckFurc.SendData " [" & frmBot.txtFrom.Text & " - " & c & "]"
End If
frmBot.dbaMess.Recordset.MoveNext
Loop
If c = 0 Then frmBot.sckFurc.SendData " MailBox Empty."
frmBot.sckFurc.SendData vbLf
End If
End Sub
