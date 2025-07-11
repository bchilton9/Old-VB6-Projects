Attribute VB_Name = "delete"
Sub dodelete(Furre, sndr, Txt)
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
frmBot.dbaMess.Recordset.delete
frmBot.sckFurc.SendData "wh " & Furre & " Message deleted." & vbLf
Else
frmBot.sckFurc.SendData "wh " & Furre & " Invalid message number." & vbLf
End If


End If
End Sub
