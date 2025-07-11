Attribute VB_Name = "join"
Sub dojoin(Furre, Txt)
On Error Resume Next


frmBot.members.Recordset.MoveFirst
Do Until frmBot.txtName.Text = Furre Or frmBot.members.Recordset.EOF
frmBot.members.Recordset.MoveNext
Loop

If frmBot.txtName.Text <> Furre Then
frmBot.members.Recordset.AddNew
frmBot.txtName.Text = Furre
frmBot.txtDate.Text = Now
frmBot.txtStatus.Text = "1"
frmBot.sckFurc.SendData "wh " & Furre & " You are now registered with MailSys." & vbLf

Else
frmBot.sckFurc.SendData "wh " & Furre & " You are allready registered with MailSys." & vbLf
End If


End Sub
