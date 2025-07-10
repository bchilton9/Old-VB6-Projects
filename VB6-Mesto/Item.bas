Attribute VB_Name = "Item"
Sub additem(num, furre)

frmBot.mbr.Recordset.MoveFirst
Do Until frmBot.txtName.Text = furre Or frmBot.mbr.Recordset.EOF
frmBot.mbr.Recordset.MoveNext
Loop

If frmBot.txtName.Text = furre Then

frmBot.mbr.Recordset.Edit

If frmBot.txtInv1.Text = num Then
frmBot.txtInv1q.Text = frmBot.txtInv1q.Text + 1
ElseIf frmBot.txtInv1.Text = 0 Then
frmBot.txtInv1.Text = num
frmBot.txtInv1q.Text = frmBot.txtInv1q.Text + 1

ElseIf frmBot.txtInv2.Text = num Then
frmBot.txtInv2q.Text = frmBot.txtInv2q.Text + 1
ElseIf frmBot.txtInv2.Text = 0 Then
frmBot.txtInv2.Text = num
frmBot.txtInv2q.Text = frmBot.txtInv2q.Text + 1

ElseIf frmBot.txtInv3.Text = num Then
frmBot.txtInv3q.Text = frmBot.txtInv3q.Text + 1
ElseIf frmBot.txtInv3.Text = 0 Then
frmBot.txtInv3.Text = num
frmBot.txtInv3q.Text = frmBot.txtInv3q.Text + 1

ElseIf frmBot.txtInv4.Text = num Then
frmBot.txtInv4q.Text = frmBot.txtInv4q.Text + 1
ElseIf frmBot.txtInv4.Text = 0 Then
frmBot.txtInv4.Text = num
frmBot.txtInv4q.Text = frmBot.txtInv4q.Text + 1

ElseIf frmBot.txtInv5.Text = num Then
frmBot.txtInv5q.Text = frmBot.txtInv5q.Text + 1
ElseIf frmBot.txtInv5.Text = 0 Then
frmBot.txtInv5.Text = num
frmBot.txtInv5q.Text = frmBot.txtInv5q.Text + 1

ElseIf frmBot.txtInv6.Text = num Then
frmBot.txtInv6q.Text = frmBot.txtInv6q.Text + 1
ElseIf frmBot.txtInv6.Text = 0 Then
frmBot.txtInv6.Text = num
frmBot.txtInv6q.Text = frmBot.txtInv6q.Text + 1
frmBot.sckFurc.SendData "wh " & furre & " Your bag is full! You must sell some items befor geting any more or you may lose them out of your bag!" & vbLf

Else

frmBot.sckFurc.SendData "wh " & furre & " Your bag is full! You lost the item out of your bag." & vbLf

End If

frmBot.mbr.Recordset.Update

Else
frmBot.sckFurc.SendData "wh " & furre & " Im sorry. You are not a member. The item will be lost! Please Join Mesto City to prevent this from happening agine!" & vbLf
End If


End Sub

Sub addgold(furre, amount)

frmBot.mbr.Recordset.MoveFirst
Do Until frmBot.txtName.Text = furre Or frmBot.mbr.Recordset.EOF
frmBot.mbr.Recordset.MoveNext
Loop

If frmBot.txtName.Text = furre Then

       frmBot.mbr.Recordset.Edit
       frmBot.txtGold.Text = frmBot.txtGold.Text + amount
       frmBot.mbr.Recordset.Update

End If

End Sub

Sub subgold(furre, amount)

frmBot.mbr.Recordset.MoveFirst
Do Until frmBot.txtName.Text = furre Or frmBot.mbr.Recordset.EOF
frmBot.mbr.Recordset.MoveNext
Loop

If frmBot.txtName.Text = furre Then

       frmBot.mbr.Recordset.Edit
       frmBot.txtGold.Text = frmBot.txtGold.Text - amount
       frmBot.mbr.Recordset.Update

End If

End Sub
