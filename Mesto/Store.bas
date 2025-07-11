Attribute VB_Name = "Store"
Sub dostore(furre, storenum, pocket)

Dim addgold As Integer

frmBot.mbr.Recordset.MoveFirst
Do Until frmBot.txtName.Text = furre Or frmBot.mbr.Recordset.EOF
frmBot.mbr.Recordset.MoveNext
Loop

If frmBot.txtName.Text = furre Then

If pocket = 1 Then
   If frmBot.txtInv1.Text = 0 Then
       frmBot.sckFurc.SendData Chr(34) & "store " & storenum & " no item" & vbLf
    Else
       frmBot.items.Recordset.MoveFirst
       Do Until frmBot.txtInv1.Text = frmBot.txtItemID.Text Or frmBot.items.Recordset.EOF
       frmBot.items.Recordset.MoveNext
       Loop
       addgold = frmBot.txtItemSell.Text * frmBot.txtInv1q.Text
       frmBot.mbr.Recordset.Edit
       frmBot.txtInv1.Text = 0
       frmBot.txtInv1q.Text = 0
       frmBot.txtGold.Text = frmBot.txtGold.Text + addgold
       frmBot.mbr.Recordset.Update
       frmBot.sckFurc.SendData Chr(34) & "store " & storenum & " sold" & vbLf
   End If
End If
If pocket = 2 Then
   If frmBot.txtInv2.Text = 0 Then
       frmBot.sckFurc.SendData Chr(34) & "store " & storenum & " no item" & vbLf
    Else
       frmBot.items.Recordset.MoveFirst
       Do Until frmBot.txtInv2.Text = frmBot.txtItemID.Text Or frmBot.items.Recordset.EOF
       frmBot.items.Recordset.MoveNext
       Loop
       addgold = frmBot.txtItemSell.Text * frmBot.txtInv2q.Text
       frmBot.mbr.Recordset.Edit
       frmBot.txtInv2.Text = 0
       frmBot.txtInv2q.Text = 0
       frmBot.txtGold.Text = frmBot.txtGold.Text + addgold
       frmBot.mbr.Recordset.Update
       frmBot.sckFurc.SendData Chr(34) & "store " & storenum & " sold" & vbLf
   End If
End If
If pocket = 3 Then
   If frmBot.txtInv3.Text = 0 Then
       frmBot.sckFurc.SendData Chr(34) & "store " & storenum & " no item" & vbLf
    Else
       frmBot.items.Recordset.MoveFirst
       Do Until frmBot.txtInv3.Text = frmBot.txtItemID.Text Or frmBot.items.Recordset.EOF
       frmBot.items.Recordset.MoveNext
       Loop
       addgold = frmBot.txtItemSell.Text * frmBot.txtInv3q.Text
       frmBot.mbr.Recordset.Edit
       frmBot.txtInv3.Text = 0
       frmBot.txtInv3q.Text = 0
       frmBot.txtGold.Text = frmBot.txtGold.Text + addgold
       frmBot.mbr.Recordset.Update
       frmBot.sckFurc.SendData Chr(34) & "store " & storenum & " sold" & vbLf
   End If
End If
If pocket = 4 Then
   If frmBot.txtInv4.Text = 0 Then
       frmBot.sckFurc.SendData Chr(34) & "store " & storenum & " no item" & vbLf
    Else
       frmBot.items.Recordset.MoveFirst
       Do Until frmBot.txtInv4.Text = frmBot.txtItemID.Text Or frmBot.items.Recordset.EOF
       frmBot.items.Recordset.MoveNext
       Loop
       addgold = frmBot.txtItemSell.Text * frmBot.txtInv4q.Text
       frmBot.mbr.Recordset.Edit
       frmBot.txtInv4.Text = 0
       frmBot.txtInv4q.Text = 0
       frmBot.txtGold.Text = frmBot.txtGold.Text + addgold
       frmBot.mbr.Recordset.Update
       frmBot.sckFurc.SendData Chr(34) & "store " & storenum & " sold" & vbLf
   End If
End If
If pocket = 5 Then
   If frmBot.txtInv5.Text = 0 Then
       frmBot.sckFurc.SendData Chr(34) & "store " & storenum & " no item" & vbLf
    Else
       frmBot.items.Recordset.MoveFirst
       Do Until frmBot.txtInv5.Text = frmBot.txtItemID.Text Or frmBot.items.Recordset.EOF
       frmBot.items.Recordset.MoveNext
       Loop
       addgold = frmBot.txtItemSell.Text * frmBot.txtInv5q.Text
       frmBot.mbr.Recordset.Edit
       frmBot.txtInv5.Text = 0
       frmBot.txtInv5q.Text = 0
       frmBot.txtGold.Text = frmBot.txtGold.Text + addgold
       frmBot.mbr.Recordset.Update
       frmBot.sckFurc.SendData Chr(34) & "store " & storenum & " sold" & vbLf
   End If
End If
If pocket = 6 Then
   If frmBot.txtInv6.Text = 0 Then
       frmBot.sckFurc.SendData Chr(34) & "store " & storenum & " no item" & vbLf
    Else
       frmBot.items.Recordset.MoveFirst
       Do Until frmBot.txtInv6.Text = frmBot.txtItemID.Text Or frmBot.items.Recordset.EOF
       frmBot.items.Recordset.MoveNext
       Loop
       addgold = frmBot.txtItemSell.Text * frmBot.txtInv6q.Text
       frmBot.mbr.Recordset.Edit
       frmBot.txtInv6.Text = 0
       frmBot.txtInv6q.Text = 0
       frmBot.txtGold.Text = frmBot.txtGold.Text + addgold
       frmBot.mbr.Recordset.Update
       frmBot.sckFurc.SendData Chr(34) & "store " & storenum & " sold" & vbLf
   End If
End If

Else
       frmBot.sckFurc.SendData Chr(34) & "store " & storenum & " nomember" & vbLf
End If

End Sub


Sub dobuy(furre, storenum, itemnum)

Dim price As Integer
Dim mygold As Integer

frmBot.mbr.Recordset.MoveFirst
Do Until frmBot.txtName.Text = furre Or frmBot.mbr.Recordset.EOF
frmBot.mbr.Recordset.MoveNext
Loop

If frmBot.txtName.Text = furre Then

If frmBot.txtInv1.Text <> 0 And frmBot.txtInv2.Text <> 0 And frmBot.txtInv3.Text <> 0 And frmBot.txtInv4.Text <> 0 And frmBot.txtInv5.Text <> 0 And frmBot.txtInv6.Text <> 0 And frmBot.txtInv1.Text <> itemnum And frmBot.txtInv2.Text <> itemnum And frmBot.txtInv3.Text <> itemnum And frmBot.txtInv4.Text <> itemnum And frmBot.txtInv5.Text <> itemnum And frmBot.txtInv6.Text <> itemnum Then
       frmBot.sckFurc.SendData Chr(34) & "store " & storenum & " full" & vbLf
Else

       frmBot.items.Recordset.MoveFirst
       Do Until frmBot.txtItemID.Text = itemnum Or frmBot.items.Recordset.EOF
       frmBot.items.Recordset.MoveNext
       Loop

price = frmBot.txtItemBuy.Text
mygold = frmBot.txtGold.Text

If price <= mygold Then
       subgold furre, price
       additem itemnum, furre
       frmBot.sckFurc.SendData Chr(34) & "store " & storenum & " bought" & vbLf
Else
       frmBot.sckFurc.SendData Chr(34) & "store " & storenum & " no gold" & vbLf
End If
End If

Else
       frmBot.sckFurc.SendData Chr(34) & "store " & storenum & " nomember" & vbLf
End If

End Sub
