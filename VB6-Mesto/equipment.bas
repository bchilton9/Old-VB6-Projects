Attribute VB_Name = "equipment"
Sub equipitem(furre, pocket)

Dim myhp As Integer
Dim mymana As Integer
Dim mydmg As Integer
Dim addhp As Integer
Dim addmana As Integer
Dim adddmg As Integer

frmBot.mbr.Recordset.MoveFirst
Do Until frmBot.txtName.Text = furre Or frmBot.mbr.Recordset.EOF
frmBot.mbr.Recordset.MoveNext
Loop

If frmBot.txtName.Text = furre Then

If pocket = 1 Then
txtInv = frmBot.txtInv1.Text
txtInvQ = frmBot.txtInv1q.Text
ElseIf pocket = 2 Then
txtInv = frmBot.txtInv2.Text
txtInvQ = frmBot.txtInv2q.Text
ElseIf pocket = 3 Then
txtInv = frmBot.txtInv3.Text
txtInvQ = frmBot.txtInv3q.Text
ElseIf pocket = 4 Then
txtInv = frmBot.txtInv4.Text
txtInvQ = frmBot.txtInv4q.Text
ElseIf pocket = 5 Then
txtInv = frmBot.txtInv5.Text
txtInvQ = frmBot.txtInv5q.Text
ElseIf pocket = 6 Then
txtInv = frmBot.txtInv6.Text
txtInvQ = frmBot.txtInv6q.Text
End If

   If txtInv = 0 Then
       frmBot.sckFurc.SendData "wh " & furre & " There is no items in that pocket." & vbLf
    Else
       frmBot.items.Recordset.MoveFirst
       Do Until txtInv = frmBot.txtItemID.Text Or frmBot.items.Recordset.EOF
       frmBot.items.Recordset.MoveNext
       Loop


       If frmBot.txtItemType.Text = 0 Then
        frmBot.sckFurc.SendData "wh " & furre & " " & frmBot.txtItemName.Text & " is not equipable." & vbLf
       
       Else
          If frmBot.txtWeapon.Text <> 0 And frmBot.txtItemType.Text = 1 Then
             frmBot.sckFurc.SendData "wh " & furre & " You already have a weapon equiped." & vbLf
          ElseIf frmBot.txtEquip1.Text <> 0 And frmBot.txtItemType.Text = 2 Then
             frmBot.sckFurc.SendData "wh " & furre & " You already have a Chest item equiped." & vbLf
          ElseIf frmBot.txtEquip2.Text <> 0 And frmBot.txtItemType.Text = 3 Then
             frmBot.sckFurc.SendData "wh " & furre & " You already have a Leg item equiped." & vbLf
          ElseIf frmBot.txtEquip3.Text <> 0 And frmBot.txtItemType.Text = 4 Then
             frmBot.sckFurc.SendData "wh " & furre & " You already have a Hand item equiped." & vbLf
          ElseIf frmBot.txtEquip4.Text <> 0 And frmBot.txtItemType.Text = 5 Then
             frmBot.sckFurc.SendData "wh " & furre & " You already have a Foot item equiped." & vbLf
             
          Else
          

          myqty = txtInvQ
          
myhp = frmBot.txtHp.Text
mymana = frmBot.txtMana.Text
mydmg = frmBot.txtDmg.Text
addhp = frmBot.txtItemHP.Text
addmana = frmBot.txtItemMana.Text
adddmg = frmBot.txtItemDmg.Text
          
       frmBot.mbr.Recordset.Edit
       
       frmBot.txtHp.Text = myhp + addhp
       frmBot.txtHPLeft.Text = myhp + addhp
       frmBot.txtMana.Text = mymana + addmana
       frmBot.txtManaLeft.Text = mymana + addmana
       frmBot.txtDmg.Text = mydmg + adddmg

       If myqty > 1 Then
       If pocket = 1 Then
       frmBot.txtInv1q.Text = frmBot.txtInv1q.Text - 1
       ElseIf pocket = 2 Then
       frmBot.txtInv2q.Text = frmBot.txtInv2q.Text - 1
       ElseIf pocket = 3 Then
       frmBot.txtInv3q.Text = frmBot.txtInv3q.Text - 1
       ElseIf pocket = 4 Then
       frmBot.txtInv4q.Text = frmBot.txtInv4q.Text - 1
       ElseIf pocket = 5 Then
       frmBot.txtInv5q.Text = frmBot.txtInv5q.Text - 1
       ElseIf pocket = 6 Then
       frmBot.txtInv6q.Text = frmBot.txtInv6q.Text - 1
       End If
       
       Else
       
       If pocket = 1 Then
       frmBot.txtInv1.Text = 0
       frmBot.txtInv1q.Text = 0
       ElseIf pocket = 2 Then
       frmBot.txtInv2.Text = 0
       frmBot.txtInv2q.Text = 0
       ElseIf pocket = 3 Then
       frmBot.txtInv3.Text = 0
       frmBot.txtInv3q.Text = 0
       ElseIf pocket = 4 Then
       frmBot.txtInv4.Text = 0
       frmBot.txtInv4q.Text = 0
       ElseIf pocket = 5 Then
       frmBot.txtInv5.Text = 0
       frmBot.txtInv5q.Text = 0
       ElseIf pocket = 6 Then
       frmBot.txtInv6.Text = 0
       frmBot.txtInv6q.Text = 0
       End If
       
       End If
       
       If frmBot.txtItemType.Text = 1 Then
          frmBot.txtWeapon.Text = frmBot.txtItemID.Text
       ElseIf frmBot.txtItemType.Text = 2 Then
          frmBot.txtEquip1.Text = frmBot.txtItemID.Text
        ElseIf frmBot.txtItemType.Text = 3 Then
          frmBot.txtEquip2.Text = frmBot.txtItemID.Text
        ElseIf frmBot.txtItemType.Text = 4 Then
          frmBot.txtEquip3.Text = frmBot.txtItemID.Text
        ElseIf frmBot.txtItemType.Text = 5 Then
          frmBot.txtEquip4.Text = frmBot.txtItemID.Text
        End If
       
       
       frmBot.mbr.Recordset.Update
       frmBot.sckFurc.SendData "wh " & furre & " Your " & frmBot.txtItemName.Text & " are now equiped." & vbLf
       
          End If

       
       End If
             
   End If





Else
frmBot.sckFurc.SendData "wh " & furre & " Im sorry. You are not a member. You do not have any equipment. Please Join Mesto City!" & vbLf
End If

End Sub

Sub unequipitem(furre, slot)

Dim itemnum As Integer
Dim myhp As Integer
Dim mymana As Integer
Dim mydmgnum As Integer
Dim losshp As Integer
Dim lossmana As Integer
Dim lossdmg As Integer

frmBot.mbr.Recordset.MoveFirst
Do Until frmBot.txtName.Text = furre Or frmBot.mbr.Recordset.EOF
frmBot.mbr.Recordset.MoveNext
Loop

If slot = 1 Then
itemnum = frmBot.txtEquip1.Text
ElseIf slot = 2 Then
itemnum = frmBot.txtEquip2.Text
ElseIf slot = 3 Then
itemnum = frmBot.txtEquip3.Text
ElseIf slot = 4 Then
itemnum = frmBot.txtEquip4.Text
ElseIf slot = 0 Then
itemnum = frmBot.txtWeapon.Text
End If

If frmBot.txtName.Text = furre Then

If itemnum = 0 Then
frmBot.sckFurc.SendData "wh " & furre & " You dont have an item in that slot." & vbLf

Else

If frmBot.txtInv1.Text <> 0 And frmBot.txtInv2.Text <> 0 And frmBot.txtInv3.Text <> 0 And frmBot.txtInv4.Text <> 0 And frmBot.txtInv5.Text <> 0 And frmBot.txtInv6.Text <> 0 And frmBot.txtInv1.Text <> itemnum And frmBot.txtInv2.Text <> itemnum And frmBot.txtInv3.Text <> itemnum And frmBot.txtInv4.Text <> itemnum And frmBot.txtInv5.Text <> itemnum And frmBot.txtInv6.Text <> itemnum Then
frmBot.sckFurc.SendData "wh " & furre & " Your Bag is full. Please sell something befor unequiping anything." & vbLf
Else

frmBot.items.Recordset.MoveFirst
Do Until itemnum = frmBot.txtItemID.Text Or frmBot.items.Recordset.EOF
frmBot.items.Recordset.MoveNext
Loop

myhp = frmBot.txtHp.Text
mymana = frmBot.txtMana.Text

mydmgnum = frmBot.txtDmg.Text

losshp = frmBot.txtItemHP.Text
lossmana = frmBot.txtItemMana.Text
lossdmg = frmBot.txtItemDmg.Text

frmBot.mbr.Recordset.Edit

frmBot.txtHp.Text = myhp - losshp
frmBot.txtHPLeft.Text = myhp - losshp
frmBot.txtMana.Text = mymana - lossmana
frmBot.txtManaLeft.Text = mymana - lossmana
frmBot.txtDmg.Text = mydmg - lossdmg

If slot = 1 Then
frmBot.txtEquip1.Text = 0
ElseIf slot = 2 Then
frmBot.txtEquip2.Text = 0
ElseIf slot = 3 Then
frmBot.txtEquip3.Text = 0
ElseIf slot = 4 Then
frmBot.txtEquip4.Text = 0
ElseIf slot = 0 Then
frmBot.txtWeapon.Text = 0
End If

frmBot.mbr.Recordset.Update

additem itemnum, furre

frmBot.sckFurc.SendData "wh " & furre & " Your " & frmBot.txtItemName.Text & " has been unequiped and placed in your bag." & vbLf

End If
End If


Else
frmBot.sckFurc.SendData "wh " & furre & " Im sorry. You are not a member. You do not have any equipment. Please Join Mesto City!" & vbLf
End If

End Sub

Sub dostats(furre)

frmBot.mbr.Recordset.MoveFirst
Do Until frmBot.txtName.Text = furre Or frmBot.mbr.Recordset.EOF
frmBot.mbr.Recordset.MoveNext
Loop

If frmBot.txtName.Text = furre Then

frmBot.sckFurc.SendData "wh " & furre & " Stats for " & furre & ". [Cidason ID #: " & frmBot.txtID.Text & "] [Cidason sence: " & frmBot.txtDate.Text & "] [Level: " & frmBot.txtlvl.Text & "] [HP: " & frmBot.txtHp.Text & "] [Mana: " & frmBot.txtMana.Text & "] [Damage: " & frmBot.txtDmg.Text & "]" & vbLf

Else
frmBot.sckFurc.SendData "wh " & furre & " Im sorry. You are not a member. You do not have any stats. Please Join Mesto City!" & vbLf
End If
End Sub

Sub doequip(furre)

frmBot.mbr.Recordset.MoveFirst
Do Until frmBot.txtName.Text = furre Or frmBot.mbr.Recordset.EOF
frmBot.mbr.Recordset.MoveNext
Loop

If frmBot.txtName.Text = furre Then

frmBot.sckFurc.SendData "wh " & furre & " " & furre & " equipment list. [Weapon: "

If frmBot.txtWeapon.Text <> 0 Then
frmBot.items.Recordset.MoveFirst
Do Until frmBot.txtWeapon.Text = frmBot.txtItemID.Text Or frmBot.items.Recordset.EOF
frmBot.items.Recordset.MoveNext
Loop
frmBot.sckFurc.SendData frmBot.txtItemName.Text
Else
frmBot.sckFurc.SendData "Paws"
End If
frmBot.sckFurc.SendData "] [Chest: "

If frmBot.txtEquip1.Text <> 0 Then
frmBot.items.Recordset.MoveFirst
Do Until frmBot.txtEquip1.Text = frmBot.txtItemID.Text Or frmBot.items.Recordset.EOF
frmBot.items.Recordset.MoveNext
Loop
frmBot.sckFurc.SendData frmBot.txtItemName.Text
Else
frmBot.sckFurc.SendData "Fur"
End If
frmBot.sckFurc.SendData "] [Legs: "

If frmBot.txtEquip2.Text <> 0 Then
frmBot.items.Recordset.MoveFirst
Do Until frmBot.txtEquip2.Text = frmBot.txtItemID.Text Or frmBot.items.Recordset.EOF
frmBot.items.Recordset.MoveNext
Loop
frmBot.sckFurc.SendData frmBot.txtItemName.Text
Else
frmBot.sckFurc.SendData "Fur"
End If
frmBot.sckFurc.SendData "] [Hands: "


If frmBot.txtEquip3.Text <> 0 Then
frmBot.items.Recordset.MoveFirst
Do Until frmBot.txtEquip3.Text = frmBot.txtItemID.Text Or frmBot.items.Recordset.EOF
frmBot.items.Recordset.MoveNext
Loop
frmBot.sckFurc.SendData frmBot.txtItemName.Text
Else
frmBot.sckFurc.SendData "Paws"
End If
frmBot.sckFurc.SendData "] [Feet: "


If frmBot.txtEquip4.Text <> 0 Then
frmBot.items.Recordset.MoveFirst
Do Until frmBot.txtEquip4.Text = frmBot.txtItemID.Text Or frmBot.items.Recordset.EOF
frmBot.items.Recordset.MoveNext
Loop
frmBot.sckFurc.SendData frmBot.txtItemName.Text
Else
frmBot.sckFurc.SendData "Paws"
End If


frmBot.sckFurc.SendData "]" & vbLf
Else
frmBot.sckFurc.SendData "wh " & furre & " Im sorry. You are not a member. You do not have any equipment. Please Join Mesto City!" & vbLf
End If
End Sub

Sub dobag(furre)
cont = 0


frmBot.mbr.Recordset.MoveFirst
Do Until frmBot.txtName.Text = furre Or frmBot.mbr.Recordset.EOF
frmBot.mbr.Recordset.MoveNext
Loop

If frmBot.txtName.Text = furre Then

frmBot.sckFurc.SendData "wh " & furre & " Contents of your bag: "

If frmBot.txtGold.Text <> 0 Then
frmBot.sckFurc.SendData "[Gold Pices: " & frmBot.txtGold.Text & "]"
cont = cont + 1
End If

If frmBot.txtInv1.Text <> 0 Then
frmBot.items.Recordset.MoveFirst
Do Until frmBot.txtInv1.Text = frmBot.txtItemID.Text Or frmBot.items.Recordset.EOF
frmBot.items.Recordset.MoveNext
Loop
frmBot.sckFurc.SendData " [Pocket 1: " & frmBot.txtItemName.Text & "(" & frmBot.txtInv1q.Text & ")]"
cont = cont + 1
End If

If frmBot.txtInv2.Text <> 0 Then
frmBot.items.Recordset.MoveFirst
Do Until frmBot.txtInv2.Text = frmBot.txtItemID.Text Or frmBot.items.Recordset.EOF
frmBot.items.Recordset.MoveNext
Loop
frmBot.sckFurc.SendData " [Pocket 2: " & frmBot.txtItemName.Text & "(" & frmBot.txtInv2q.Text & ")]"
cont = cont + 1
End If

If frmBot.txtInv3.Text <> 0 Then
frmBot.items.Recordset.MoveFirst
Do Until frmBot.txtInv3.Text = frmBot.txtItemID.Text Or frmBot.items.Recordset.EOF
frmBot.items.Recordset.MoveNext
Loop
frmBot.sckFurc.SendData " [Pocket 3: " & frmBot.txtItemName.Text & "(" & frmBot.txtInv3q.Text & ")]"
cont = cont + 1
End If

If frmBot.txtInv4.Text <> 0 Then
frmBot.items.Recordset.MoveFirst
Do Until frmBot.txtInv4.Text = frmBot.txtItemID.Text Or frmBot.items.Recordset.EOF
frmBot.items.Recordset.MoveNext
Loop
frmBot.sckFurc.SendData " [Pocket 4: " & frmBot.txtItemName.Text & "(" & frmBot.txtInv4q.Text & ")]"
cont = cont + 1
End If

If frmBot.txtInv5.Text <> 0 Then
frmBot.items.Recordset.MoveFirst
Do Until frmBot.txtInv5.Text = frmBot.txtItemID.Text Or frmBot.items.Recordset.EOF
frmBot.items.Recordset.MoveNext
Loop
frmBot.sckFurc.SendData " [Pocket 5: " & frmBot.txtItemName.Text & "(" & frmBot.txtInv5q.Text & ")]"
cont = cont + 1
End If

If frmBot.txtInv6.Text <> 0 Then
frmBot.items.Recordset.MoveFirst
Do Until frmBot.txtInv6.Text = frmBot.txtItemID.Text Or frmBot.items.Recordset.EOF
frmBot.items.Recordset.MoveNext
Loop
frmBot.sckFurc.SendData " [Pocket 6: " & frmBot.txtItemName.Text & "(" & frmBot.txtInv6q.Text & ")]"
cont = cont + 1
End If

If cont = 0 Then
frmBot.sckFurc.SendData " Empty"
End If

frmBot.sckFurc.SendData vbLf


Else
frmBot.sckFurc.SendData "wh " & furre & " Im sorry. You are not a member. You do not have a bag. Please Join Mesto City to get a bag!" & vbLf
End If
End Sub

