Attribute VB_Name = "buy"

Sub buyweapon(Furre, wea)
Open "C:\Jovati\members.txt" For Input As #1
Input #1, fName, mnum
Do Until (fName = Furre) Or (EOF(1))
Input #1, fName, mnum
Loop
Close #1
If fName = Furre Then
    Open "C:\Jovati\memfiles\" & mnum & ".txt" For Input As #1
    Input #1, fName, mnum, lvl, clas, gold, xp, weap, armo, sone, stwo, sthe, sfor, sfiv
    Close #1
    price = wea * 10
    If wea < weap Then
        frmBot.sckFurc.SendData "wh " & Furre & " you would be dumb to buy a less of a weapon." & vbLf
    Else
    If price < gold Then
        ngold = gold - price
        Open "C:\Jovati\memfiles\" & mnum & ".txt" For Output As #1
        Write #1, fName, mnum, lvl, clas, ngold, xp, wea, armo, sone, stwo, sthe, sfor, sfiv
        Close #1
If wea = 1 Then wea = "Dagger"
If wea = 2 Then wea = "Knife"
If wea = 3 Then wea = "Hand ax"
If wea = 4 Then wea = "Quarterstaff"
If wea = 5 Then wea = "Spear"
If wea = 6 Then wea = "Warhammer"
If wea = 7 Then wea = "Battle ax"
If wea = 8 Then wea = "Morneing Star"
If wea = 9 Then wea = "Flail"
If wea = 10 Then wea = "Mace"
If wea = 11 Then wea = "Broad Sword"
If wea = 12 Then wea = "Short Bow"
If wea = 13 Then wea = "Crossbow"
If wea = 14 Then wea = "Shord Sword"
If wea = 15 Then wea = "Long Sword"
If wea = 16 Then wea = "TwoHand Sword"
        frmBot.sckFurc.SendData "wh " & Furre & " you now have a " & wea & vbLf
    Else
        frmBot.sckFurc.SendData "wh " & Furre & " you dont have enuff gold to buy that." & vbLf
    End If
    End If
Else
    frmBot.sckFurc.SendData "wh " & Furre & " You are not a member." & vbLf
End If
End Sub

Sub buyarmo(Furre, wea)
Open "C:\Jovati\members.txt" For Input As #1
Input #1, fName, mnum
Do Until (fName = Furre) Or (EOF(1))
Input #1, fName, mnum
Loop
Close #1
If fName = Furre Then
    Open "C:\Jovati\memfiles\" & mnum & ".txt" For Input As #1
    Input #1, fName, mnum, lvl, clas, gold, xp, weap, armo, sone, stwo, sthe, sfor, sfiv
    Close #1
    price = wea * 10
    If wea < weap Then
        frmBot.sckFurc.SendData "wh " & Furre & " you would be dumb to buy a less of armor." & vbLf
    Else
    If price < gold Then
        ngold = gold - price
        Open "C:\Jovati\memfiles\" & mnum & ".txt" For Output As #1
        Write #1, fName, mnum, lvl, clas, ngold, xp, weap, wea, sone, stwo, sthe, sfor, sfiv
        Close #1
If wea = 0 Then wea = "Fir"
If wea = 1 Then wea = "Padded"
If wea = 2 Then wea = "Leather"
If wea = 3 Then wea = "Chain Mail"
If wea = 4 Then wea = "Splint Mail"
If wea = 5 Then wea = "Ring Mail"
If wea = 6 Then wea = "Scale Mail"
If wea = 7 Then wea = "Banded Mail"
If wea = 8 Then wea = "Plate Mail"
        frmBot.sckFurc.SendData "wh " & Furre & " you now have " & wea & vbLf
    Else
        frmBot.sckFurc.SendData "wh " & Furre & " you dont have enuff gold to buy that." & vbLf
    End If
    End If
Else
    frmBot.sckFurc.SendData "wh " & Furre & " You are not a member." & vbLf
End If
End Sub
