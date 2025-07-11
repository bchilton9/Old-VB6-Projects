Attribute VB_Name = "stsreg"


Sub doregister(Furre, clas)
Open "C:\Jovati\members.txt" For Input As #1
Input #1, fName, mnum
Do Until (fName = Furre) Or (EOF(1))
Input #1, fName, mnum
Loop
Close #1


If fName = Furre Then
frmBot.sckFurc.SendData "wh " & Furre & " You are allready a member." & vbLf
Else
Open "C:\Jovati\memnum.txt" For Input As #1
Input #1, nnum
Close #1
num = nnum + 1

Open "C:\Jovati\memnum.txt" For Output As #1
Write #1, num
Close #1

Open "C:\Jovati\memfiles\" & num & ".txt" For Output As #1
Write #1, Furre, num, 1, clas, 0, 0, 0, 0, 0, 0, 0, 0, 0
Close #1
Open "C:\Jovati\members.txt" For Append As #1
Write #1, Furre, num
Close #1
frmBot.sckFurc.SendData "wh " & Furre & " You are now a member." & vbLf
End If


End Sub
Sub dostats(Furre)
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

    If clas = "Wizard" Then
    stren = 9
    dex = 12
    intel = 18
    wis = 16
    cha = 13
    hp = lvl * 15
    man = lvl * 20
    End If
    If clas = "Fighter" Then
    stren = 15
    dex = 13
    intel = 10
    wis = 9
    cha = 10
    hp = lvl * 20
    man = 0
    End If
    If clas = "Thief" Then
    stren = 11
    dex = 18
    intel = 12
    wis = 10
    cha = 12
    hp = lvl * 15
    man = lvl * 10
    End If
    If clas = "Paladin" Then
    stren = 14
    dex = 12
    intel = 9
    wis = 14
    cha = 17
    hp = lvl * 20
    man = lvl * 10
    End If
    If clas = "Priest" Then
    stren = 14
    dex = 13
    intel = 11
    wis = 17
    cha = 10
    hp = lvl * 10
    man = lvl * 10
    End If
If weap = 0 Then weap = "Paws"
If weap = 1 Then weap = "Dagger"
If weap = 2 Then weap = "Knife"
If weap = 3 Then weap = "Hand ax"
If weap = 4 Then weap = "Quarterstaff"
If weap = 5 Then weap = "Spear"
If weap = 6 Then weap = "Warhammer"
If weap = 7 Then weap = "Battle ax"
If weap = 8 Then weap = "Morneing Star"
If weap = 9 Then weap = "Flail"
If weap = 10 Then weap = "Mace"
If weap = 11 Then weap = "Broad Sword"
If weap = 12 Then weap = "Short Bow"
If weap = 13 Then weap = "Crossbow"
If weap = 14 Then weap = "Shord Sword"
If weap = 15 Then weap = "Long Sword"
If weap = 16 Then weap = "TwoHand Sword"

If armo = 0 Then armo = "Fur"
If armo = 1 Then armo = "Padded"
If armo = 2 Then armo = "Leather"
If armo = 3 Then armo = "Chain Mail"
If armo = 4 Then armo = "Splint Mail"
If armo = 5 Then armo = "Ring Mail"
If armo = 6 Then armo = "Scale Mail"
If armo = 7 Then armo = "Banded Mail"
If armo = 8 Then armo = "Plate Mail"

frmBot.sckFurc.SendData "wh " & Furre & " Stats: [Member# - " & mnum & "] [Class - " & clas & "] [LvL - " & lvl & "] [Strength - " & stren & "] [Dexterity - " & dex & "] [Intelligence - " & intel & "] [Wisdom - " & wis & "] [Exp - " & xp & "%] [Health - " & hp & "] [Mana - " & man & "] [Gold - " & gold & "] [Charisma - " & cha & "] [Weapon - " & weap & "] [Armor - " & armo & "]" & vbLf
Else
frmBot.sckFurc.SendData "wh " & Furre & " You are not a member. Whisper me JOIN to learn how to become a member." & vbLf
End If
End Sub
