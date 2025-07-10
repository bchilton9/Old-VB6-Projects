Attribute VB_Name = "Whisper"
Sub DoWhisper(Furre, Msg, Index As Integer)
On Error Resume Next
'When anyone whispers the bot anything it whispers back
'The Like operator is used for string comparisons

If Msg Like "help" And Index = 0 Then
whspnum = 1
ElseIf Msg Like "stats" And Index = 0 Then
whspnum = 2
ElseIf Msg Like "bar" And Index = 0 Then
whspnum = 3
ElseIf Msg Like "join" And Index = 0 Then
whspnum = 4
ElseIf Msg Like "join fighter" And Index = 0 Then
whspnum = 5
clas = "Fighter"
ElseIf Msg Like "join wizard" And Index = 0 Then
whspnum = 5
clas = "Wizard"
ElseIf Msg Like "join thief" And Index = 0 Then
whspnum = 5
clas = "Thief"
ElseIf Msg Like "join paladin" And Index = 0 Then
whspnum = 5
clas = "Paladin"
ElseIf Msg Like "join priest" And Index = 0 Then
whspnum = 5
clas = "Preiest"
ElseIf Msg Like "fighter" And Index = 0 Then
whspnum = 6
ElseIf Msg Like "wizard" And Index = 0 Then
whspnum = 7
ElseIf Msg Like "thief" And Index = 0 Then
whspnum = 8
ElseIf Msg Like "paladin" And Index = 0 Then
whspnum = 9
ElseIf Msg Like "priest" And Index = 0 Then
whspnum = 10
'ElseIf Msg Like " And Index = 0 Then
'whspnum = 11
ElseIf Msg Like "sparing" And Index = 0 Then
whspnum = 12
ElseIf Msg Like "help" And Index = 2 Then
whspnum = 13
ElseIf Msg Like "menu" And Index = 2 Then
whspnum = 14
ElseIf Msg Like "buy weapon 1" And Index = 0 Then
wea = 1
whspnum = 15
ElseIf Msg Like "buy weapon 2" And Index = 0 Then
wea = 2
whspnum = 15
ElseIf Msg Like "buy weapon 3" And Index = 0 Then
wea = 3
whspnum = 15
ElseIf Msg Like "buy weapon 4" And Index = 0 Then
wea = 4
whspnum = 15
ElseIf Msg Like "buy weapon 5" And Index = 0 Then
wea = 5
whspnum = 15
ElseIf Msg Like "buy weapon 6" And Index = 0 Then
wea = 6
whspnum = 15
ElseIf Msg Like "buy weapon 7" And Index = 0 Then
wea = 7
whspnum = 15
ElseIf Msg Like "buy weapon 8" And Index = 0 Then
wea = 8
whspnum = 15
ElseIf Msg Like "buy weapon 9" And Index = 0 Then
wea = 9
whspnum = 15
ElseIf Msg Like "buy weapon 10" And Index = 0 Then
wea = 10
whspnum = 15
ElseIf Msg Like "buy weapon 11" And Index = 0 Then
wea = 11
whspnum = 15
ElseIf Msg Like "buy weapon 12" And Index = 0 Then
wea = 12
whspnum = 15
ElseIf Msg Like "buy weapon 13" And Index = 0 Then
wea = 13
whspnum = 15
ElseIf Msg Like "buy weapon 14" And Index = 0 Then
wea = 14
whspnum = 15
ElseIf Msg Like "buy weapon 15" And Index = 0 Then
wea = 15
whspnum = 15
ElseIf Msg Like "buy weapon 16" And Index = 0 Then
wea = 16
whspnum = 15
ElseIf Msg Like "buy armor 1" And Index = 0 Then
wea = 1
whspnum = 16
ElseIf Msg Like "buy armor 2" And Index = 0 Then
wea = 2
whspnum = 16
ElseIf Msg Like "buy armor 3" And Index = 0 Then
wea = 3
whspnum = 16
ElseIf Msg Like "buy armor 4" And Index = 0 Then
wea = 4
whspnum = 16
ElseIf Msg Like "buy armor 5" And Index = 0 Then
wea = 5
whspnum = 16
ElseIf Msg Like "buy armor 6" And Index = 0 Then
wea = 6
whspnum = 16
ElseIf Msg Like "buy armor 7" And Index = 0 Then
wea = 7
whspnum = 16
ElseIf Msg Like "buy armor 8" And Index = 0 Then
wea = 8
whspnum = 16
ElseIf Msg Like "buy" And Index = 0 Then
whspnum = 17
ElseIf Msg Like "weapon" And Index = 0 Then
whspnum = 18
ElseIf Msg Like "armor" And Index = 0 Then
whspnum = 19
ElseIf Msg Like "close" Then
    If Furre = "keny" Then
        whspnum = 20
    Else
        whspnum = 0
    End If
ElseIf Msg Like "reset" Then whspnum = 21
ElseIf Msg Like "open" Then
    If Furre = "keny" Then
        whspnum = 22
    Else
        whspnum = 0
    End If
Else
whspnum = 0
End If


If whspnum = 0 Then frmBot.sckFurc(Index).SendData "wh " & Furre & "  I dont understand. Try /Pena HELP insted." & vbLf
If whspnum = 1 Then frmBot.sckFurc(Index).SendData "wh " & Furre & " To learn how I work whisper me a command. My commands are JOIN, SPARING, STATS, BUY, BAR." & vbLf
If whspnum = 2 Then dostats Furre, Index
If whspnum = 3 Then frmBot.sckFurc(Index).SendData "wh " & Furre & " Dwain runs the pub here in Tidal. Whisper him HELP for more info." & vbLf
If whspnum = 4 Then frmBot.sckFurc(Index).SendData "wh " & Furre & " To join you must first chose a class. Classes are Fighter, Wizard, Thief, Paladin, and Priest. Whisper me a class name to lern more. After you have chosen a class whisper me JOIN CLASS" & vbLf
If whspnum = 5 Then doregister Furre, clas, Index
If whspnum = 6 Then frmBot.sckFurc(Index).SendData "wh " & Furre & " This mighty Fighter is a brave adventurer and good friend. He defends his companions against monsters and outher enemies with fierce devotion." & vbLf
If whspnum = 7 Then frmBot.sckFurc(Index).SendData "wh " & Furre & " This powerful Wizard controls vast magical energies, shaping them and casting them as mighty spells. He studies strange tongues and obscure facts and devotes much of his time to magical research." & vbLf
If whspnum = 8 Then frmBot.sckFurc(Index).SendData "wh " & Furre & " This cunning Thief makes his way through the world using his wits, stealth, and roguish talents. His companions depend on his skills to aid them in avoiding locks, traps, and outher hidden dangers." & vbLf
If whspnum = 9 Then frmBot.sckFurc(Index).SendData "wh " & Furre & " The Paladin. This holy warrior stands pure and true against the evils of the world. He upholds all that is good, living for the ideals of righteousness, justice, honesty, and chivalry. He strives to be a liveing example of these virtues so that outhers might learn from him as wall as gain by his actions." & vbLf
If whspnum = 10 Then frmBot.sckFurc(Index).SendData "wh " & Furre & " The Priest serves as a protector and healer for his companions. When evil threatens, he woun't hesitate to hunt it down and destroy it. He calls upon the power of his faith to cast powerful spells to aid his allies and distory his enemies." & vbLf
'If whspnum = 11 Then frmBoT.sckFurc(Index).SendData "wh " & Furre & " Empty." & vbLf
If whspnum = 12 Then frmBot.sckFurc(Index).SendData "wh " & Furre & " Furres must move into a sparing arena for the fight to begine. It will be automated from then on." & vbLf
If whspnum = 13 Then frmBot.sckFurc(Index).SendData "wh " & Furre & " To order something type #ITEM# replaceing ITEM with anything you want Whisper me MENU for a list of what I can make." & vbLf
If whspnum = 14 Then frmBot.sckFurc(Index).SendData "wh " & Furre & " Food: HOTDOG HAMBURGER. Drinks: BEER ROOTBEER" & vbLf
If whspnum = 15 Then buyweapon Furre, wea, Index
If whspnum = 16 Then buyarmo Furre, wea, Index
If whspnum = 17 Then frmBot.sckFurc(Index).SendData "wh " & Furre & " To buy weapons and armor whisper me BUY WEAPON # or BUY ARMOR # replaceing # with weapon or armor number. For lists whisper WEAPON or ARMOR to me. Items are 10X the item number in Gold." & vbLf
If whspnum = 18 Then frmBot.sckFurc(Index).SendData "wh " & Furre & " Weapons: [Dagger - 1] [Knife - 2] [Hand ax - 3] [Quarterstaff - 4] [Spear - 5] [Warhammer - 6] [Battle ax - 7] [Morneing Star - 8] [Flail - 9] [Mace - 10] [Broad Sword - 11] [Short Bow - 12] [Crossbow - 13] [Shord Sword - 14] [Long Sword - 15] [TwoHand Sword - 16]." & vbLf
If whspnum = 19 Then frmBot.sckFurc(Index).SendData "wh " & Furre & " Armor: [Padded - 1] [Leather - 2] [Chain Mail - 3] [Splint Mail - 4] [Ring Mail - 5] [Scale Mail - 6] [Banded Mail - 7] [Plate Mail - 8]." & vbLf
If whspnum = 20 Then
    frmBot.reset Index
    frmBot.sckFurc(Index).SendData "wh " & Furre & " Arena Closed" & vbLf & "m 7" & vbLf & Chr(34) & "emit This arena is closed!" & vbLf
frmBot.sckFurc(Index).SendData Chr(34) & "emit Arena " & Index & " is closed!" & vbLf
End If
If whspnum = 21 Then
    frmBot.reset Index
    frmBot.sckFurc(Index).SendData Chr(34) & "emitloud Arena " & Index & " was reset by " & Furre & "!" & vbLf
End If
If whspnum = 22 Then
    If Index = 1 Then
        frmBot.sckFurc(Index).SendData "m 3" & vbLf
    End If
    frmBot.sckFurc(Index).SendData Chr(34) & "emit This arena is closed!" & vbLf
    frmBot.sckFurc(Index).SendData Chr(34) & "emit Arena " & Index & " is now open!" & vbLf
End If
End Sub

Sub dostats(Furre, Index As Integer)
On Error Resume Next
Open "members.txt" For Input As #1
Input #1, fName, mnum
Do Until (fName = Furre) Or (EOF(1))
Input #1, fName, mnum
Loop
Close #1
If fName = Furre Then
Open "memfiles\" & mnum & ".txt" For Input As #1
Input #1, fName, mnum, lvl, clas, gold, xp, weap, armo, sone, stwo, sthe, sfor, sfiv
Close #1

    If clas = "Wizard" Then
    hp = lvl * 15
    man = lvl * 20
    End If
    If clas = "Fighter" Then
    hp = lvl * 20
    man = 0
    End If
    If clas = "Thief" Then
    hp = lvl * 15
    man = lvl * 10
    End If
    If clas = "Paladin" Then
    hp = lvl * 20
    man = lvl * 10
    End If
    If clas = "Priest" Then
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

frmBot.sckFurc(Index).SendData "wh " & Furre & " Stats: [Member# - " & mnum & "] [Class - " & clas & "] [LvL - " & lvl & "] [Exp - " & xp & "%] [Health - " & hp & "] [Mana - " & man & "] [Gold - " & gold & "] [Weapon - " & weap & "] [Armor - " & armo & "]" & vbLf
Else
frmBot.sckFurc(Index).SendData "wh " & Furre & " You are not a member. Whisper me JOIN to learn how to become a member." & vbLf
End If
End Sub

Sub doregister(Furre, clas, Index As Integer)
On Error Resume Next
Open "members.txt" For Input As #1
Input #1, fName, mnum
Do Until (fName = Furre) Or (EOF(1))
Input #1, fName, mnum
Loop
Close #1

Open "memnum.txt" For Input As #1
Input #1, nnum
Close #1
num = nnum + 1

Open "memnum.txt" For Output As #1
Write #1, num
Close #1

If fName = Furre Then
frmBot.sckFurc(Index).SendData "wh " & Furre & " You are allready a member." & vbLf
Else
Open "memfiles\" & num & ".txt" For Output As #1
Write #1, Furre, num, 1, clas, 0, 0, 0, 0, 0, 0, 0, 0, 0
Close #1
Open "members.txt" For Append As #1
Write #1, Furre, num
Close #1
frmBot.sckFurc(Index).SendData "wh " & Furre & " You are now a member." & vbLf
End If


End Sub

Sub buyweapon(Furre, wea, Index As Integer)
On Error Resume Next
Open "members.txt" For Input As #1
Input #1, fName, mnum
Do Until (fName = Furre) Or (EOF(1))
Input #1, fName, mnum
Loop
Close #1
If fName = Furre Then
    Open "memfiles\" & mnum & ".txt" For Input As #1
    Input #1, fName, mnum, lvl, clas, gold, xp, weap, armo, sone, stwo, sthe, sfor, sfiv
    Close #1
    price = wea * 10
    If wea < weap Then
        frmBot.sckFurc(Index).SendData "wh " & Furre & " you would be dumb to buy a less of a weapon." & vbLf
    Else
    If price < gold Then
        ngold = gold - price
        Open "memfiles\" & mnum & ".txt" For Output As #1
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
        frmBot.sckFurc(Index).SendData "wh " & Furre & " you now have a " & wea & vbLf
    Else
        frmBot.sckFurc(Index).SendData "wh " & Furre & " you dont have enuff gold to buy that." & vbLf
    End If
    End If
Else
    frmBot.sckFurc(Index).SendData "wh " & Furre & " You are not a member." & vbLf
End If
End Sub

Sub buyarmo(Furre, wea, Index)
On Error Resume Next
Open "members.txt" For Input As #1
Input #1, fName, mnum
Do Until (fName = Furre) Or (EOF(1))
Input #1, fName, mnum
Loop
Close #1
If fName = Furre Then
    Open "memfiles\" & mnum & ".txt" For Input As #1
    Input #1, fName, mnum, lvl, clas, gold, xp, weap, armo, sone, stwo, sthe, sfor, sfiv
    Close #1
    price = wea * 10
    If wea < weap Then
        frmBot.sckFurc(Index).SendData "wh " & Furre & " you would be dumb to buy a less of armor." & vbLf
    Else
    If price < gold Then
        ngold = gold - price
        Open "memfiles\" & mnum & ".txt" For Output As #1
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
        frmBot.sckFurc(Index).SendData "wh " & Furre & " you now have " & wea & vbLf
    Else
        frmBot.sckFurc(Index).SendData "wh " & Furre & " you dont have enuff gold to buy that." & vbLf
    End If
    End If
Else
    frmBot.sckFurc(Index).SendData "wh " & Furre & " You are not a member." & vbLf
End If
End Sub
