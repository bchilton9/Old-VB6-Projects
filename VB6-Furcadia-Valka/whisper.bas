Attribute VB_Name = "whisper"
Sub DoWhisper(Furre, Msg)
If Msg Like "help" Then
whspnum = 1
ElseIf Msg Like "stats" Then
whspnum = 2
ElseIf Msg Like "na" Then
whspnum = 3
ElseIf Msg Like "join" Then
whspnum = 4
ElseIf Msg Like "join fighter" Then
whspnum = 5
clas = "Fighter"
ElseIf Msg Like "join wizard" Then
whspnum = 5
clas = "Wizard"
ElseIf Msg Like "join thief" Then
whspnum = 5
clas = "Thief"
ElseIf Msg Like "join paladin" Then
whspnum = 5
clas = "Paladin"
ElseIf Msg Like "join priest" Then
whspnum = 5
clas = "Preiest"
ElseIf Msg Like "fighter" Then
whspnum = 6
ElseIf Msg Like "wizard" Then
whspnum = 7
ElseIf Msg Like "thief" Then
whspnum = 8
ElseIf Msg Like "paladin" Then
whspnum = 9
ElseIf Msg Like "priest" Then
whspnum = 10
ElseIf Msg Like "naa" Then
whspnum = 11
ElseIf Msg Like "fight" Then
whspnum = 12
ElseIf Msg Like "naaa" Then
whspnum = 13
ElseIf Msg Like "naaaa" Then
whspnum = 14
ElseIf Msg Like "buy weapon 1" Then
wea = 1
whspnum = 15
ElseIf Msg Like "buy weapon 2" Then
wea = 2
whspnum = 15
ElseIf Msg Like "buy weapon 3" Then
wea = 3
whspnum = 15
ElseIf Msg Like "buy weapon 4" Then
wea = 4
whspnum = 15
ElseIf Msg Like "buy weapon 5" Then
wea = 5
whspnum = 15
ElseIf Msg Like "buy weapon 6" Then
wea = 6
whspnum = 15
ElseIf Msg Like "buy weapon 7" Then
wea = 7
whspnum = 15
ElseIf Msg Like "buy weapon 8" Then
wea = 8
whspnum = 15
ElseIf Msg Like "buy weapon 9" Then
wea = 9
whspnum = 15
ElseIf Msg Like "buy weapon 10" Then
wea = 10
whspnum = 15
ElseIf Msg Like "buy weapon 11" Then
wea = 11
whspnum = 15
ElseIf Msg Like "buy weapon 12" Then
wea = 12
whspnum = 15
ElseIf Msg Like "buy weapon 13" Then
wea = 13
whspnum = 15
ElseIf Msg Like "buy weapon 14" Then
wea = 14
whspnum = 15
ElseIf Msg Like "buy weapon 15" Then
wea = 15
whspnum = 15
ElseIf Msg Like "buy weapon 16" Then
wea = 16
whspnum = 15
ElseIf Msg Like "buy armor 1" Then
wea = 1
whspnum = 16
ElseIf Msg Like "buy armor 2" Then
wea = 2
whspnum = 16
ElseIf Msg Like "buy armor 3" Then
wea = 3
whspnum = 16
ElseIf Msg Like "buy armor 4" Then
wea = 4
whspnum = 16
ElseIf Msg Like "buy armor 5" Then
wea = 5
whspnum = 16
ElseIf Msg Like "buy armor 6" Then
wea = 6
whspnum = 16
ElseIf Msg Like "buy armor 7" Then
wea = 7
whspnum = 16
ElseIf Msg Like "buy armor 8" Then
wea = 8
whspnum = 16
ElseIf Msg Like "buy" Then
whspnum = 17
ElseIf Msg Like "weapon" Then
whspnum = 18
ElseIf Msg Like "armor" Then
whspnum = 19
Else
whspnum = 0
End If
If whspnum = 0 Then frmBot.sckFurc.SendData "wh " & Furre & " I dont understand. Try Whispering me help." & vbLf
If whspnum = 1 Then frmBot.sckFurc.SendData "wh " & Furre & " Commands, JOIN, FIGHT, STATS, BUY." & vbLf
If whspnum = 2 Then stsreg.dostats Furre
If whspnum = 3 Then frmBot.sckFurc.SendData "wh " & Furre & " I dont understand. Try Whispering me help." & vbLf
If whspnum = 4 Then frmBot.sckFurc.SendData "wh " & Furre & " To join you must first chose a class. Classes are Fighter, Wizard, Thief, Paladin, and Priest. Whisper me a class name to lern more. After you have chosen a class whisper me JOIN CLASS" & vbLf
If whspnum = 5 Then stsreg.doregister Furre, clas
If whspnum = 6 Then frmBot.sckFurc.SendData "wh " & Furre & " The mighty Fighter is a brave adventurer and good friend. He defends his companions against monsters and outher enemies with fierce devotion." & vbLf
If whspnum = 7 Then frmBot.sckFurc.SendData "wh " & Furre & " The powerful Wizard controls vast magical energies, shaping them and casting them as mighty spells. He studies strange tongues and obscure facts and devotes much of his time to magical research." & vbLf
If whspnum = 8 Then frmBot.sckFurc.SendData "wh " & Furre & " The cunning Thief makes his way through the world using his wits, stealth, and roguish talents. His companions depend on his skills to aid them in avoiding locks, traps, and outher hidden dangers." & vbLf
If whspnum = 9 Then frmBot.sckFurc.SendData "wh " & Furre & " The Paladin. This holy warrior stands pure and true against the evils of the world. He upholds all that is good, living for the ideals of righteousness, justice, honesty, and chivalry. He strives to be a liveing example of these virtues so that outhers might learn from him as wall as gain by his actions." & vbLf
If whspnum = 10 Then frmBot.sckFurc.SendData "wh " & Furre & " The Priest serves as a protector and healer for his companions. When evil threatens, he woun't hesitate to hunt it down and destroy it. He calls upon the power of his faith to cast powerful spells to aid his allies and distory his enemies." & vbLf
If whspnum = 11 Then frmBot.sckFurc.SendData "wh " & Furre & " I dont understand. Try Whispering me help." & vbLf
If whspnum = 12 Then frmBot.sckFurc.SendData "wh " & Furre & " Fighting as all anomated. Move into an open sparing ring and let the fight begine." & vbLf
If whspnum = 13 Then frmBot.sckFurc.SendData "wh " & Furre & " I dont understand. Try Whispering me help." & vbLf
If whspnum = 14 Then frmBot.sckFurc.SendData "wh " & Furre & " I dont understand. Try Whispering me help." & vbLf
If whspnum = 15 Then buy.buyweapon Furre, wea
If whspnum = 16 Then buy.buyarmo Furre, wea
If whspnum = 17 Then frmBot.sckFurc.SendData "wh " & Furre & " To buy weapons and armor whisper me BUY WEAPON # or BUY ARMOR # replaceing # with weapon or armor number. For lists whisper WEAPON or ARMOR to me. Items are 10X the item number in Gold." & vbLf
If whspnum = 18 Then frmBot.sckFurc.SendData "wh " & Furre & " Weapons: [Dagger - 1] [Knife - 2] [Hand ax - 3] [Quarterstaff - 4] [Spear - 5] [Warhammer - 6] [Battle ax - 7] [Morneing Star - 8] [Flail - 9] [Mace - 10] [Broad Sword - 11] [Short Bow - 12] [Crossbow - 13] [Shord Sword - 14] [Long Sword - 15] [TwoHand Sword - 16]." & vbLf
If whspnum = 19 Then frmBot.sckFurc.SendData "wh " & Furre & " Armor: [Padded - 1] [Leather - 2] [Chain Mail - 3] [Splint Mail - 4] [Ring Mail - 5] [Scale Mail - 6] [Banded Mail - 7] [Plate Mail - 8]." & vbLf
End Sub
