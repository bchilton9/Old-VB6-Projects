Attribute VB_Name = "whispers"
Sub DoWhisper(furre, Msg)
On Error Resume Next

'If Msg Like "help" Then
'frmBot.sckFurc.SendData "wh " & furre & " My Whispers work with commands. In my wispers commands you can use will be inclosed in brackets like [COMMAND]. To use a command whisper me the command with out the brackets (exp: /mesto COMMAND). Help Commands: [CIDASON COMMANDS] [JOIN]." & vbLf

If Msg Like "bag" Then
dobag furre

ElseIf Msg Like "stats" Then
dostats furre

ElseIf Msg Like "equipment" Then
doequip furre

'ElseIf Msg Like "join" Then
'frmBot.sckFurc.SendData "wh " & furre & " If you would like to join Mesto City please see Jeji in the Cidasonship Office located just a few steps north west of where you enter the city." & vbLf

'ElseIf Msg Like "cidason commands" Then
'frmBot.sckFurc.SendData "wh " & furre & " Cidason Commands: [BAG] [EQUIPMENT] [STATS] [EQUIP]." & vbLf

'ElseIf Msg Like "equip" Then
'frmBot.sckFurc.SendData "wh " & furre & " To equip items out of your bag whisper me the command [EQUIP #] replaceing the # with the number of the bag pocket that has the item you would like to equip. To unequip the item whisper me the command [UNEQUIP SLOT] replaceing the SLOT with the slot you want to unequip. Slots are [WEAPON] [CHEST] [LEGS] [HANDS] [FEET]" & vbLf

ElseIf Msg Like "equip 1" Then
equipitem furre, 1

ElseIf Msg Like "equip 2" Then
equipitem furre, 2

ElseIf Msg Like "equip 3" Then
equipitem furre, 3

ElseIf Msg Like "equip 4" Then
equipitem furre, 4

ElseIf Msg Like "equip 5" Then
equipitem furre, 5

ElseIf Msg Like "equip 6" Then
equipitem furre, 6

ElseIf Msg Like "unequip weapon" Then
unequipitem furre, 0

ElseIf Msg Like "unequip chest" Then
unequipitem furre, 1

ElseIf Msg Like "unequip legs" Then
unequipitem furre, 2

ElseIf Msg Like "unequip hands" Then
unequipitem furre, 3

ElseIf Msg Like "unequip feet" Then
unequipitem furre, 4


Else
frmBot.sckFurc.SendData "wh " & furre & " I dont understand." & vbLf

End If

End Sub
