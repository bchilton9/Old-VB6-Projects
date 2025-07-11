Attribute VB_Name = "whispers"
Sub DoHelp(Furre, Msg, Txt)

If Msg Like "*help*" Then
    frmBot.sckFurc.SendData "wh " & Furre & " http://www.erenetwork.com/mailsys||||||||||||||||||||||||||||||||||||||||||||||||Thank you for choosing Mailsys as your Furcadian mail service. Whisper me one of the following to learn how I run:|||||||||||#mail, #entertainment, #news, #stats.||||||||||||||||||||||||||||||||||||||||||||||||||||If you ever need help with mailsys just put a # infront of the command and ill tell you how it works.|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||Press F8 to go to the MailSys website. At the website you can learn eavrything about MailSys.||||||||||||||||||||||||||||||||||||" & vbLf

ElseIf Msg Like "*mail*" Then
frmBot.sckFurc.SendData "wh " & Furre & " ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||The following commands will help you Learn my mail service:||||||||||||||||||||||||#join, #read, #delete, #send, #check, #suggest, #card.|||||||||||||||||||||||||||||" & vbLf

ElseIf Msg Like "*join*" Then
frmBot.sckFurc.SendData "wh " & Furre & " |||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||If you wish to use my messaging service you must first sign up.|||||||||||||||||||||It is easy.|||||||||||||||||||||||||||||||Simply type '/Mailsys JOIN' and I will create your account.||||||||||||||||||||||||" & vbLf

ElseIf Msg Like "*read*" Then
frmBot.sckFurc.SendData "wh " & Furre & " ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||To check your messages type '/Mailsys Read' and I will give you a list of all who have left you a message followed by a message number.|||||||||||||||||||||||||||Type '/Mailsys Read #' to read the message replacing the # with the message number.||" & vbLf

ElseIf Msg Like "*delete*" Then
frmBot.sckFurc.SendData "wh " & Furre & " ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||To delete a message on your list type '/Mailsys Delete #' Replacing the # with the message number.|||||||||||||||||||||||||||" & vbLf

ElseIf Msg Like "*send*" Then
frmBot.sckFurc.SendData "wh " & Furre & " ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||To send a message type '/Mailsys SEND Furrename MESSAGE Messagebody' replacing Furrename with the name of the person you wish to mail, and messagebody with what you want the message to say.|||" & vbLf

ElseIf Msg Like "*check*" Then
frmBot.sckFurc.SendData "wh " & Furre & " ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||To see if a certain furry is registered with my serivice, type '/Mailsys CHECK furrename' replaceing furrename with the name of the person you wish to check on.|" & vbLf

ElseIf Msg Like "*suggest*" Then
frmBot.sckFurc.SendData "wh " & Furre & " ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||Type '/Mailsys SUGGEST Furrename' Replacing Furrename with the Furry's name and I will check every 30 minutes for them to be on and send them a message.|||" & vbLf

ElseIf Msg Like "*card" Then
frmBot.sckFurc.SendData "wh " & Furre & " http://www.erenetwork.com/mailsys/images/index.html||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||To send a greeting card to another furry type '/Mailsys CARD Furrename IMAGE # MESSAGE MM' replacing Furrename with the furry's name, # with the number of the image, and MM with the message.||||||||||||||||||||||||||||||||||||||||||||||Example '/Mailsys CARD Felorin IMAGE 3 Furcadia is great!'.||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||To view the available images please press F8.|||||||||||||||||||||||||||||||||||||||||||" & vbLf

ElseIf Msg Like "*insult*" Then
frmBot.sckFurc.SendData "wh " & Furre & " ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||I can randomly generate Insults. If you would like to see one just type '/Mailsys insult'.|||||||||||||||||||||||||||||||||||Coming Soon - Send your random Insult.|||" & vbLf

ElseIf Msg Like "*namegen*" Then
frmBot.sckFurc.SendData "wh " & Furre & " ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||Mystic Name Giver has now been brought to furcadia.|||||||||||||||||||||||||||||||||Just whisper me the following styles and I will generate you an authentic name: Albino, Alver, Deverry, Elf, Felana, Galler, Orc.|||||||||||||||||||||||||||||||||||||||||" & vbLf

ElseIf Msg Like "*entertainment*" Then
frmBot.sckFurc.SendData "wh " & Furre & " ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||Entertainment! Just stuff for fun:|||||||#insult, #namegen.||||||||||||||||||||||||||||" & vbLf

ElseIf Msg Like "*news*" Then
DoWhisper Furre, Msg, Txt

ElseIf Msg Like "*stats*" Then
DoWhisper Furre, Msg, Txt

Else
    frmBot.sckFurc.SendData "wh " & Furre & " I dont understand. Whisper me #HELP to learn how to use my service." & vbLf
End If
End Sub
Sub DoWhisper(Furre, Msg, Txt)

If Msg Like "*help*" Then
DoHelp Furre, Msg, Txt

ElseIf Msg Like "albino" Then
albino
frmBot.sckFurc.SendData "wh " & Furre & " Your Mystic Name is " & Chr(34) & frmBot.usrname.Text & Chr(34) & vbLf

ElseIf Msg Like "alver" Then
Alver
frmBot.sckFurc.SendData "wh " & Furre & " Your Mystic Name is " & Chr(34) & frmBot.usrname.Text & Chr(34) & vbLf

ElseIf Msg Like "deverry" Then
Deverry
frmBot.sckFurc.SendData "wh " & Furre & " Your Mystic Name is " & Chr(34) & frmBot.usrname.Text & Chr(34) & vbLf

ElseIf Msg Like "elf" Then
elf
frmBot.sckFurc.SendData "wh " & Furre & " Your Mystic Name is " & Chr(34) & frmBot.usrname.Text & Chr(34) & vbLf

ElseIf Msg Like "felanna" Then
felana
frmBot.sckFurc.SendData "wh " & Furre & " Your Mystic Name is " & Chr(34) & frmBot.usrname.Text & Chr(34) & vbLf

ElseIf Msg Like "galler" Then
galler
frmBot.sckFurc.SendData "wh " & Furre & " Your Mystic Name is " & Chr(34) & frmBot.usrname.Text & Chr(34) & vbLf

ElseIf Msg Like "orc" Then
orc
frmBot.sckFurc.SendData "wh " & Furre & " Your Mystic Name is " & Chr(34) & frmBot.usrname.Text & Chr(34) & vbLf

ElseIf Msg Like "join" Then
    join.dojoin Furre, Txt
ElseIf Msg Like "read" Then
    read.doread Furre, Txt
ElseIf Msg Like "proxy" Then
    read.doproxy Furre, Txt
ElseIf Msg Like "remove" Then
    sugchk.remove Furre
ElseIf Msg Like "unforward" Then
    Forword.dounforward Furre, Txt
    
ElseIf Msg Like "*stats*" Then
    Open "C:\mailsys\memnum.txt" For Input As #1
    Input #1, memn
    Close #1
    Open "C:\mailsys\sent.txt" For Input As #3
    Input #3, sent
    Close #3
    yes = 0
    If frmBot.Day >= 10 Then yes = yes + 1
    If frmBot.Day >= 100 Then yes = yes + 1
    If frmBot.Hour >= 10 Then yes = yes + 1
    If frmBot.Minute >= 10 Then yes = yes + 1
    If yes = 0 Then modLine = "||||||||||"
    If yes = 1 Then modLine = "|||||||||"
    If yes = 2 Then modLine = "||||||||"
    If yes = 3 Then modLine = "|||||||"
    If yes = 4 Then modLine = "||||||"
    frmBot.sckFurc.SendData "wh " & Furre & " ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||The softwhere I use to run my services ||||was made with VB6.|||||||||||||||||||||||It is curently version " & frmBot.vers & ".|||||||||||||||||||I have been on the clock for |||||||||||||||||" & frmBot.Day & " Day(s) " & frmBot.Hour & " Hour(s) " & frmBot.Minute & " Minute(s)." & modLine & "I have " & memn & " Furre's useing my services.|||||I have delivered " & sent & " message's.||||||||||||||I am an official bot of the AICC. ||||||||||||||||[http://AICC.erenetwork.com]||||||||And SilverSide Interactive.|||||||||||||||||||||||||[http://www.dulledge.net]||||||||||||||" & vbLf

ElseIf Msg Like "*news*" Then
    Open "C:\mailsys\news.txt" For Input As #1
    Input #1, news
    Close #1
    frmBot.sckFurc.SendData "wh " & Furre & " " & news & vbLf

ElseIf Msg Like "*insult*" Then
geninsult
frmBot.sckFurc.SendData "wh " & Furre & " " & frmBot.Insult.Text & vbLf

Else
    frmBot.sckFurc.SendData "wh " & Furre & " I dont understand. Whisper me #help to learn how to use my service." & vbLf
End If
End Sub
