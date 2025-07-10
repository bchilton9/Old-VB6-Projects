Attribute VB_Name = "whispers"
Sub DoWhisper(Furre, Msg)

If Msg Like "help" Then
frmBot.sckFurc.SendData "wh " & Furre & " Whisper one of the following commands without the brackets to get the corresponding information: [clients] [url] [vote] [tutorials] [members] [join] [requests] [taneests]" & vbLf

ElseIf Msg Like "clients" Then
frmBot.sckFurc.SendData "wh " & Furre & " We currently help with the following bot creation clients: FurBot, zMUD, MUSHclient, and Visual Basic" & vbLf

ElseIf Msg Like "url" Then
frmBot.sckFurc.SendData "wh " & Furre & " Visit the Archives of the AICC at http://www.erenetwork.com/aicc/.  [F8 Now]" & vbLf

ElseIf Msg Like "vote" Then
frmBot.sckFurc.SendData "wh " & Furre & " If you like our site, vote for us on the Furcadia Users Database. http://fud.axaqy.ro/vote.php?site_id=13" & vbLf

ElseIf Msg Like "tutorials" Then
frmBot.sckFurc.SendData "wh " & Furre & " We currently have seven FurBot tutorials and four MUSHclient tutorials. To access them, you can bump the bookshelves on the southeast side of the library or go to the tutorials section of our website. [http://www.erenetwork.com/aicc] [F8 Now]" & vbLf

ElseIf Msg Like "factions" Then
frmBot.sckFurc.SendData "wh " & Furre & " The AICC has four factions: Archivists, Datamancers, Debuggers, and Overseers. Whisper the name of a faction you wish to know more about" & vbLf

ElseIf Msg Like "join" Then
frmBot.sckFurc.SendData "wh " & Furre & " Membership is by invite from another member only. The referring member will have to review your botmaking skills, then you may be interviewed by Red|Dragon. Memberships may be revoked at anytime with or without reason" & vbLf

ElseIf Msg Like "requests" Then
frmBot.sckFurc.SendData "wh " & Furre & " We do not currently have any formal system to take requests for bots. If you are in need of a bot, ask one of our members politely, one time. If he/she decides to help you, do NOT pester them to finish it. Bots cannot and should not be given a deadline for finishing. If he/she decides they are unable to help you, do NOT pester them to help you" & vbLf

ElseIf Msg Like "Archivists" Then
frmBot.sckFurc.SendData "wh " & Furre & " The Archivists write tutorials for aspiring botmakers to use as a guide" & vbLf

ElseIf Msg Like "Datamancers" Then
frmBot.sckFurc.SendData "wh " & Furre & " The Datamancers create example bots for aspiring botmakers to use as a reference. They also make bots for specific purposes" & vbLf

ElseIf Msg Like "Debuggers" Then
frmBot.sckFurc.SendData "wh " & Furre & " The Debuggers help furres find problems with their bots, and then find solutions to the problems" & vbLf

ElseIf Msg Like "Overseers" Then
frmBot.sckFurc.SendData "wh " & Furre & " The Overseers greet new furres to the dream, offer tours, and explain our rules and policies" & vbLf

ElseIf Msg Like "actions" Then
frmBot.sckFurc.SendData "wh " & Furre & " The following are commands you can use with zMUD, MUSHclient, and Visual Basic to control your bot: get, who, use, sit, lie, liedown, stand, " & Chr(34) & "YourMessageHere, :YourEmoteHere, -YourShoutHere" & vbLf

ElseIf Msg Like "movement" Then
frmBot.sckFurc.SendData "wh " & Furre & " The following are commands used to move your bot with zMUD, MUSHclient, and Visual Basic: m 1, m 3, m 7, m 9, <, >" & vbLf

ElseIf Msg Like "taneests" Then
frmBot.sckFurc.SendData "wh " & Furre & " The following furres are Taneest(a)s of the AICC: Mys', Red Dragon, C.H McCormick." & vbLf

ElseIf Msg Like "transfer * to *" Then
msgArray = Split(Msg, " "): modGuild.Transfer_Points Furre, msgArray(1), msgArray(3)
    
ElseIf Msg = "balance" Then modGuild.Balance Furre
    
ElseIf Msg = "members" Then modGuild.List_Members Furre

ElseIf Msg = "yes" Then modGuild.Help_Request Msg, HelpReq, Furre
    
ElseIf Msg = "no" Then frmBot.sckFurc.SendData "wh " & Furre & " Bot help request cancelled." & vbLf


End If

    If PO = Furre Then
        If Left(Msg, 5) = "send " Then
            msgArray = Split(Msg, " ")
            For x = 2 To UBound(msgArray)
                If Message <> Empty Then Message = Message & " " & msgArray(x)
                If Message = Empty Then Message = msgArray(x)
            Next
            modPostOffice.Send_Message Furre, msgArray(1), Message
        End If

        If Msg = "get msgs" And PO <> Furre Then modPostOffice.Read_Messages Furre
    
        If Msg Like "read *" Then modPostOffice.Read_Msg Furre, Right(Msg, Len(Msg) - 5)

        If Msg Like "clear *" Then modPostOffice.Clear_Sender Furre, Right(Msg, Len(Msg) - 6)
    
        If Msg = "clear" Then modPostOffice.Clear_Msgs Furre
    End If
    
    
    'If Furre = "Mys'" Then
        
        'If Msg Like "ban *" Then modGuild.Ban_Furre Furre, Right(Msg, Len(Msg) - 4)

        'If Msg Like "share *" Then modGuild.Share_Furre Furre, Right(Msg, Len(Msg) - 6)

        'If Msg Like "reg *" Then modGuild.Register_Furre Furre, Right(Msg, Len(Msg) - 4)

        'If Msg Like "del *" Then modGuild.Delete_Furre Furre, Right(Msg, Len(Msg) - 4)
    'End If
    
End Sub
