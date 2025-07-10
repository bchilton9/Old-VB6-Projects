Attribute VB_Name = "modSigns"
Public Sign As Integer
Public PO As String, BF As String
Public Prgm As Boolean

Sub SignTrg(Trg, Pos, TrgPos)
'Entry whisper
If Trg = "!" Then frmBot.sckFurc.SendData "l  @!)" & vbLf: Sign = 1
'Help request at help desk
If Pos = " J {" Then frmBot.sckFurc.SendData "l " & Pos & vbLf: Sign = 2
'FurBot Sign
If Pos = " P o" Then frmBot.sckFurc.SendData "l " & Pos & vbLf: Sign = 3
'FurBot Shelves - Left to Right
If Pos = " P m" Then frmBot.sckFurc.SendData "l " & Pos & vbLf: Sign = 4
If Pos = " P l" Then frmBot.sckFurc.SendData "l " & Pos & vbLf: Sign = 5
If TrgPos = " P j" And Pos = " Q k" Then frmBot.sckFurc.SendData "l " & Pos & vbLf: Sign = 6
If TrgPos = " Q j" And Pos = " Q k" Then frmBot.sckFurc.SendData "l " & Pos & vbLf: Sign = 7
If Pos = " Q l" Then frmBot.sckFurc.SendData "l " & Pos & vbLf: Sign = 8
If Pos = " R m" Then frmBot.sckFurc.SendData "l " & Pos & vbLf: Sign = 9
If Pos = " R n" Then frmBot.sckFurc.SendData "l " & Pos & vbLf: Sign = 10
'zMUD Sign 1
If Pos = " N k" Then frmBot.sckFurc.SendData "l " & Pos & vbLf: Sign = 11
'zMUD Shelves - Left to Right
If Pos = " N i" Then frmBot.sckFurc.SendData "l " & Pos & vbLf: Sign = 12
If Pos = " N h" Then frmBot.sckFurc.SendData "l " & Pos & vbLf: Sign = 13
If Pos = " O g" Then frmBot.sckFurc.SendData "l " & Pos & vbLf: Sign = 14
If Pos = " O f" Then frmBot.sckFurc.SendData "l " & Pos & vbLf: Sign = 15
If Pos = " P e" Then frmBot.sckFurc.SendData "l " & Pos & vbLf: Sign = 16
If Pos = " P d" Then frmBot.sckFurc.SendData "l " & Pos & vbLf: Sign = 17
If Pos = " Q c" Then frmBot.sckFurc.SendData "l " & Pos & vbLf: Sign = 18
If Pos = " Q c" Then frmBot.sckFurc.SendData "l " & Pos & vbLf: Sign = 19
If Pos = " Q d" Then frmBot.sckFurc.SendData "l " & Pos & vbLf: Sign = 20
If Pos = " R e" Then frmBot.sckFurc.SendData "l " & Pos & vbLf: Sign = 21
If Pos = " R f" Then frmBot.sckFurc.SendData "l " & Pos & vbLf: Sign = 22
If Pos = " S g" Then frmBot.sckFurc.SendData "l " & Pos & vbLf: Sign = 23
If Pos = " S h" Then frmBot.sckFurc.SendData "l " & Pos & vbLf: Sign = 24
If Pos = " T i" Then frmBot.sckFurc.SendData "l " & Pos & vbLf: Sign = 25
'zMUD Sign 2
If Pos = " T k" Then frmBot.sckFurc.SendData "l " & Pos & vbLf: Sign = 26
'MUSHclient Sign
If Pos = " K h" Then frmBot.sckFurc.SendData "l " & Pos & vbLf: Sign = 27
'MUSHclient Shelves - Left to Right
If Pos = " K f" Then frmBot.sckFurc.SendData "l " & Pos & vbLf: Sign = 28
If Pos = " L e" Then frmBot.sckFurc.SendData "l " & Pos & vbLf: Sign = 29
If Pos = " L d" Then frmBot.sckFurc.SendData "l " & Pos & vbLf: Sign = 30
If Pos = " M c" Then frmBot.sckFurc.SendData "l " & Pos & vbLf: Sign = 31
If Pos = " M b" Then frmBot.sckFurc.SendData "l " & Pos & vbLf: Sign = 32
If Pos = " N a" Then frmBot.sckFurc.SendData "l " & Pos & vbLf: Sign = 33
If Pos = " N `" Then frmBot.sckFurc.SendData "l " & Pos & vbLf: Sign = 34
If Pos = " O _" Then frmBot.sckFurc.SendData "l " & Pos & vbLf: Sign = 35
If Pos = " O ^" Then frmBot.sckFurc.SendData "l " & Pos & vbLf: Sign = 36
If Pos = " P ]" Then frmBot.sckFurc.SendData "l " & Pos & vbLf: Sign = 37
If Pos = " P \" Then frmBot.sckFurc.SendData "l " & Pos & vbLf: Sign = 38
If Pos = " Q [" Then frmBot.sckFurc.SendData "l " & Pos & vbLf: Sign = 39
'Visual Basic Sign
If Pos = " V h" Then frmBot.sckFurc.SendData "l " & Pos & vbLf: Sign = 40
'Visual Basic Shelves - Right to Left
If Pos = " V f" Then frmBot.sckFurc.SendData "l " & Pos & vbLf: Sign = 41
If Pos = " V e" Then frmBot.sckFurc.SendData "l " & Pos & vbLf: Sign = 42
If Pos = " U d" Then frmBot.sckFurc.SendData "l " & Pos & vbLf: Sign = 43
If Pos = " U c" Then frmBot.sckFurc.SendData "l " & Pos & vbLf: Sign = 44
If Pos = " T b" Then frmBot.sckFurc.SendData "l " & Pos & vbLf: Sign = 45
If Pos = " T a" Then frmBot.sckFurc.SendData "l " & Pos & vbLf: Sign = 46
If Pos = " S `" Then frmBot.sckFurc.SendData "l " & Pos & vbLf: Sign = 47
If Pos = " S _" Then frmBot.sckFurc.SendData "l " & Pos & vbLf: Sign = 48
If Pos = " R ^" Then frmBot.sckFurc.SendData "l " & Pos & vbLf: Sign = 49
If Pos = " R ]" Then frmBot.sckFurc.SendData "l " & Pos & vbLf: Sign = 50
If Pos = " Q \" Then frmBot.sckFurc.SendData "l " & Pos & vbLf: Sign = 51
If Pos = " Q [" Then frmBot.sckFurc.SendData "l " & Pos & vbLf: Sign = 52
'Bank
If Pos = " F X" Then frmBot.sckFurc.SendData "l " & Pos & vbLf: Sign = 53
If Pos = " G Y" Then frmBot.sckFurc.SendData "l " & Pos & vbLf: Sign = 54
'Post Office Sign
If Pos = " \ c" And TrgPos = " \ b" Then frmBot.sckFurc.SendData "l " & Pos & vbLf: Sign = 55
'Post Office
If Pos = " ] c" Then frmBot.sckFurc.SendData "l " & Pos & vbLf: Sign = 56
If Pos = " \ d" Then frmBot.sckFurc.SendData "l " & Pos & vbLf: Sign = 57
'Help Desk Teleport
If Trg = "1" Then frmBot.sckFurc.SendData "l " & Pos & vbLf: Sign = 58
'Meeting Area Sign
If Pos = " P H" Then frmBot.sckFurc.SendData "l " & Pos & vbLf: Sign = 59
'Bank Sign
If Pos = " D `" Then frmBot.sckFurc.SendData "l " & Pos & vbLf: Sign = 60
'Classroom Under Construction Sign
If Pos = " C c" Then frmBot.sckFurc.SendData "l " & Pos & vbLf: Sign = 61
'If Prgm = True Then frmBot.sckFurc.SendData Chr(34) + "emitloud Sign #" & Sign & vbLf
End Sub

Sub YouSee(Furre, Sign)
done = False
If Sign = 1 Then
done = True
    Open "eject.txt" For Input As #1
        Do Until Furre = Baddie Or EOF(1)
            Line Input #1, Baddie
            If Furre = Baddie Then frmBot.sckFurc.SendData Chr(34) + "eject " & Furre & vbLf: Eject = True
        Loop
    Close #1

    Open "share.txt" For Input As #1
        Do Until Furre = Goodie Or EOF(1)
            Line Input #1, Goodie
            If Furre = Goodie Then frmBot.sckFurc.SendData Chr(34) + "share " & Furre & vbLf: Share = True
        Loop
    Close #1

    If Eject = True Then frmBot.sckFurc.SendData "wh " & Furre & " You are banned from all AICC dreams. If you feel this is a mistake, contact Mys'." & vbLf
    If Share = True Then frmBot.sckFurc.SendData "wh " & Furre & " Hello, " & Furre & ", welcome back to the academy. :)" & vbLf
    If Share = False And Eject = False Then frmBot.sckFurc.SendData "wh " & Furre & " You seem to have stumbled upon The AICC Academy of Botmakers. Within this social guild you can learn how to make bots using a variety of clients. For more information, you can visit our site at [http://www.erenetwork.com/aicc/] or contact one of our members. Whisper me [commands] for a list of options." & vbLf

End If

If Sign = 2 Then
done = True
frmBot.sckFurc.SendData "wh " & Furre & " Do you need help with a bot? (yes/no)" & vbLf: HelpReq = True
End If

If Sign = 58 Then
done = True
    Open "members.txt" For Input As #1
        Do Until EOF(1)
            Line Input #1, Helper
            If Furre = Helper Then Teleport = True
        Loop
    Close #1
    If Teleport = True Then frmBot.sckFurc.SendData "wh " & Furre & " Thank you for taking the time to help furres at the Bot Help Desk. #SA" & vbLf: frmBot.sckFurc.SendData "use" & vbLf
    If Teleport = False Then frmBot.sckFurc.SendData "wh " & Furre & " Only members may enter the Help Desk." & vbLf: frmBot.sckFurc.SendData "m 1" & vbLf
End If

If Sign = 3 Then
done = True
frmBot.sckFurc.SendData "wh " & Furre & " You can find the URLs to several FurBot tutorials on these shelves." & vbLf
ElseIf Sign = 4 Then
done = True
frmBot.sckFurc.SendData "wh " & Furre & " Create & Connect a FurBot [http://www.erenetwork.com/aicc/Tutorials/FurBot/NewBot.htm]" & vbLf
ElseIf Sign = 5 Then
done = True
frmBot.sckFurc.SendData "wh " & Furre & " Auto-whispers with FurBot [http://www.erenetwork.com/aicc/Tutorials/FurBot/Whispers.htm]" & vbLf
ElseIf Sign = 6 Then
done = True
frmBot.sckFurc.SendData "wh " & Furre & " DS Responses with FurBot [http://www.erenetwork.com/aicc/Tutorials/FurBot/DSresponse.htm]" & vbLf
ElseIf Sign = 7 Then
done = True
frmBot.sckFurc.SendData "wh " & Furre & " Signs With FurBot [http://www.erenetwork.com/aicc/Tutorials/FurBot/Signs.htm]" & vbLf
ElseIf Sign = 8 Then
done = True
frmBot.sckFurc.SendData "wh " & Furre & " Logs with FurBot [http://www.erenetwork.com/aicc/Tutorials/FurBot/Logs.htm]" & vbLf
ElseIf Sign = 9 Then
done = True
frmBot.sckFurc.SendData "wh " & Furre & " Databases with FurBot [http://www.erenetwork.com/aicc/Tutorials/FurBot/Databases.htm]" & vbLf
ElseIf Sign = 10 Then
done = True
frmBot.sckFurc.SendData "wh " & Furre & " Store a Roll Value [http://www.erenetwork.com/aicc/Tutorials/FurBot/RollValue.htm]" & vbLf
ElseIf Sign = 11 Then
done = True
frmBot.sckFurc.SendData "wh " & Furre & "  You can the URLs to several zMUD tutorials on these shelves." & vbLf
ElseIf Sign = 26 Then
done = True
frmBot.sckFurc.SendData "wh " & Furre & "  You can the URLs to several zMUD tutorials on these shelves." & vbLf
ElseIf Sign = 27 Then
done = True
frmBot.sckFurc.SendData "wh " & Furre & " You can the URLs to several MUSHclient tutorials on these shelves." & vbLf
ElseIf Sign = 28 Then
done = True
frmBot.sckFurc.SendData "wh " & Furre & " Create & Connect a MUSHclient bot [http://www.erenetwork.com/aicc/Tutorials/MUSHclient/NewBot.htm]" & vbLf
ElseIf Sign = 29 Then
done = True
frmBot.sckFurc.SendData "wh " & Furre & " Basic Commands with MUSHclient [http://www.erenetwork.com/aicc/Tutorials/MUSHclient/BasicCommands.htm]" & vbLf
ElseIf Sign = 30 Then
done = True
frmBot.sckFurc.SendData "wh " & Furre & " Wildcards with MUSHclient [http://www.erenetwork.com/aicc/Tutorials/Wildcards/Wildcards.htm]" & vbLf
ElseIf Sign = 31 Then
done = True
frmBot.sckFurc.SendData "wh " & Furre & " VBScripting for Beginners with MUSHclient [http://www.erenetwork.com/aicc/Tutorials/MUSHclient/VBS4Beginners.htm]" & vbLf
ElseIf Sign = 40 Then
done = True
frmBot.sckFurc.SendData "wh " & Furre & " We do not currently have any Visual Basic tutorials available." & vbLf
ElseIf Sign = 55 Then
done = True
frmBot.sckFurc.SendData "wh " & Furre & " This is the Academy Post Office. You can leave messages for a member for them to read later. Step on the pad to read or send a message." & vbLf
ElseIf Sign = 56 Then
done = True
frmBot.sckFurc.SendData "wh " & Furre & " Whisper me 'get msgs', without the quotes, to get a list of your messages. Whisper me 'read Furre' to read messages from a certain furre. Whisper 'send Furre Msg' to me to send a message to a certain furre." & vbLf
ElseIf Sign = 57 Then
done = True
frmBot.sckFurc.SendData "wh " & Furre & " #SE Thanks for using the Academy Post Office. #SE" & vbLf
ElseIf Sign = 59 Then
done = True
frmBot.sckFurc.SendData "wh " & Furre & " This is the Academy Meeting Area. We plan to have many meetings, Q&A sessions, and other such gatherings here." & vbLf
ElseIf Sign = 60 Then
done = True
frmBot.sckFurc.SendData "wh " & Furre & " This is the Academy Bank. Members can transfer funds or check the balance of their accounts here." & vbLf
ElseIf Sign = 53 Then
done = True
frmBot.sckFurc.SendData "wh " & Furre & " Welcome to the Academy Bank! To transfer points to another member, whisper 'transfer Amount to Furre' to me. To check your account balance, whisper 'balance' to me." & vbLf
ElseIf Sign = 54 Then
done = True
frmBot.sckFurc.SendData "wh " & Furre & " #SE Thank you for using the AICC Academy bank! #SE" & vbLf
ElseIf Sign = 61 Then
done = True
frmBot.sckFurc.SendData "wh " & Furre & " This stairwell leads to the main socializing area and sparring arena." & vbLf
ElseIf Sign = 12 Then
done = True
frmBot.sckFurc.SendData "wh " & Furre & " Create & Connect a zMUD bot [http://www.erenetwork.com/aicc/Tutorials/zMUD/connect/connect.html]" & vbLf
ElseIf Sign = 13 Then
done = True
frmBot.sckFurc.SendData "wh " & Furre & " Basic Commands with zMUD [http://www.erenetwork.com/aicc/Tutorials/zMUD/move/move.html]" & vbLf
ElseIf Sign = 14 Then
done = True
frmBot.sckFurc.SendData "wh " & Furre & " Loading a bot with zMUD [http://www.erenetwork.com/aicc/Tutorials/zMUD/load/load.html]" & vbLf
ElseIf Sign = 15 Then
done = True
frmBot.sckFurc.SendData "wh " & Furre & " Triggers with zMUD [http://www.erenetwork.com/aicc/Tutorials/zMUD/triggers/triggers.html]" & vbLf
ElseIf Sign = 16 Then
done = True
frmBot.sckFurc.SendData "wh " & Furre & " Signs with zMUD [http://www.erenetwork.com/aicc/Tutorials/zMUD/signs/signs.html]" & vbLf
ElseIf Sign = 17 Then
done = True
frmBot.sckFurc.SendData "wh " & Furre & " 8-Ball Commands with zMUD [http://www.erenetwork.com/aicc/Tutorials/zMUD/8ball/8ball.html]" & vbLf


End If

If done = False Then frmBot.sckFurc.SendData "wh " & Furre & " Information not available." & vbLf


Teleport = False
Share = False
Eject = False
End Sub


