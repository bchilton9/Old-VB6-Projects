Attribute VB_Name = "modGuild"
Public Membership As Boolean, HelpReq As Boolean, Send As Boolean, Receive As Boolean, TransCheck As Boolean, Share As Boolean, Eject As Boolean
Public Member As String, Applicant As String
Public Points As Integer, sAmnt As Integer, rAmnt As Integer, Amount As Integer

Sub Membership_Check(Membership, Member)
    Open "members.txt" For Input As #1
        Do Until EOF(1)
            Line Input #1, fName
            If Member = fName Then Membership = True
            If MemList <> Empty Then MemList = MemList & vbLf & fName
            If MemList = Empty Then MemList = fName
        Loop
    Close #1
End Sub

Sub Help_Request(Choice, HelpReq, Furre)
    frmBot.sckFurc.SendData "wh " & Furre & " Your request for bot help has been sent. A member of the AICC *may* contact you in a few moments. Please do not abuse the help request privelege." & vbLf
    Open "members.txt" For Input As #1
        Do Until EOF(1)
            Line Input #1, Helper
            frmBot.sckFurc.SendData "wh " & Helper & " " & Furre & " has requested help with a bot. Please contact him/her as soon as possible. If you are unable to respond to bot help requests at the current time, put me on ignore." & vbLf
        Loop
    Close #1
    HelpReq = False
End Sub

Sub Ban_Furre(Furre, Criminal)
    Open "eject.txt" For Append As #1
        Print #1, Right(Msg, Len(Msg) - 4)
    Close #1
    sckFurc.SendData "wh " & Furre & " " & Right(Msg, Len(Msg) - 4) & " has been banned from all AICC dreams." & vbLf
End Sub

Sub Share_Furre(Furre, Taneest)
    Open "share.txt" For Append As #1
        Print #1, Right(Msg, Len(Msg) - 4)
        sckFurc.SendData "wh " & Furre & " " & Right(Msg, Len(Msg) - 4) & " has been added to the share list for all AICC dreams." & vbLf
    Close #1
End Sub

Sub Register_Furre(Furre, Applicant)
    Membership_Check Membership, Applicant
    If Membership = True Then
        frmBot.sckFurc.SendData "wh " & Furre & " " & Applicant & " is already a member." & vbLf
    Else
        Open "members.txt" For Append As #1
            Print #1, Applicant
        Close #1
            Open "accounts.txt" For Append As #1
            Print #1, Applicant & ": 0"
        Close #1
        Do Until NameLength = Len(Applicant)
            NameLength = NameLength + 1
            If Mid(Applicant, NameLength, 1) = "|" Then Mid(Applicant, NameLength, 1) = "_"
        Loop
        NameLength = 0
        Open Applicant & ".txt" For Output As #1 Len = 50
        Close #1
        frmBot.sckFurc.SendData "wh " & Furre & " " & Applicant & " is now a member of the AICC." & vbLf
        Do Until NameLength = Len(Applicant)
            NameLength = NameLength + 1
            If Mid(Applicant, NameLength, 1) = "_" Then Mid(Applicant, NameLength, 1) = "|"
        Loop
        NameLength = 0
        frmBot.sckFurc.SendData "wh " & Applicant & " You are now a member of the AICC." & vbLf
    End If
Membership = False
End Sub

Sub Delete_Furre(Furre, XMember)
    Membership_Check Membership, XMember
    If Membership = True Then
        Open "members.txt" For Input As #1
            Do Until EOF(1)
                Line Input #1, Mem
                If MemList <> Empty Then MemList = MemList & vbLf & Mem
                If MemList = Empty Then MemList = Mem
            Loop
        Close #1
        
        Open "accounts.txt" For Input As #1
            Do Until EOF(1)
                Line Input #1, Acct
                If AcctList <> Empty Then AcctList = AcctList & vbLf & Acct
                If AcctList = Empty Then AcctList = Acct
            Loop
        Close #1
        Mems = Split(MemList, vbLf)
        Accts = Split(AcctList, vbLf)
        Open "members.txt" For Output As #1 Len = Len(XMember)
        Open "accounts.txt" For Output As #2 Len = Len(XMember)
        For Each Mem In Mems
        If Mem <> XMember Then Print #1, Mem
        Next
        For Each Account In Accts
        If Left(Account, Len(XMember)) <> XMember Then Print #2, Account
        Next
        Close #1
        Close #2
        Do Until NameLength = Len(XMember)
            NameLength = NameLength + 1
            If Mid(XMember, NameLength, 1) = "|" Then Mid(XMember, NameLength, 1) = "_"
        Loop
        NameLength = 0
        Kill XMember & ".txt"
        Do Until NameLength = Len(XMember)
            NameLength = NameLength + 1
            If Mid(XMember, NameLength, 1) = "_" Then Mid(XMember, NameLength, 1) = "|"
        Loop
        NameLength = 0
        frmBot.sckFurc.SendData "wh " & Furre & " " & XMember & " is no longer a member of the AICC." & vbLf
    End If
Membership = False
End Sub

Sub Transfer_Points(Sender, Amount, Recipient)
    Membership = False
    Membership_Check Membership, Sender
    If Membership = True Then Send = True
    
    Membership = False
    Membership_Check Membership, Sender
    If Membership = True Then Receive = True

    If Send = True And Receive = False Then frmBot.sckFurc.SendData "wh " & Sender & " " & Recipient & " is not a member of the AICC." & vbLf
    If Send = False And Receive = True Then frmBot.sckFurc.SendData "wh " & Sender & " You are not a member of the AICC." & vbLf
    If Send = False And Receive = False Then frmBot.sckFurc.SendData "wh " & Sender & " You nor " & Recipient & " are members of the AICC." & vbLf

    If Send = True And Receive = True Then
        TransCheck = True
        Balance Sender
        sAmnt = Points
        Balance Recipient
        rAmnt = Points
        TransCheck = False
        If sAmnt >= Amount Then
            sAmnt = sAmnt - Amount
            rAmnt = rAmny + Amount
            
            Open "accounts.txt" For Input As #1
                Do Until EOF(1)
                    Line Input #1, Acct
                    If AcctList <> Empty Then AcctList = AcctList & vbLf & Acct
                    If AcctList = Empty Then AcctList = Acct
                Loop
            Close #1
            
            Accts = Split(AcctList, vbLf)
                
            
            Open "accounts.txt" For Output As #1
                For Each NewAcct In Accts
                    If Left(NewAcct, Len(Sender)) = Sender Then NewAcct = Sender & ": " & sAmnt
                    If Left(NewAcct, Len(Recipient)) = Recipient Then NewAcct = Recipient & ": " & rAmnt
                    If NewAcct <> pAcct Then Print #1, NewAcct
                    pAcct = NewAcct
                Next
            Close #1
            frmBot.sckFurc.SendData "wh " & Sender & " You have transfered " & Amount & " points to " & Recipient & "'s account, leaving you with " & sAmnt & "." & vbLf
            frmBot.sckFurc.SendData "wh " & Recipient & " " & Sender & " has transfered " & Amount & " points to your account, leaving you with " & rAmnt & "." & vbLf
        Else
            frmBot.sckFurc.SendData "wh " & Sender & " You do not have enough points to make this transaction." & vbLf
        End If
    End If
End Sub

Sub Balance(Member)
    Membership_Check Membership, Member
    If Membership = True Then
        Open "accounts.txt" For Input As #1
            Do Until EOF(1)
                Line Input #1, AcctInfo
                If Left(AcctInfo, Len(Member)) = Member Then Points = Right(AcctInfo, Len(AcctInfo) - (Len(Member) + 2))
            Loop
        Close #1
        If TransCheck = False Then frmBot.sckFurc.SendData "wh " & Member & " You have " & Points & " points in your account." & vbLf
    End If
    Membership = False
End Sub

Sub List_Members(Furre)
    Open "members.txt" For Input As #1
        Do Until EOF(1)
            Line Input #1, Mem
            If lstMem <> Empty Then lstMem = lstMem & ", " & Mem
            If lstMem = Empty Then lstMem = Mem
        Loop
    Close #1
    frmBot.sckFurc.SendData "wh " & Furre & " The following furres are members of the AICC: " & lstMem & "." & vbLf
End Sub

