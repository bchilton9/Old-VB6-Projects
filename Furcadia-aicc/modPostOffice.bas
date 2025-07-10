Attribute VB_Name = "modPostOffice"

Sub Send_Message(Sender, Recipient, Message)
    Debug.Print "Send Activated"
    Membership_Check Membership, Recipient
    If Membership = True Then
        Do Until NameLength = Len(Recipient)
            NameLength = NameLength + 1
            If Mid(Recipient, NameLength, 1) = "|" Then Mid(Recipient, NameLength, 1) = "_"
        Loop
        NameLength = 0
        Open Recipient & ".txt" For Append As #1
            Print #1, Sender & ": " & Message
        Close #1
        Do Until NameLength = Len(Recipient)
            NameLength = NameLength + 1
            If Mid(Recipient, NameLength, 1) = "_" Then Mid(Recipient, NameLength, 1) = "|"
        Loop
        NameLength = 0
        frmPC.sckFurc.SendData "wh " & Sender & " Your message has been delivered to " & Recipient & "." & vbLf
    Else
        frmPC.sckFurc.SendData "wh " & Sender & " " & Recipient & " is not a member of the AICC." & vbLf
    End If
    Membership = False
    Message = Empty
End Sub

Sub Read_Messages(Reader)
    Debug.Print "Get Activated"
Dim lstMsgs As String
    Membership_Check Membership, Reader
    If Membership = True Then
        Do Until NameLength = Len(Reader)
            NameLength = NameLength + 1
            If Mid(Reader, NameLength, 1) = "|" Then Mid(Reader, NameLength, 1) = "_"
        Loop
        NameLength = 0
        Open Reader & ".txt" For Input As #1
            Do Until EOF(1)
                Line Input #1, MsgInfo
                pMsg = Split(MsgInfo, ": ")
                If lstMsgs <> Empty Then lstMsgs = lstMsgs & ", " & pMsg(0)
                If lstMsgs = Empty Then lstMsgs = pMsg(0)
            Loop
        Do Until NameLength = Len(Reader)
            NameLength = NameLength + 1
            If Mid(Reader, NameLength, 1) = "_" Then Mid(Reader, NameLength, 1) = "|"
        Loop
        NameLength = 0
        Close #1
        frmPC.sckFurc.SendData "wh " & Reader & " You have messages from the following furres: " & lstMsgs & "." & vbLf
    End If
    lstMsgs = Empty
    Membership = False
End Sub

Sub Read_Msg(Reader, Sender)
    Debug.Print "Read Activated"
Dim lstMsgs As String
    Membership_Check Membership, Reader
    If Membership = True Then
        Do Until NameLength = Len(Reader)
            NameLength = NameLength + 1
            If Mid(Reader, NameLength, 1) = "|" Then Mid(Reader, NameLength, 1) = "_"
        Loop
        NameLength = 0
        Open Reader & ".txt" For Input As #1
            Do Until EOF(1)
                Line Input #1, Amsg
                If lstMsgs = Empty Then lstMsgs = Amsg
                If lstMsgs <> Empty Then lstMsgs = lstMsgs & vbLf & Amsg
                Do Until NameLength = Len(Reader)
                    NameLength = NameLength + 1
                    If Mid(Reader, NameLength, 1) = "_" Then Mid(Reader, NameLength, 1) = "|"
                Loop
                NameLength = 0
                If Left(Amsg, Len(Sender)) = Sender Then frmPC.sckFurc.SendData "wh " & Reader & " The following is a message from " & Sender & ": " & Right(Amsg, Len(Amsg) - (Len(Sender) + 2)) & vbLf
            Loop
        Close #1
        
        Msgs = Split(lstMsgs, vbLf)
        
        Do Until NameLength = Len(Reader)
            NameLength = NameLength + 1
            If Mid(Reader, NameLength, 1) = "|" Then Mid(Reader, NameLength, 1) = "_"
        Loop
        NameLength = 0
        
        Open Reader & ".txt" For Output As #1
            For Each Omsg In Msgs
                If Left(Omsg, Len(Sender)) <> Sender Then Print #1, Omsg
            Next
        Close #1
        
        Do Until NameLength = Len(Reader)
            NameLength = NameLength + 1
            If Mid(Reader, NameLength, 1) = "_" Then Mid(Reader, NameLength, 1) = "|"
        Loop
        NameLength = 0
    
    End If
    Membership = False
End Sub

Sub Clear_Sender(Reader, Sender)
    Membership_Check Membership, Reader
    If Membership = True Then
        Do Until NameLength = Len(Reader)
            NameLength = NameLength + 1
            If Mid(Reader, NameLength, 1) = "|" Then Mid(Reader, NameLength, 1) = "_"
        Loop
        NameLength = 0
        Open Reader & ".txt" For Input As #1
            Do Until EOF(1)
                Line Input #1, Dmsg
                If Left(Dmsg, Len(Sender)) <> Sender Then
                    If lstMsgs = Empty Then lstMsgs = Dmsg
                    If lstMsgs <> Empty Then lstMsgs = lstMsgs & vbLf & Dmsg
                End If
            Loop
        Close #1
        
        Msgs = Split(lstMsgs, vbLf)
        
        Open Reader & ".txt" For Output As #1
            For Each Rmsg In Msgs
                If pMsg <> Rmsg Then Print #1, Rmsg
                pMsg = Rmsg
            Next
        Close #1
        Do Until NameLength = Len(Reader)
            NameLength = NameLength + 1
            If Mid(Reader, NameLength, 1) = "_" Then Mid(Reader, NameLength, 1) = "|"
        Loop
        NameLength = 0
        frmPC.sckFurc.SendData "wh " & Reader & " All messages from " & Sender & " have been cleared from your mailbox." & vbLf
    Else
        frmPC.sckFurc.SendData "wh " & Reader & " You are not a member of the AICC." & vbLf
    End If
End Sub

Sub Clear_Msgs(Reader)
    Membership_Check Membership, Reader
    If Membership = True Then
        Do Until NameLength = Len(Reader)
            NameLength = NameLength + 1
            If Mid(Reader, NameLength, 1) = "|" Then Mid(Reader, NameLength, 1) = "_"
        Loop
        NameLength = 0
        Open Reader & ".txt" For Output As #1
            Print #1, Empty
        Close #1
        Do Until NameLength = Len(Reader)
            NameLength = NameLength + 1
            If Mid(Reader, NameLength, 1) = "_" Then Mid(Reader, NameLength, 1) = "|"
        Loop
        NameLength = 0
        frmPC.sckFurc.SendData "wh " & Reader & " Your mailbox has been cleared." & vbLf
    Else
        frmPC.sckFurc.SendData "wh " & Reader & " You are not a member of the AICC." & vbLf
    End If
End Sub
