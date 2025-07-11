Attribute VB_Name = "addmem"
Sub addmemb(Furre)
    member = 0
    Open "members.mem" For Input As #2
    Do Until EOF(2)
        Input #2, nam, rank
        If nam = Furre Then member = 1
    Loop
    Close #2
    If member = 0 Then
        Open "members.mem" For Append As #1
        Write #1, Furre, 0
        Close #1
        Open "messages.dat" For Append As #1
        Write #1, Furre, "MailSys", "Welcome to MailSys"
        Close #1
        sckFurc.SendData "wh " & Furre & " You have been added to the MailSys members file." & vbLf
    End If
End Sub
