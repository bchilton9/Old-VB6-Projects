Attribute VB_Name = "delete"
Dim messnum As Integer

Sub dodelete(Furre, snd)
    messnum = 0
    
    Open "messages.dat" For Input As #1
    Open "temp.dat" For Output As #2
    Do
    Input #1, mto, mfrom, mess
    If mto = Furre Then
        messnum = messnum + 1
            If messnum = snd Then
                frmBot.sckFurc.SendData "wh " & Furre & " Message deleted" & vbLf
            Else
                Write #2, mto, mfrom, mess
            End If
    Else
    Write #2, mto, mfrom, mess
    End If
    Loop Until EOF(1)
    Close #1, #2
    
    
    Open "messages.dat" For Output As #1
    Open "temp.dat" For Input As #2
    Do
    Input #2, mto, mfrom, mess
    Write #1, mto, mfrom, mess
    Loop Until EOF(2)
    Close #1, #2
    
    Open "temp.dat" For Output As #2
    Write #2, ""
    Close #2
End Sub

