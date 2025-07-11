Attribute VB_Name = "timer"
Sub timeon()
On Error GoTo reconect
        frmBot.Space = frmBot.Space + 1
        If frmBot.Space = 5 Then
            frmBot.sckFurc.SendData "flame" & vbLf
            frmBot.Space = 0
        End If
        
        
'        frmBot.urgc = frmBot.urgc + 1
'        If frmBot.urgc >= 30 Then
'            frmBot.urgc = 0
'            Open "C:\mailsys\suggestq.txt" For Input As #1
'            Input #1, qun
'            Close #1
'            If qun >= 1 Then
'                Open "C:\mailsys\suggest.txt" For Input As #1
'                Input #1, fName, sndr
'                frmBot.sckFurc.SendData "wh " & fName & " " & sndr & " Suggested that you join MailSys Whisper me #HELP to learn how to use my service. If you dont want to recive thease anymore whisper me REMOVE." & vbLf
'                Do Until (EOF(1))
'                Input #1, fName, sndr
'                frmBot.sckFurc.SendData "wh " & fName & " " & sndr & " Suggested that you join MailSys Whisper me #HELP to learn how to use my service. If you dont want to recive thease anymore whisper me REMOVE." & vbLf
'                Loop
'                Close #1
'            End If
'        End If
   
       
       
    If frmBot.prem = 1 Then
        frmBot.premt = frmBot.premt + 1
        If frmBot.premt >= 10 Then
            Open "C:\mailsys\prem.txt" For Input As #1
            Input #1, premote
            Close #1
            frmBot.sckFurc.SendData Chr(34) & premote & vbLf
            frmBot.premt = 0
        End If
    End If

    Open "memnum.txt" For Input As #1
    Input #1, nnum
    Close #1
    Open "sent.txt" For Input As #3
    Input #3, sent
    Close #3
    frmBot.txtsent = sent
    frmBot.txtmem = nnum
    
    frmBot.Minute = frmBot.Minute + 1
    If frmBot.Minute >= 60 Then
        frmBot.Hour = frmBot.Hour + 1
        frmBot.Minute = 0
    End If
    If frmBot.Hour >= 24 Then
        frmBot.Day = frmBot.Day + 1
        frmBot.Hour = 0
    End If
    frmBot.timon = frmBot.Day & ":" & frmBot.Hour & ":" & frmBot.Minute
    frmBot.sckFurc.SendData "desc " & frmBot.descrip & " [Uptime: "
    If frmBot.Day >= 1 Then frmBot.sckFurc.SendData frmBot.Day & " Day(s) "
    If frmBot.Hour >= 1 Then frmBot.sckFurc.SendData frmBot.Hour & " Hour(s) "
    frmBot.sckFurc.SendData frmBot.Minute & " Minute(s)]" & vbLf
Exit Sub
reconect:
    frmBot.sckFurc.Close
    frmBot.Connected = False
    frmBot.txtcnt = "False"
    frmBot.sckFurc.RemoteHost = frcHost
    frmBot.sckFurc.RemotePort = frcPort
    frmBot.sckFurc.Connect

End Sub
