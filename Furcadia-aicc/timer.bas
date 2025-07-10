Attribute VB_Name = "timer"
Sub timeon()
    
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
preerr:
    Close #1, #3
    frmBot.sckFurc.Close
    frmBot.Connected = False
    frmBot.txtcnt = "False"
    frmBot.sckFurc.RemoteHost = frmBot.frcHost
    frmBot.sckFurc.RemotePort = frmBot.frcPort
    frmBot.sckFurc.Connect
    Resume stoptrying
descerr:
    Close #1, #3
    frmBot.sckFurc.Close
    frmBot.Connected = False
    frmBot.txtcnt = "False"
    frmBot.sckFurc.RemoteHost = frmBot.frcHost
    frmBot.sckFurc.RemotePort = frmBot.frcPort
    frmBot.sckFurc.Connect
    Resume stoptrying
sugerr:
    Close #1, #3
    frmBot.sckFurc.Close
    frmBot.Connected = False
    frmBot.txtcnt = "False"
    frmBot.sckFurc.RemoteHost = frmBot.frcHost
    frmBot.sckFurc.RemotePort = frmBot.frcPort
    frmBot.sckFurc.Connect
    Resume stoptrying
stoptrying:
End Sub

