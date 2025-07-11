Attribute VB_Name = "setting"
Sub dosettings()
On Error Resume Next
frmBot.BotName = "MailSys"
frmBot.BotPass = "0519aa"
frmBot.vers = "4.2"
frmBot.descrip = "http://www.erenetwork.com/mailsys                                                 Thank you for choosing Mailsys as your Furcadian mail service. Whisper me one of the following to learn how I run:          #mail, #entertainment, #news, #stats.                                                   or go to the MailSys website.          Press F8 now!!!                          "
frmBot.ColorCode = "! B1+99979 # "
frmBot.frcHost = "66.28.224.193"
frmBot.frcPort = "6000"

frmBot.Minute = 0
frmBot.Day = 0
frmBot.urgc = 0
frmBot.Hour = 0
frmBot.premt = 0
frmBot.Desc = frmBot.descrip & " [Uptime: 0 Minute(s)]"
frmBot.Connected = False

Open "memnum.txt" For Input As #1
Input #1, nnum
Close #1
Open "sent.txt" For Input As #3
Input #3, sent
Close #3
frmBot.txtsent = sent
frmBot.txtmem = nnum

'frmBot.count.Recordset.MoveFirst
'Do Until frmBot.mdbCount.Recordset.EOF
'frmBot.count.Recordset.MoveNext
'If frmBot.txtCname.Text = "Member" Then
'frmBot.txtmem.Text = frmBot.txtVal.Text
'ElseIf frmBot.txtCname.Text = "Sent" Then
'frmBot.txtsent.Text = frmBot.txtVal.Text
'End If
'Loop

End Sub
