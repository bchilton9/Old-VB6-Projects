Attribute VB_Name = "setting"
Sub dosettings()
frmBot.BotName = "MailSys"
frmBot.BotPass = "0519aa"
frmBot.vers = "3.0"
frmBot.descrip = "http://www.erenetwork.com/mailsys                                                 Thank you for choosing Mailsys as your Furcadian mail service. Whisper me one of the following to learn how I run:          #mail, #entertainment, #news, #stats.                                                   or go to the MailSys website.          Press F8 now!!!                          "
frmBot.ColorCode = "! B1+99979 # "
frmBot.frcHost = "66.28.224.193"
frmBot.frcPort = "6000"




Open "C:\mailsys\memnum.txt" For Input As #1
Input #1, nnum
Close #1
Open "C:\mailsys\sent.txt" For Input As #3
Input #3, sent
Close #3
Open "C:\mailsys\errorq.txt" For Input As #3
Input #3, errq
Close #3
frmBot.txterr = errq
frmBot.txtsent = sent
frmBot.txtmem = nnum
frmBot.Minute = 0
frmBot.Day = 0
frmBot.urgc = 0
frmBot.Hour = 0
frmBot.premt = 0
frmBot.Desc = frmBot.descrip & " [Uptime: 0 Minute(s)]"
frmBot.Connected = False
End Sub
