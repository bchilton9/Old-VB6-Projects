Attribute VB_Name = "chat"
Sub Text(txt As String, Color)

Msg = MsgBox("Need to fix the text for " & txt, vbCritical)

End Sub


Sub HTMLToRich(inHTML As String)


Dim lngLastFontColor As Long
Dim lngChar As Long, strTag As String, lngSpot As Long, strChar As String
Dim strBuf As String, strBuf2 As String, lngBuf As Long, strBuf3 As String
Dim strHTML As String

strHTML$ = vbLf & inHTML

lngLastFontColor& = -1

   For lngChar& = 1 To Len(strHTML$)
      strChar$ = Mid$(strHTML$, lngChar&, 1)
      
        If strChar$ = "<" Then
           lngSpot& = InStr(lngChar& + 1, strHTML$, ">")
              If lngSpot& Then
              
                 strTag$ = LCase$(Mid$(strHTML$, lngChar& + 1, lngSpot& - lngChar& - 1))
                    
                   If LCase(Left$(strTag$, 5)) = "font " Then
                      strBuf$ = Right$(strTag$, Len(strTag$) - 5)

                               strBuf$ = Right$(strBuf$, Len(strBuf$) - InStr(strBuf$, "="))
                               strBuf3$ = ""
                                  For lngBuf& = 1 To Len(strBuf$)
                                     strBuf2$ = Mid$(strBuf$, lngBuf&, 1)
                                     If strBuf2$ <> """" And strBuf2$ <> "#" Then strBuf3$ = strBuf3$ & strBuf2$
                                  Next lngBuf&
                               lngLastFontColor& = HexToDecimal(strBuf3$)

                   End If
                   
                 lngChar& = lngSpot&
              End If
        Else
           frmMain.txtChat.SelStart = Len(frmMain.txtChat.Text)
           frmMain.txtChat.SelLength = 0
           frmMain.txtChat.SelText = strChar$
           frmMain.txtChat.SelStart = Len(frmMain.txtChat.Text) - 1
           frmMain.txtChat.SelLength = 1
           frmMain.txtChat.SelColor = lngLastFontColor&
        End If
     

      
    Next lngChar&


End Sub

Function HexToDecimal(ByVal strHex As String) As Long

Dim lngDecimal As Long, strCharHex As String, lngColor As Long
Dim lngChar As Long

If Left$(strHex$, 1) = "#" Then strHex$ = Right$(strHex$, 6)
  
strHex$ = Right$(strHex$, 2) & Mid$(strHex$, 3, 2) & Left$(strHex$, 2)

  For lngChar& = Len(strHex$) To 1 Step -1
    strCharHex$ = Mid$(UCase$(strHex$), lngChar&, 1)
    
       Select Case strCharHex$
          Case 0 To 9
             lngDecimal& = CLng(strCharHex$)
          Case Else 'A,B,C,D,E,F
             lngDecimal& = CLng(Chr$((Asc(strCharHex$) - 17))) + 10
       End Select
       
    lngColor& = lngColor& + lngDecimal& * 16 ^ (Len(strHex$) - lngChar&)
  Next lngChar&
  
HexToDecimal = lngColor&

End Function
