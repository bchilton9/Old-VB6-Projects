Attribute VB_Name = "modPatcher"
Public Const downloadurl = "http://www.freewebtown.com/tidel/patch"

Public Function GetInternetFile(Inet1 As Inet, myURL As String, DestDIR As String) As Boolean

    'On Local Error GoTo 100
    On Error Resume Next
    Dim myData() As Byte
    If Inet1.StillExecuting = True Then Exit Function
    myData() = Inet1.OpenURL(myURL, icByteArray)


    For X = Len(myURL) To 1 Step -1
        If Left$(Right$(myURL, X), 1) = "/" Then RealFile$ = Right$(myURL, X - 1)
    Next X
    myFile$ = DestDIR + "\" + RealFile$
    Open myFile$ For Binary Access Write As #1
    Put #1, , myData()
    Close #1
    
    GetInternetFile = True
    Exit Function

100 X = MsgBox("An error has occured in the file transfer or write.  Please try again later.", vbInformation)
    GetInternetFile = False
    Resume 105
105 End Function

