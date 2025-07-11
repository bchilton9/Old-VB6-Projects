Attribute VB_Name = "Maps"
Public obs(0 To 11, 0 To 11) As String
Public cur As Integer

Sub inputMaps()
On Error Resume Next

For i = 0 To 54



frmtex.tex(i) = LoadPicture(App.Path & "\tiles\" & i & ".bmp")

Next i

End Sub

Sub drawMap(map As Integer)
On Error Resume Next

Dim strMap As String
Open App.Path & "\Maps\map" & map & ".txt" For Input As #1
For j = 1 To 12
    For i = 1 To 12
    Input #1, strMap
    Form1.picMap.PaintPicture frmtex.tex(strMap).Picture, (j - 1) * 30, (i - 1) * 30
    Next i
Next j
Close #1

End Sub

Sub inputobst(map As Integer)
On Error Resume Next

Dim strMap As String
Open App.Path & "\Maps\map" & map & ".txt" For Input As #1
For l = 1 To 2
For i = 0 To 11
    For j = 0 To 11
    Input #1, strMap
    If l = 2 Then obs(i, j) = strMap
    Next j
Next i
Next l
Close #1
End Sub
