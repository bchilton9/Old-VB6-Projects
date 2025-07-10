Attribute VB_Name = "Maps"
Type Item
loc As Integer
X As Single
Y As Single
obName As String
type As Integer
field1 As Variant
field2 As Variant
field3 As Variant
pic As Integer
End Type

Public items(0 To 20) As Item
Public obs(0 To 11, 0 To 11) As String
Sub inputMaps()
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
inputItems

inputobst (Guy.Location)
End Sub

Sub inputobst(map As Integer)
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
Sub inputItems()
Open App.Path & "\Misc\items.txt" For Input As #1
For i = 1 To 4
Input #1, items(i).loc, items(i).X, items(i).Y, items(i).obName, items(i).type, items(i).field1, items(i).field2, items(i).field3, items(i).pic
Next i
Close #1
End Sub
