Attribute VB_Name = "else"
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Const SRCCOPY = &HCC0020
Public Const SRCPAINT = &HEE0086
Public Const SRCAND = &H8800C6


Public Type Character
name As String
Location As String
Frame As Integer
Animation As Integer
Bitmap As String
Mask As String
Container As Integer
MskContainer As Integer
X As Variant
Y As Variant
Flippers As Boolean
Height As Integer
coins As Integer
Index As Integer
End Type

Public Guy As Character
Public Kiki As Character
Public Olga As Character
Public Merlon As Character
Public coin(1 To 7) As Character
Public sign(1 To 9)

