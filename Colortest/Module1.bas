Attribute VB_Name = "Module1"
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, _
ByVal X As Long, _
ByVal Y As Long, _
ByVal nWidth As Long, _
ByVal nHeight As Long, _
ByVal hSrcDC As Long, _
ByVal xSrc As Long, _
ByVal ySrc As Long, _
ByVal nSrcWidth As Long, _
ByVal nSrcHeight As Long, _
ByVal dwRop As Long) As Long

