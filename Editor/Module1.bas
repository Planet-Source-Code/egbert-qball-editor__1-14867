Attribute VB_Name = "Types_Declares"
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Public Type BrickS
X As Long
Y As Long
Height As Long
Width As Long
Visible As Boolean
Pic As Long
End Type
