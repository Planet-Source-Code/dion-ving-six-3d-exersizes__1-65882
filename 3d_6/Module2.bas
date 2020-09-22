Attribute VB_Name = "Module2"
Private Declare Function CreatePolyPolygonRgn Lib "gdi32" (lpPoint As pointApi, lpPolyCounts As Long, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function PaintRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Public Type pointApi
 x As Long
 y As Long
End Type

Public Type vertexAry
 vtx(0 To 3) As pointApi
 clr As Long
 depth As Single
End Type

Function drawAPoly(vtxary As vertexAry)
 Dim hRgn As Long, cnt(0 To 0) As Long
 
 Form1.Picture1.FillColor = vtxary.clr
 cnt(0) = 4
 hRgn = CreatePolyPolygonRgn(vtxary.vtx(0), cnt(0), 1, 2)
 PaintRgn Form1.Picture1.hdc, hRgn
 DeleteObject hRgn

End Function
