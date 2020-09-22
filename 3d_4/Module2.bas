Attribute VB_Name = "Module2"
Private Declare Function CreatePolyPolygonRgn Lib "gdi32" (lpPoint As pointApi, lpPolyCounts As Long, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function PaintRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Public Type pointApi
 x As Long
 y As Long
End Type

Private Type vertexAry
 vtx(0 To 3) As pointApi
 clr As Long
 depth As Single
End Type

Function polyDraw(pointArray() As renderStruct)
 Dim i As Integer, j As Integer, vertAry(5) As vertexAry, tmpVert As vertexAry
 
 tStr = "124356871265348715732684" ' 8 cube corners making 6 sides
 For i = 0 To 5 ' 6 sides
  vertAry(i).depth = -9999
  For j = 0 To 3 ' 4 corners per side
   With pointArray(Val(Mid(tStr, i * 4 + j + 1, 1)) - 1)
    vertAry(i).vtx(j).x = .x
    vertAry(i).vtx(j).y = .y
    If .z > vertAry(i).depth Then
     vertAry(i).depth = .z
    End If
   End With
  Next j
  vertAry(i).clr = RGB(i * 10 + 128, i * 12 + 128, i * 14 + 128)
 Next i

' simple sort to draw back shapes first
 For i = 0 To 4
  For j = i + 1 To 5
   If vertAry(i).depth < vertAry(j).depth Then
    tmpVert = vertAry(i)
    vertAry(i) = vertAry(j)
    vertAry(j) = tmpVert
   End If
  Next j
  drawAPoly vertAry(i)
 Next i
 drawAPoly vertAry(5)

 Form1.Picture1.Refresh

End Function

Function drawAPoly(vtxary As vertexAry)
 Dim hRgn As Long, cnt(0 To 0) As Long
 
 Form1.Picture1.FillColor = vtxary.clr
 cnt(0) = 4
 hRgn = CreatePolyPolygonRgn(vtxary.vtx(0), cnt(0), 1, 2)
 PaintRgn Form1.Picture1.hdc, hRgn
 DeleteObject hRgn

End Function
