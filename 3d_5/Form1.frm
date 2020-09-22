VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9105
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12660
   LinkTopic       =   "Form1"
   ScaleHeight     =   9105
   ScaleWidth      =   12660
   StartUpPosition =   3  'Windows Default
   Begin VB.VScrollBar VScroll1 
      Height          =   1155
      LargeChange     =   10
      Left            =   9900
      Max             =   800
      TabIndex        =   20
      Top             =   7620
      Value           =   400
      Width           =   255
   End
   Begin VB.ListBox List1 
      Height          =   1230
      ItemData        =   "Form1.frx":0000
      Left            =   6780
      List            =   "Form1.frx":000D
      TabIndex        =   19
      Top             =   7620
      Width           =   2715
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Index           =   0
      LargeChange     =   10
      Left            =   882
      Max             =   800
      TabIndex        =   9
      Top             =   8070
      Value           =   400
      Width           =   1335
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Index           =   1
      LargeChange     =   10
      Left            =   2841
      Max             =   800
      TabIndex        =   8
      Top             =   8070
      Value           =   400
      Width           =   1335
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Index           =   2
      LargeChange     =   10
      Left            =   4800
      Max             =   800
      TabIndex        =   7
      Top             =   8070
      Value           =   400
      Width           =   1335
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Index           =   5
      LargeChange     =   10
      Left            =   4800
      Max             =   1600
      TabIndex        =   3
      Top             =   8670
      Width           =   1335
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Index           =   4
      LargeChange     =   10
      Left            =   2835
      Max             =   1600
      TabIndex        =   2
      Top             =   8670
      Width           =   1335
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Index           =   3
      LargeChange     =   10
      Left            =   870
      Max             =   800
      TabIndex        =   1
      Top             =   8670
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      FillStyle       =   0  'Solid
      Height          =   7155
      Left            =   180
      ScaleHeight     =   473
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   801
      TabIndex        =   0
      Top             =   300
      Width           =   12075
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Zoom"
      Height          =   195
      Left            =   9810
      TabIndex        =   21
      Top             =   8880
      Width           =   435
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   195
      Index           =   5
      Left            =   5130
      TabIndex        =   18
      Top             =   8400
      Width           =   675
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   195
      Index           =   4
      Left            =   3165
      TabIndex        =   17
      Top             =   8400
      Width           =   675
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   195
      Index           =   3
      Left            =   1200
      TabIndex        =   16
      Top             =   8400
      Width           =   675
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   195
      Index           =   2
      Left            =   5130
      TabIndex        =   15
      Top             =   7800
      Width           =   675
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   195
      Index           =   1
      Left            =   3165
      TabIndex        =   14
      Top             =   7800
      Width           =   675
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   195
      Index           =   0
      Left            =   1215
      TabIndex        =   13
      Top             =   7800
      Width           =   675
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Z Pos"
      Height          =   195
      Index           =   2
      Left            =   4275
      TabIndex        =   12
      Top             =   8100
      Width           =   420
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Y Pos"
      Height          =   195
      Index           =   1
      Left            =   2325
      TabIndex        =   11
      Top             =   8100
      Width           =   420
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "X Pos"
      Height          =   195
      Index           =   0
      Left            =   360
      TabIndex        =   10
      Top             =   8100
      Width           =   420
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Z Axis"
      Height          =   195
      Index           =   5
      Left            =   4275
      TabIndex        =   6
      Top             =   8700
      Width           =   435
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Y Axis"
      Height          =   195
      Index           =   4
      Left            =   2310
      TabIndex        =   5
      Top             =   8700
      Width           =   435
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "X Axis"
      Height          =   195
      Index           =   3
      Left            =   345
      TabIndex        =   4
      Top             =   8700
      Width           =   435
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
 Dim i As Integer, j As Integer, cnt As Integer
 Me.Show
 
 eye = 800

 viewPort.x = Picture1.ScaleWidth / 2
 viewPort.y = Picture1.ScaleHeight / 2
 
 allObjs(1).whd.w = 100
 allObjs(1).whd.h = 100
 allObjs(1).whd.d = 100
 allObjs(1).xyz.x = 0
 allObjs(1).xyz.y = 50
 allObjs(1).xyz.z = 0
 allObjs(1).deg.xDeg = 0
 allObjs(1).deg.yDeg = 0
 allObjs(1).deg.zDeg = 0
 
 allObjs(2).whd.w = 10
 allObjs(2).whd.h = 50
 allObjs(2).whd.d = 60
 allObjs(2).xyz.x = 120
 allObjs(2).xyz.y = 50
 allObjs(2).xyz.z = -80
 allObjs(2).deg.xDeg = 10
 allObjs(2).deg.yDeg = 20
 allObjs(2).deg.zDeg = 30
 
 For i = -6 To 6
  For j = -5 To 5
   clrAry(cnt) = RGB(10 * i + 128, 10 * j + 128, 10 * (i + j) + 128)
   cnt = cnt + 1
  Next j
 Next i
 
 List1.ListIndex = 0
 
End Sub

Function drawWorld()
 Dim xp As Single, zp As Single, x As Single, y As Single, z As Single
 Dim gridCords As xyzPosStruct, tRen(1) As renderStruct, t As Integer
 Dim cnt As Integer, emptyObj As objStruct
 Dim pntsAry(7) As renderStruct
 Dim cords As xyzPosStruct
 Dim tmpObj As objStruct
 Dim vertAry(11) As vertexAry ' 12 shapes in two objects
 Dim tmpVert As vertexAry, idx As Integer
 
 
If 1 = 2 Then ' never true
' draw a grid to show the world space
 For xp = -100 To 100 Step 10
  For zp = -100 To 100 Step 10
   gridCords.x = -100 * 3
   gridCords.z = zp * 3
   tRen(0) = getXY(gridCords)
   gridCords.x = 100 * 3
   tRen(1) = getXY(gridCords)
   Picture1.Line (tRen(0).x, tRen(0).y)-(tRen(1).x, tRen(1).y), vbBlack
   
   gridCords.x = xp * 3
   gridCords.z = -100 * 3
   tRen(0) = getXY(gridCords)
   gridCords.z = 100 * 3
   tRen(1) = getXY(gridCords)
   Picture1.Line (tRen(0).x, tRen(0).y)-(tRen(1).x, tRen(1).y), vbBlack
  Next zp
 Next xp
End If

 
 
 For lp = 1 To 2 ' two shapes
' map 8 points on a cube around this object position
  cnt = 0
  For x = -1 To 1 Step 2
   For y = -1 To 1 Step 2
    For z = -1 To 1 Step 2
     tmpObj = emptyObj
' 1 of 8 cube corners + object origin
     tmpObj.xyz.x = allObjs(lp).whd.w / 2 * x + allObjs(lp).xyz.x
     tmpObj.xyz.y = allObjs(lp).whd.h / 2 * y + allObjs(lp).xyz.y
     tmpObj.xyz.z = allObjs(lp).whd.d / 2 * z + allObjs(lp).xyz.z
' rotate based on object center orientation
' translate to origin
     tRen(0) = rotTranMul(tmpObj, allObjs(lp))
' translate again
     tmpObj.xyz.x = tRen(0).x - allObjs(lp).xyz.x
     tmpObj.xyz.y = tRen(0).w - allObjs(lp).xyz.y
     tmpObj.xyz.z = tRen(0).z - allObjs(lp).xyz.z
' save the screen xy cordinates
     pntsAry(cnt) = getXY(tmpObj.xyz)
     cnt = cnt + 1
     Next z
   Next y
  Next x

  tStr = "124356871265348715732684" ' 8 cube corners making 6 sides
  For i = 0 To 5 ' 6 sides
   idx = (lp - 1) * 6 + i
   vertAry(idx).depth = -9999
   For j = 0 To 3 ' 4 corners per side
    With pntsAry(Val(Mid(tStr, i * 4 + j + 1, 1)) - 1)
     vertAry(idx).vtx(j).x = .x
     vertAry(idx).vtx(j).y = .y
     If .z > vertAry(idx).depth Then
      vertAry(idx).depth = .z
     End If
    End With
   Next j
   vertAry(idx).clr = clrAry(idx)
  Next i
  
 Next lp
 
 Picture1.Cls
 
' simple sort to draw back shapes first
 For i = 0 To 10
  For j = i + 1 To 11
   If vertAry(i).depth < vertAry(j).depth Then
    tmpVert = vertAry(i)
    vertAry(i) = vertAry(j)
    vertAry(j) = tmpVert
   End If
  Next j
  drawAPoly vertAry(i)
 Next i
 drawAPoly vertAry(11)
 
 Form1.Picture1.Refresh

End Function

Private Sub HScroll1_Change(Index As Integer)
 Dim which As Integer
 
 which = List1.ListIndex
 Select Case Index
  Case 0
   allObjs(which).xyz.x = HScroll1(0).Value - 400
  Case 1
   allObjs(which).xyz.y = HScroll1(1).Value - 400
  Case 2
   allObjs(which).xyz.z = HScroll1(2).Value - 400
  Case 3
   allObjs(which).deg.xDeg = HScroll1(3).Value Mod 360
  Case 4
   allObjs(which).deg.yDeg = HScroll1(4).Value Mod 360
  Case 5
   allObjs(which).deg.zDeg = HScroll1(5).Value Mod 360
 End Select
 
 Label2(0) = allObjs(which).xyz.x
 Label2(1) = allObjs(which).xyz.y
 Label2(2) = allObjs(which).xyz.z
 
 Label2(3) = allObjs(which).deg.xDeg & Chr(176)
 Label2(4) = allObjs(which).deg.yDeg & Chr(176)
 Label2(5) = allObjs(which).deg.zDeg & Chr(176)
 
 drawWorld
 
End Sub

Private Sub List1_Click()
 Dim which As Integer
 
 which = List1.ListIndex
 HScroll1(0).Value = allObjs(which).xyz.x + 400
 HScroll1(1).Value = allObjs(which).xyz.y + 400
 HScroll1(2).Value = allObjs(which).xyz.z + 400
 HScroll1(3).Value = allObjs(which).deg.xDeg
 HScroll1(4).Value = allObjs(which).deg.yDeg
 HScroll1(5).Value = allObjs(which).deg.zDeg
 
End Sub

Private Sub VScroll1_Change()
 drawWorld
End Sub
