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
      Height          =   7155
      Left            =   180
      ScaleHeight     =   473
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   801
      TabIndex        =   0
      Top             =   300
      Width           =   12075
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
 
 Me.Show
 
 eye = 800

 viewPort.x = Picture1.ScaleWidth / 2
 viewPort.y = Picture1.ScaleHeight / 2
 
 drawWorld
 
End Sub

Function drawWorld()
 Dim xp As Single, zp As Single
 Dim gridCords As xyzPosStruct, tRen(1) As renderStruct

 Picture1.Cls
 
 For xp = -100 To 100 Step 10
  For zp = -100 To 100 Step 10
   gridCords.x = -100
   gridCords.z = zp
   tRen(0) = getXY(gridCords)
   gridCords.x = 100
   tRen(1) = getXY(gridCords)
   Picture1.Line (tRen(0).x, tRen(0).y)-(tRen(1).x, tRen(1).y), vbBlack
   
   gridCords.x = xp
   gridCords.z = -100
   tRen(0) = getXY(gridCords)
   gridCords.z = 100
   tRen(1) = getXY(gridCords)
   Picture1.Line (tRen(0).x, tRen(0).y)-(tRen(1).x, tRen(1).y), vbBlack
  Next zp
 Next xp

End Function

Private Sub HScroll1_Change(Index As Integer)
 
 Select Case Index
  Case 0
   worldSpace.xyz.x = HScroll1(0).Value - 400
  Case 1
   worldSpace.xyz.y = HScroll1(1).Value - 400
  Case 2
   worldSpace.xyz.z = HScroll1(2).Value - 400
  Case 3
   worldSpace.deg.xDeg = HScroll1(3).Value Mod 360
  Case 4
   worldSpace.deg.yDeg = HScroll1(4).Value Mod 360
  Case 5
   worldSpace.deg.zDeg = HScroll1(5).Value Mod 360
 End Select
 
 Label2(0) = worldSpace.xyz.x
 Label2(1) = worldSpace.xyz.y
 Label2(2) = worldSpace.xyz.z
 
 Label2(3) = worldSpace.deg.xDeg & Chr(176)
 Label2(4) = worldSpace.deg.yDeg & Chr(176)
 Label2(5) = worldSpace.deg.zDeg & Chr(176)
 
 drawWorld
 
End Sub
