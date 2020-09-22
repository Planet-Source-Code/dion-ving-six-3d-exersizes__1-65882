Attribute VB_Name = "Module1"
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public exitFlag As Boolean

Public Const radian As Double = 0.017453292519943

Public Type xyzPosStruct
 x As Single
 y As Single
 z As Single
End Type

Public Type orientStruct
 xDeg As Single
 yDeg As Single
 zDeg As Single
End Type

Public Type dimsStruct
 w As Single
 h As Single
 d As Single
End Type

Public Type objStruct
 xyz As xyzPosStruct
 deg As orientStruct
 whd As dimsStruct
End Type

Public Type renderStruct
 x As Single
 y As Single
 z As Single
 w As Single
End Type

Public Type xyPair
 x As Single
 y As Single
End Type

Public worldSpace As objStruct
Public viewPort As xyPair
Public eye As Single

Function getXY(xyzCords As xyzPosStruct) As renderStruct
 Dim tmpSrcObj As objStruct, rtmObj As renderStruct
 
 tmpSrcObj.xyz = xyzCords
 rtmObj = rotTranMul(tmpSrcObj, worldSpace)
 getXY.x = rtmObj.x * eye / (rtmObj.z + eye) + viewPort.x
 getXY.y = rtmObj.w * eye / (rtmObj.z + eye) + (eye / (rtmObj.z + eye)) + viewPort.y
 getXY.z = rtmObj.z
 getXY.w = eye / ((rtmObj.z + eye) + 0.001)

End Function

Function rotTranMul(srcObj As objStruct, refOBJ As objStruct) As renderStruct
 Dim trig(7, 1) As Single, i As Integer, tSing As Single
 Dim tmpObj As renderStruct
 
 
' rotate x,y,z
 tSing = refOBJ.deg.xDeg
 For i = 0 To 2
  trig(i, 0) = Sin(tSing * radian)
  trig(i, 1) = Cos(tSing * radian)
  trig(i + 4, 0) = Sin(-tSing * radian)
  trig(i + 4, 1) = Cos(-tSing * radian)
  tSing = refOBJ.deg.yDeg
  If i = 1 Then
   tSing = refOBJ.deg.zDeg
  End If
 Next i
       
' translate x,y,z
 tmpObj.x = srcObj.xyz.x - refOBJ.xyz.x
 tmpObj.y = srcObj.xyz.y - refOBJ.xyz.y
 tmpObj.z = srcObj.xyz.z - refOBJ.xyz.z

' multpliy
 tmpObj.w = tmpObj.x * trig(2, 1) + tmpObj.y * trig(2, 0)
 tmpObj.y = tmpObj.y * trig(2, 1) + tmpObj.x * trig(6, 0)
 tmpObj.x = tmpObj.w * trig(1, 1) + (-tmpObj.z * trig(1, 0))
 tmpObj.z = -((-tmpObj.z * trig(1, 1)) + (tmpObj.w * trig(5, 0)))
 tmpObj.w = tmpObj.y * trig(0, 1) + (-tmpObj.z * trig(0, 0))
 tmpObj.z = -((-tmpObj.z * trig(0, 1)) + (tmpObj.y * trig(4, 0)))
 
 rotTranMul = tmpObj
 
End Function

