# GetInsertionPoint (Note:  Newer function available)

## Syntax

SapObject.SapModel.FrameObj.GetInsertionPoint

## VB6 Procedure

Function GetInsertionPoint(ByVal Name As String, ByRef
CardinalPoint As Long, ByRef Mirror2 As Boolean, ByRef StiffTransform
As Boolean, ByRef Offset1() As Double, ByRef Offset2() As Double, ByRef
CSys As String) As Long

## Parameters

Name

The name of an existing frame object.

CardinalPoint

This is a numeric value from 1 to 11 that specifies
the cardinal point for the frame object. The cardinal point specifies
the relative position of the frame section on the line representing the
frame object.

1 = bottom left

2 = bottom center

3 = bottom right

4 = middle left

5 = middle center

6 = middle right

7 = top left

8 = top center

9 = top right

10 = centroid

11
= shear center

Mirror2

If this item is True, the frame object section is assumed
to be mirrored (flipped) about its local 2-axis.

StiffTransform

If this item is True, the frame object stiffness is
transformed for cardinal point and joint offsets from the frame section
centroid.

Offset1

This is an array of three joint offset distances, in
the coordinate directions specified by CSys, at the I-End of the frame
object. [L]

Offset1(0) = Offset in
the 1-axis or X-axis direction

Offset1(1) = Offset in
the 2-axis or Y-axis direction

Offset1(2)
= Offset in the 3-axis or Z-axis direction

Offset2

This is an array of three joint offset distances, in
the coordinate directions specified by CSys, at the J-End of the frame
object. [L]

Offset2(0) = Offset in
the 1-axis or X-axis direction

Offset2(1) = Offset in
the 2-axis or Y-axis direction

Offset2(2) = Offset in the
3-axis or Z-axis direction

CSys

This is either Local or the name of a defined coordinate
system. It is the coordinate system in which the Offset1 and Offset2 items
are specified.

## Remarks

This function retrieves frame object insertion point
assignments. The assignments include the cardinal point and end joint
offsets.

The function returns zero if the insertion point data
is successfully retrieved, otherwise it returns a nonzero value.

## VBA Example



```
Sub GetFrameInsertionPoint()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim i As Long

      Dim CardinalPoint As Long

      Dim Mirror2 As Boolean
      Dim StiffTransform 
 As Boolean

      Dim Offset1() As Double

      Dim Offset2() As Double

      Dim CSys As String



   'create Sap2000 object

      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'create model from template

      ret = SapModel.File.New2DFrame(PortalFrame, 
 3, 124, 3, 200)



   'assign frame insertion point

      ReDim Offset1(2)

      ReDim Offset2(2)

      For i=0 To 2

         Offset1(i)=10 + i

         Offset2(i)=20 + i

      Next i

      ret = SapModel.FrameObj.SetInsertionPoint("15", 
 7, False, True, Offset1, Offset2)



   'get frame insertion point

      ReDim Offset1(2)

      ReDim Offset2(2)

      ret = SapModel.FrameObj.GetInsertionPoint("15", 
 CardinalPoint, Mirror2, StiffTransform, Offset1, Offset2, CSys)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

This function is obsolete and has been superseded by
[GetInsertionPoint\_1](../Object_Model/Frame_Object/GetInsertionPoint_1%7BFrame_Object%7D.md)
as of version 24.1. This function is maintained for backwards compatibility.

## See Also

[SetInsertionPoint](SetInsertionPoint%7BFrame_Object%7D.md)
