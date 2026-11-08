# SetInsertionPoint\_1

## Syntax

SapObject.SapModel.FrameObj.SetInsertionPoint\_1

## VB6 Procedure

Function SetInsertionPoint\_1(ByVal Name As String, ByVal
CardinalPoint As Long, ByVal Mirror2, ByVal Mirror3,  As Boolean,
ByVal StiffTransform As Boolean, ByRef Offset1() As Double, ByRef Offset2()
As Double, Optional ByVal CSys As String = "Local", Optional
ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing frame object or group, depending
on the value of the ItemType item.

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

11 = shear center

Mirror2

If this item is True, the frame object section is assumed
to be mirrored (flipped) about its local 2-axis.

Mirror3

If this item is True, the frame object section is assumed
to be mirrored (flipped) about its local 3-axis.

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

Offset1(2) = Offset in the
3-axis or Z-axis direction

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

This is Local or the name of a defined coordinate system.
It is the coordinate system in which the Offset1 and Offset2 items are
specified.

ItemType

This is one of the following items in the eItemType
enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the assignment is made to the
frame object specified by the Name item.

If this item is Group, the assignment is made to all
frame objects in the group specified by the Name item.

If this item is SelectedObjects, assignment is made
to all selected frame objects, and the Name item is ignored.

## Remarks

This function assigns frame object insertion point data.
The assignments include the cardinal point and end joint offsets.

The function returns zero if the insertion point data
is successfully assigned, otherwise it returns a nonzero value.

## VBA Example



```
Sub AssignFrameInsertionPoint()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim i As Long

      Dim Offset1() As Double

      Dim Offset2() As Double



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

      ret = SapModel.FrameObj.SetInsertionPoint_1("15", 
 7, False, False, True, Offset1, Offset2)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 24.1.

This function supersedes [SetInsertionPoint](../../Obsolete_Functions/SetInsertionPoint%7BFrame_Object%7D.md),
adding the Mirror3 parameter.

## See Also

[GetInsertionPoint\_1](GetInsertionPoint_1%7BFrame_Object%7D.md)
