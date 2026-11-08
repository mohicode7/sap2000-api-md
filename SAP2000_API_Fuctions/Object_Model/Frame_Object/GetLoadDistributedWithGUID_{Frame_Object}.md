# GetLoadDistributedWithGUID {Frame Object}

## Syntax

SapObject.SapModel.FrameObj.GetLoadDistributedWthGUID

## VB6 Procedure

Function GetLoadDistributed(ByVal Name As String, ByRef NumberItems As Long, ByRef FrameName() As String, ByRef LoadPat() As String, ByRef MyType() As Long, ByRef CSys() As String, ByRef Dir() As Long, ByRef RD1() As Double, ByRef RD2() As Double, ByRef Dist1() As Double, ByRef Dist2() As Double, ByRef Val1() As Double, ByRef Val2() As Double, ByRef GUID() As String, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing frame object or group, depending on the value of the ItemType item.

NumberItems

The total number of distributed loads retrieved for the specified frame objects.

FrameName

This is an array that includes the name of the frame object associated with each distributed load.

LoadPat

This is an array that includes the name of the coordinate system in which the distributed loads are specified.

MyType

This is an array that includes 1 or 2, indicating the type of distributed load.

1 = Force

2 = Moment

CSys

This is an array that includes the name of the coordinate system in which each distributed load is defined. It may be Local or the name of a defined coordinate system.

Dir

This is an array that includes an integer between 1 and 11, indicating the direction of the load.

1 = Local 1 axis (only applies when CSys is Local)

2 = Local 2 axis (only applies when CSys is Local)

3 = Local 3 axis (only applies when CSys is Local)

4 = X direction (does not apply when CSys is Local)

5 = Y direction (does not apply when CSys is Local)

6 = Z direction (does not apply when CSys is Local)

7 = Projected X direction (does not apply when CSys is Local)

8 = Projected Y direction (does not apply when CSys is Local)

9 = Projected Z direction (does not apply when CSys is Local)

10 = Gravity direction (only applies when CSys is Global)

11 = Projected Gravity direction (only applies when CSys is Global)

The positive gravity direction (see Dir = 10 and 11) is in the negative Global Z direction.

RD1

This is an array that includes the relative distance from the I-End of the frame object to the start of the distributed load.

RD2

This is an array that includes the relative distance from the I-End of the frame object to the end of the distributed load.

Dist1

This is an array that includes the actual distance from the I-End of the frame object to the start of the distributed load. [L]

Dist2

This is an array that includes the actual distance from the I-End of the frame object to the end of the distributed load. [L]

Val1

This is an array that includes the load value at the start of the distributed load. [F/L] when MyType is 1 and [FL/L] when MyType is 2

Val2

This is an array that includes the load value at the end of the distributed load. [F/L] when MyType is 1 and [FL/L] when MyType is 2

GUID

This is an array that includes the global unique ID of the distributed load.

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the assignments are retrieved for the frame object specified by the Name item.

If this item is Group, the assignments are retrieved for all frame objects in the group specified by the Name item.

If this item is SelectedObjects, assignments are retrieved for all selected frame objects, and the Name item is ignored.

## Remarks

This function is identical to GetLoadDistributed but it includes an extra return parameter, which includes the global unique IDs of the retrieved distributed loads.

The function returns zero if the load assignments are successfully retrieved, otherwise it returns a nonzero value.

## Release Notes

Initial release in version 17.2.0.

## See Also

[SetLoadDistributedWithGUID](../../../SetLoadDistributed_%7BFrame_Object%7D.md)

[DeleteLoadDistributedWithGUID](DeleteLoadDistributedWithGUID_%7BFrame_Object%7D.md)
