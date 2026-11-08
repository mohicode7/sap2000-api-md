# GetLoadForceWithGUID {Point Object}

## Syntax

SapObject.SapModel.PointObj.GetLoadForceWithGUID

## VB6 Procedure

Function GetLoadForceWithGUID(ByVal Name As String, ByRef NumberItems As Integer, ByRef PointName() As String, ByRef LoadPat() As String, ByRef LcStep() As Integer, ByRef CSys() As String, ByRef F1() As Double, ByRef F2() As Double, ByRef F3() As Double, ByRef M1() As Double, ByRef M2() As Double, ByRef M3() As Double, ByRef GUID() As String, Optional ByVal ItemType As eItemType = eItemType.Objects) As Long

## Parameters

Name

The name of an existing point object or group, depending on the value of the ItemType item.

NumberItems

This is the total number of joint force load assignments returned.

PointName

This is an array that includes the name of the point object to which the specified load assignment applies.

LoadPat

This is an array that includes the name of the load pattern for the load.

LCStep

This is an array that includes the load pattern step for the load. In most cases, this item does not apply and will be returned as 0.

CSys

This is an array that includes the name of the coordinate system for the load. This is Local or the name of a defined coordinate system.

F1

This is an array that includes the assigned translational force in the local 1-axis or coordinate system X-axis direction, depending on the specified CSys. [F]

F2

This is an array that includes the assigned translational force in the local 2-axis or coordinate system Y-axis direction, depending on the specified CSys. [F]

F3

This is an array that includes the assigned translational force in the local 3-axis or coordinate system Z-axis direction, depending on the specified CSys. [F]

M1

This is an array that includes the assigned moment about the local 1-axis or coordinate system X-axis, depending on the specified CSys. [FL]

M2

This is an array that includes the assigned moment about the local 2-axis or coordinate system Y-axis, depending on the specified CSys. [FL]

M3

This is an array that includes the assigned moment about the local 3-axis or coordinate system Z-axis, depending on the specified CSys. [FL]

GUID

This is an array that includes the global unique ID of the distributed load.

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the load assignments are retrieved for the point object specified by the Name item.

If this item is Group, the assignments are retrieved for all point objects in the group specified by the Name item.

If this item is SelectedObjects, the load assignments are retrieved for all selected point objects, and the Name item is ignored.

## Remarks

This function retrieves the joint force load assignments to point objects.

The function returns zero if the load assignments are successfully retrieved, otherwise it returns a nonzero value.

## Release Notes

Initial release in version 17.2.0.

## See Also

[SetLoadForceWithGUID](SetLoadForceWithGUID_%7BPoint_Object%7D.md)

[DeleteLoadForceWithGUID](DeleteLoadForceWithGUID_%7BPoint_Object%7D.md)
