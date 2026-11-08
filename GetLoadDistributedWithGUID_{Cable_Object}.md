# GetLoadDistributedWithGUID

## Syntax

SapObject.SapModel.CableObj.GetLoadDistributedWithGUID

## VB6 Procedure

Function GetLoadDistributedWithGUID(ByVal Name As String, ByRef NumberItems As Long, ByRef CableName() As String, ByRef LoadPat() As String, ByRef MyType() As Long, ByRef CSys() As String, ByRef Dir() As Long, ByRef Value() As Double, ByRefGUID() As String, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing cable object or group, depending on the value of the ItemType item.

NumberItems

The total number of distributed loads retrieved for the specified cable objects.

CableName

This is an array that includes the name of the cable object associated with each distributed load.

LoadPat

This is an array that includes the name of the coordinate system in which the distributed loads are specified.

MyType

This is an array that includes 1 or 2, indicating the type of distributed load.

1 = Force

2 = Moment

CSys

This is an array that includes the name of the coordinate system in which each distributed load is defined. It may be Local or the name of a defined coordinate system.

Dir

This is 1, 2, 3, 4, 5, 6 or 10, indicating the direction of the load.

1 = Local 1 axis (only applies when CSys is Local)

2 = Local 2 axis (only applies when CSys is Local)

3 = Local 3 axis (only applies when CSys is Local)

4 = X direction (does not apply when CSys is Local)

5 = Y direction (does not apply when CSys is Local)

6 = Z direction (does not apply when CSys is Local)

10 = Gravity direction (only applies when CSys is Global)

The positive gravity direction (see Dir = 10) is in the negative Global Z direction.

Value

This is the load value of the distributed load. The distributed load is applied over the full length of the cable. [F/L] when MyType is 1 and [FL/L] when MyType is 2

GUID

This is an array that includes the global unique ID of each load.

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the assignments are retrieved for the cable object specified by the Name item.

If this item is Group, the assignments are retrieved for all cable objects in the group specified by the Name item.

If this item is SelectedObjects, assignments are retrieved for all selected cable objects, and the Name item is ignored.

## Remarks

This function retrieves the distributed load assignments to cable objects. The loads are uniformly distributed over the full length of cable objects.

The function returns zero if the load assignments are successfully retrieved, otherwise it returns a nonzero value.

## Release Notes

Initial release in version 17.2.0.

## See Also

[SetLoadDistributedWithGUID](SetLoadDistributedWithGUID_%7BCable_Object%7D.md)

[DeleteLoadDistributedWithGUID](SAP2000_API_Fuctions/Object_Model/Cable_Object/DeleteLoadDistributedWithGUID_%7BCable_Object%7D.md)
