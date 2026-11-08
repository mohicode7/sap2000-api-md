# SetLoadForceWithGUID {Point Object}

## Syntax

SapObject.SapModel.PointObj.SetLoadForceWithGUID

## VB6 Procedure

Function SetLoadForceWithGUID(ByVal Name As String, ByVal LoadPat As String, ByRef Value() As Double, ByRef GUID As String, Optional ByVal Replace As Boolean = False, Optional ByVal CSys As String = "Global") As Long

## Parameters

Name

The name of an existing point object.

LoadPat

The name of the load pattern for the point load.

Value

This is an array of six point load values.

Value(0) = F1 [F]

Value(1) = F2 [F]

Value(2) = F3 [F]

Value(3) = M1 [FL]

Value(4) = M2 [FL]

Value(5) = M3 [FL]

GUID

This is the global unique ID of a load force assigned to the point object or if it is not the global unique id of a load force assigned to the point object and it is not blank, the global unique ID which is assigned to the newly assigned load. If left blank, a new load is assigned to the point object and the value of this parameter is set to the global unique ID of the newly assigned load.

Replace

If this item is True and the input GUID is not the GUID of any load force assigned to the point object, all previous loads force, if any, assigned to the specified point object, in the specified load pattern, are deleted before making the new assignment. If the input GUID is the GUID of a load force already assigned to the point object, the parameters of the distributed load are updated with the values provided and this item is ignored.

CSys

The name of the coordinate system for the considered point load. This is Local or the name of a defined coordinate system.

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the load assignment is made to the point object specified by the Name item.

If this item is Group, the load assignment is made to all point objects in the group specified by the Name item.

If this item is SelectedObjects, the load assignment is made to all selected point objects and the Name item is ignored.

## Remarks

This function makes point load assignments to point objects.

The function returns zero if the load assignments are successfully made, otherwise it returns a nonzero value.

## Release Notes

Initial release in version 17.2.0.

## See Also

[GetLoadForceWithGUID](GetLoadForceWithGUID_%7BPoint_Object%7D.md)

[DeleteLoadForceWithGUID](DeleteLoadForceWithGUID_%7BPoint_Object%7D.md)
