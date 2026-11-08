# SetLoadDistributedWithGUID

## Syntax

SapObject.SapModel.CableObj.SetLoadDistributedWithGUID

## VB6 Procedure

Function SetLoadDistributedWithGUID (ByVal Name As String, ByVal LoadPat As String, ByVal MyType As Long, ByVal Dir As Long, ByVal Value As Double, ByRef GUID As String, Optional ByVal CSys As String = "Global", Optional ByVal Replace As Boolean = True) As Long

## Parameters

Name

The name of an existing cable object or group, depending on the value of the ItemType item.

LoadPat

The name of a defined load pattern.

MyType

This is 1 or 2, indicating the type of distributed load.

1 = Force per unit length

2 = Moment per unit length

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

This is the global unique ID of a distributed load assigned to the cable object or if it is not the global unique ID of a distributed load assigned to the cable object and it is not blank, the global unique ID which is assigned to the newly assigned load. If left blank, a new load assigned to the cable object and the value of this parameter is set to the global unique ID of the newly assigned load.

CSys

This is Local or the name of a defined coordinate system. It is the coordinate system in which the loads are specified.

Replace

If this item is True and the input GUID is not the GUID of any distributed load assigned to the cable object, all previous distributed loads, if any, assigned to the specified cable object, in the specified load pattern, are deleted before making the new assignment. If the input GUID is the GUID of a distributed load already assigned to the frame object, the parameters of the distributed load are updated with the values provided and this item is ignored.

If this item is True, all previous loads, if any, assigned to the specified cable object(s), in the specified load pattern, are deleted before making the new assignment.

## Remarks

If the cable object is already assigned a distributed load with a global unique ID matching the specified global unique ID, this function modifies that distributed load. Otherwise, this function assigns a new distributed load over the full length of the cable object and sets its global unique ID to the specified global unique ID.

This function assigns uniform distributed loads over the full length of cable objects.

The function returns zero if the loads are successfully assigned, otherwise it returns a nonzero value.

## Release Notes

Initial release in version 17.2.0.

## See Also

[GetLoadDistributedWithGUID](GetLoadDistributedWithGUID_%7BCable_Object%7D.md)

[DeleteLoadDistributedWithGUID](SAP2000_API_Fuctions/Object_Model/Cable_Object/DeleteLoadDistributedWithGUID_%7BCable_Object%7D.md)
