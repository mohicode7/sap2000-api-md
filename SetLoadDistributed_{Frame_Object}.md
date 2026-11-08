# SetLoadDistributedWithGUID {Frame Object}

## Syntax

SapObject.SapModel.FrameObj.SetLoadDistributedWithGUID

## VB6 Procedure

Function SetLoadDistributedWithGUID(ByVal Name As String, ByVal LoadPat As String, ByVal MyType As Long, ByVal Dir As Long, ByVal Dist1 As Double, ByVal Dist2 As Double, ByVal Val1 As Double, ByVal Val2 As Double, ByRef GUID As String, Optional ByVal CSys As String = "Global", Optional ByVal RelDist As Boolean = True, Optional ByVal Replace As Boolean = True) As Long

## Parameters

Name

The name of an existing frame object.

LoadPat

The name of a defined load pattern.

MyType

This is 1 or 2, indicating the type of distributed load.

1 = Force per unit length

2 = Moment per unit length

Dir

This is an integer between 1 and 11, indicating the direction of the load.

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

Dist1

This is the distance from the I-End of the frame object to the start of the distributed load. This may be a relative distance (0 <= Dist1 <= 1) or an actual distance, depending on the value of the RelDist item. [L] when RelDist is False

Dist2

This is the distance from the I-End of the frame object to the end of the distributed load. This may be a relative distance (0 <= Dist2 <= 1) or an actual distance, depending on the value of the RelDist item. [L] when RelDist is False

Val1

This is the load value at the start of the distributed load. [F/L] when MyType is 1 and [FL/L] when MyType is 2

Val2

This is the load value at the end of the distributed load. [F/L] when MyType is 1 and [FL/L] when MyType is 2

GUID

This is the global unique ID of a distributed load assigned to the frame object or if it is not the global unique id of a distributed load assigned to the frame object and it is not blank, the global unique ID which is assigned to the newly assigned load. If left blank, a new load is assigned to the frame object and the value of this parameter is set to the global unique ID of the newly assigned load.

CSys

This is Local or the name of a defined coordinate system. It is the coordinate system in which the loads are specified.

RelDist

If this item is True, the specified Dist item is a relative distance, otherwise it is an actual distance.

Replace

If this item is True and the input GUID is not the GUID of any distributed load assigned to the frame object, all previous distributed loads, if any, assigned to the specified frame object, in the specified load pattern, are deleted before making the new assignment. If the input GUID is the GUID of a distributed load already assigned to the frame object, the parameters of the distributed load are updated with the values provided and this item is ignored

## Remarks

If the frame object is already assigned a distributed load with a global unique ID matching the specified global unique ID, this function modifies that distributed load. Otherwise, this function assigns a new distributed load to the frame object and sets its global unique ID to the specified global unique ID.

The function returns zero if the loads are successfully assigned, otherwise it returns a nonzero value.

## Release Notes

Initial release in version 17.2.0.

## See Also

[GetLoadDistributedWithGUID](SAP2000_API_Fuctions/Object_Model/Frame_Object/GetLoadDistributed_%7BFrame_Object%7D.md)

[DeleteLoadDistributedWithGUID](SAP2000_API_Fuctions/Object_Model/Frame_Object/DeleteLoadDistributed_%7BFrame_Object%7D.md)
