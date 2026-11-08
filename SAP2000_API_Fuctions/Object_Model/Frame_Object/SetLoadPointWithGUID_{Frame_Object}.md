# SetLoadPointWithGUID {Frame Object}

## Syntax

SapObject.SapModel.FrameObj.SetLoadPointWithGUID

## VB6 Procedure

Function SetLoadPointWithGUID(ByVal Name As String, ByVal LoadPat As String, ByVal MyType As Integer, ByVal dir As Integer, ByVal dist As Double, ByVal val As Double, ByRef GUID As String, Optional ByVal csys As String = "Global", Optional ByVal RelDist As Boolean = True, Optional ByVal Replace As Boolean = True) As Long

## Parameters

Name

The name of an existing frame object.

LoadPat

The name of a defined load pattern.

MyType

This is 1 or 2, indicating the type of point load.

1 = Force

2 = Moment

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

Dist

This is the distance from the I-End of the frame object to the load location. This may be a relative distance (0 <= Dist <= 1) or an actual distance, depending on the value of the RelDist item. [L] when RelDist is False

Val

This is the value of the point load. [F] when MyType is 1 and [FL] when MyType is 2

GUID

This is the global unique ID of a point load assigned to the frame object or if it is not the global unique id of a point load assigned to the frame object and it is not blank, the global unique ID which is assigned to the newly assigned load. If left blank, a new load is assigned to the frame object and the value of this parameter is set to the global unique ID of the newly assigned load

CSys

This is Local or the name of a defined coordinate system. It is the coordinate system in which the loads are specified.

RelDist

If this item is True, the specified Dist item is a relative distance, otherwise it is an actual distance.

Replace

If this item is True and the input GUID is not the GUID of any point load assigned to the frame object, all previous point loads, if any, assigned to the specified frame object, in the specified load pattern, are deleted before making the new assignment. If the input GUID is the GUID of a point load already assigned to the frame object, the parameters of the point load are updated with the values provided and this item is ignored.

The function returns zero if the loads are successfully assigned, otherwise it returns a nonzero value.

## Remarks

This function assigns point loads to frame objects.

The function returns zero if the loads are successfully assigned, otherwise it returns a nonzero value.

## Release Notes

Initial release in version 17.2.0.

## See Also

[GetLoadPointWithGUID](GetLoadPoint_%7BFrame_Object%7D.md)

[DeleteLoadPointWithGUID](DeleteLoadPointWithGUID_%7BFrame_Object%7D.md)
