# DeleteLoadForceWithGUID {Point Object}

## Syntax

SapObject.SapModel.PointObj.DeleteLoadForceWithGUID

## VB6 Procedure

Function DeleteLoadForceWithGUID(ByVal Name As String, ByVal GUID As String) As Long

## Parameters

Name

The name of an existing point object.

GUID

The global unique ID of one of the point loads on that point object.

## Remarks

This function deletes the point load assignment with the specified global unique ID for the specified point object.

The function returns zero if the load assignment is successfully deleted, otherwise it returns a nonzero value.

## Release Notes

Initial release in version 17.2.0.

## See Also

[GetLoadForceWithGUID](GetLoadForceWithGUID_%7BPoint_Object%7D.md)

[SetLoadForceWithGUID](SetLoadForceWithGUID_%7BPoint_Object%7D.md)
