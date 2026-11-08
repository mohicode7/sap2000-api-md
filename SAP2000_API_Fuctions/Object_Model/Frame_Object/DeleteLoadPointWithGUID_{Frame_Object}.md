# DeleteLoadPointWithGUID {Frame Object}

## Syntax

SapObject.SapModel.FrameObj.DeleteLoadPointWithGUID

## VB6 Procedure

Function DeleteLoadPointWithGUID(ByVal Name As String, ByVal GUID As String) As Long

## Parameters

Name

The name of an existing frame object.

GUID

The global unique ID for one of the point loads on that frame object.

## Remarks

This function deletes the point load assignment with the specified global unique ID for the specified frame object.

The function returns zero if the load assignment is successfully deleted, otherwise it returns a nonzero value.

## Release Notes

Initial release in version 17.2.0.

## See Also

[GetLoadPointWithGUID](../../../GetLoadPointWithGUID_%7BFrame_Object%7D.md)

[SetLoadPointWithGUID](SetLoadPointWithGUID_%7BFrame_Object%7D.md)
