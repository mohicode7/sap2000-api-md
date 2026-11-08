# DeleteLoadDistributedWithGUID

## Syntax

SapObject.SapModel.FrameObj.DeleteLoadDistributedWithGUID

## VB6 Procedure

Function DeleteLoadDistributedWithGUID(ByVal Name As String, ByVal GUID As String) As Long

## Parameters

Name

The name of an existing frame object.

GUID

The global unique ID of one of the distributed loads on that frame object.

## Remarks

This function deletes the distributed load assignment with the specified global unique ID for the frame objects for the specified frame object.

The function returns zero if the load assignments are successfully deleted, otherwise it returns a nonzero value.

## Release Notes

Initial release in version 17.20.

## See Also

[GetLoadDistributedWithGUID](GetLoadDistributedWithGUID_%7BFrame_Object%7D.md)

[SetLoadDistributedWithGUID](../../../SetLoadDistributed_%7BFrame_Object%7D.md)
