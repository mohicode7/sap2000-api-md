# DeleteLoadDistributedWithGUID

## Syntax

SapObject.SapModel.CableObj.DeleteLoadDistributedWithGUID

## VB6 Procedure

Function DeleteLoadDistributedWithGUID(ByVal Name As String, ByVal GUID As String) As Long

## Parameters

Name

The name of an existing cable object.

GUID

The global unique ID of one of the distributed loads on that cable object.

## Remarks

This function deletes the distributed load assignment with the specified global unique ID for the specified cable object.

The function returns zero if the load assignment is successfully deleted, otherwise it returns a nonzero value.

## Release Notes

Initial release in version 17.2.0.

## See Also

[GetLoadDistributedWithGUID](../../../GetLoadDistributedWithGUID_%7BCable_Object%7D.md)

[SetLoadDistributedWithGUID](../../../SetLoadDistributedWithGUID_%7BCable_Object%7D.md)
