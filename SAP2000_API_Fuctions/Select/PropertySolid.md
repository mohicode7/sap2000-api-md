# PropertySolid

## Syntax

SapObject.SapModel.SelectObj.PropertySolid

## VB6 Procedure

Function PropertySolid(ByVal Name As String, Optional ByVal DeSelect As Boolean = False) As Long

## Parameters

Name

The name of an existing solid property.

DeSelect

The item is False if objects are to be selected and True if they are to be deselected.

## Remarks

This function selects or deselects all solid objects to which the specified property has been assigned.

The function returns zero if the selection is successfully completed, otherwise it returns nonzero.

## VBA Example

Similar to PropertyFrame

   'select by solid property  
      ret = SapModel.SelectObj.PropertySolid("SOLID1")

## Release Notes

Initial release in version 11.00.

## See Also

[PropertyFrame](PropertyFrame.md)

[PropertyCable](PropertyCable.md)

[PropertyTendon](PropertyTendon.md)

[PropertyArea](PropertyArea.md)

[PropertyLink](PropertyLink.md)

[PropertyLinkFD](PropertyLinkFD.md)

[PropertyMaterial](PropertyMaterial.md)
