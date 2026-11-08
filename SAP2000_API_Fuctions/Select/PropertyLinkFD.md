# PropertyLinkFD

## Syntax

SapObject.SapModel.SelectObj.PropertyLinkFD

## VB6 Procedure

Function PropertyLinkFD(ByVal Name As String, Optional ByVal DeSelect As Boolean = False) As Long

## Parameters

Name

The name of an existing frequency dependent link property.

DeSelect

The item is False if objects are to be selected and True if they are to be deselected.

## Remarks

This function selects or deselects all link objects to which the specified frequency dependent link property has been assigned.

The function returns zero if the selection is successfully completed, otherwise it returns nonzero.

## VBA Example

Similar to PropertyFrame

   'select by frequency dependent link property  
      ret = SapModel.SelectObj.PropertyLinkFD("FD1")

## Release Notes

Initial release in version 11.00.

## See Also

[PropertyFrame](PropertyFrame.md)

[PropertyCable](PropertyCable.md)

[PropertyTendon](PropertyTendon.md)

[PropertyArea](PropertyArea.md)

[PropertySolid](PropertySolid.md)

[PropertyLink](PropertyLink.md)

[PropertyMaterial](PropertyMaterial.md)
