# PropertyLink

## Syntax

SapObject.SapModel.SelectObj.PropertyLink

## VB6 Procedure

Function PropertyLink(ByVal Name As String, Optional ByVal DeSelect As Boolean = False) As Long

## Parameters

Name

The name of an existing link property.

DeSelect

The item is False if objects are to be selected and True if they are to be deselected.

## Remarks

This function selects or deselects all link objects to which the specified section property has been assigned.

The function returns zero if the selection is successfully completed, otherwise it returns nonzero.

## VBA Example



```
Sub SelectLinkProperty()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret as Long



   'create Sap2000 object

      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'open existing model

      ret = SapModel.File.OpenFile("C:\SapAPI\Example 6-003a.sdb")



   'select by link property

      ret = SapModel.SelectObj.PropertyLink("GAP1")



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[PropertyFrame](PropertyFrame.md)

[PropertyCable](PropertyCable.md)

[PropertyTendon](PropertyTendon.md)

[PropertyArea](PropertyArea.md)

[PropertySolid](PropertySolid.md)

[PropertyLinkFD](PropertyLinkFD.md)

[PropertyMaterial](PropertyMaterial.md)
