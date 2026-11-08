# PropertyTendon

## Syntax

SapObject.SapModel.SelectObj.PropertyTendon

## VB6 Procedure

Function PropertyTendon(ByVal Name As String, Optional ByVal DeSelect As Boolean = False) As Long

## Parameters

Name

The name of an existing tendon section property.

DeSelect

The item is False if objects are to be selected and True if they are to be deselected.

## Remarks

This function selects or deselects all tendon objects to which the specified section has been assigned.

The function returns zero if the selection is successfully completed, otherwise it returns nonzero.

## VBA Example



```
Sub SelectTendonProperty()

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

      ret = SapModel.File.OpenFile("C:\SapAPI\Example 1-009a.sdb")



   'select by tendon property

      ret = SapModel.SelectObj.PropertyTendon("TEN1")



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

[PropertyArea](PropertyArea.md)

[PropertySolid](PropertySolid.md)

[PropertyLink](PropertyLink.md)

[PropertyLinkFD](PropertyLinkFD.md)

[PropertyMaterial](PropertyMaterial.md)
