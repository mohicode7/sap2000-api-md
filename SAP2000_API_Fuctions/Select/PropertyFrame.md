# PropertyFrame

## Syntax

SapObject.SapModel.SelectObj.PropertyFrame

## VB6 Procedure

Function PropertyFrame(ByVal Name As String, Optional ByVal DeSelect As Boolean = False) As Long

## Parameters

Name

The name of an existing frame section property.

DeSelect

The item is False if objects are to be selected and True if they are to be deselected.

## Remarks

This function selects or deselects all line objects to which the specified section has been assigned.

The function returns zero if the selection is successfully completed, otherwise it returns nonzero.

## VBA Example



```
Sub SelectFrameProperty()

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



   'create model from template

      ret = SapModel.File.New2DFrame(PortalFrame, 3, 124, 3, 200)



   'select by frame property

      ret = SapModel.SelectObj.PropertyFrame("FSEC1")



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[PropertyCable](PropertyCable.md)

[PropertyTendon](PropertyTendon.md)

[PropertyArea](PropertyArea.md)

[PropertySolid](PropertySolid.md)

[PropertyLink](PropertyLink.md)

[PropertyLinkFD](PropertyLinkFD.md)

[PropertyMaterial](PropertyMaterial.md)
