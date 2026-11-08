# PlaneYZ

## Syntax

SapObject.SapModel.SelectObj.PlaneYZ

## VB6 Procedure

Function PlaneYZ(ByVal Name As String, Optional ByVal DeSelect As Boolean = False) As Long

## Parameters

Name

The name of a point object.

DeSelect

The item is False if objects are to be selected and True if they are to be deselected.

## Remarks

This function selects or deselects all objects that are in the same YZ plane (in the present coordinate system) as the specified point object.

The function returns zero if the selection is successfully completed, otherwise it returns nonzero.

## VBA Example



```
Sub SelectPlaneYZ()

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



   'select in YZ plane

      ret = SapModel.SelectObj.PlaneYZ("3")



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[PlaneXY](PlaneXY.md)

[PlaneXZ](PlaneXZ.md)

[GetPresentCoordSystem](../General_Functions/GetPresentCoordSystem.md)

[SetPresentCoordSystem](../General_Functions/SetPresentCoordSystem.md)
