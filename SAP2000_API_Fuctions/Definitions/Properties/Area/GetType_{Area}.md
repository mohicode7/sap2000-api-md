# GetTypeOAPI

## Syntax

SapObject.SapModel.PropArea.GetTypeOAPI

## VB6 Procedure

Function GetTypeOAPI(ByVal Name As String, ByRef PropType As Long) As Long

## Parameters

Name

The name of an existing area property.

PropType

This is 1, 2 or 3, indicating the type of area property.

1 = Shell

2 = Plane

3 = Asolid

## Remarks

This function retrieves the property type for the specified area property.

The function returns zero if the type is successfully retrieved; otherwise it returns nonzero.

## VBA Example



```
Sub GetAreaPropType()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim PropType As Long



   'create Sap2000 object

      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'create model from template

      ret = SapModel.File.NewWall(2, 48, 2, 48)



   'get area property type

      ret = SapModel.PropArea.GetTypeOAPI("ASEC1", PropType)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

Changed function name to GetTypeOAPI in v17.0.0.

## See Also
