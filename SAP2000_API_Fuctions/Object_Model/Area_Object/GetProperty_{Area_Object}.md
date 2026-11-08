# GetProperty

## Syntax

SapObject.SapModel.AreaObj.GetProperty

## VB6 Procedure

Function GetProperty(ByVal Name As String, ByRef PropName As String) As Long

## Parameters

Name

The name of a defined area object.

PropName

The name of the area property assigned to the area object. This item is None if no area property is assigned to the area object.

## Remarks

This function retrieves the area property assigned to an area object.

The function returns zero if the property is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetAreaObjectProp()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim PropName As String



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



   'get area property

      ret = SapModel.AreaObj.GetProperty("1", PropName)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[SetProperty](SetProperty_%7BArea_Object%7D.md)
