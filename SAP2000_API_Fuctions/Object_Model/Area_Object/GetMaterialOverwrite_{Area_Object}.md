# GetMaterialOverwrite

## Syntax

SapObject.SapModel.AreaObj.GetMaterialOverwrite

## VB6 Procedure

Function GetMaterialOverwrite(ByVal Name As String, ByRef PropName As String) As Long

## Parameters

Name

The name of a defined area object.

PropName

This is None, indicating that no material overwrite exists for the specified area object, or it is the name of an existing material property.

## Remarks

This function retrieves the material overwrite assigned to an area object, if any. The material property name is indicated as None if there is no material overwrite assignment.

The function returns zero if the material overwrite assignment is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetAreaMaterialOverwrite()

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



   'assign material overwrite

      ret = SapModel.AreaObj.SetMaterialOverwrite("3", "A992Fy50")



   'get material overwrite assignment

      ret = SapModel.AreaObj.GetMaterialOverwrite("3", PropName)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[SetMaterialOverwrite](SetMaterialOverwrite_%7BArea_Object%7D.md)
