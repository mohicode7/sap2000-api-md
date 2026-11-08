# GetMaterialOverwrite

## Syntax

SapObject.SapModel.CableObj.GetMaterialOverwrite

## VB6 Procedure

Function GetMaterialOverwrite(ByVal Name As String, ByRef PropName As String) As Long

## Parameters

Name

The name of a defined cable object.

PropName

This is None, indicating that no material overwrite exists for the specified cable object, or it is the name of an existing material property.

## Remarks

This function retrieves the material overwrite assigned to a cable object, if any. It returns None if there is no material overwrite assignment.

The function returns zero if the material overwrite assignment is successfully retrieved, otherwise it returns a nonzero value.

## VBA Example



```
Sub GetCableMaterialOverwrite()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim PropName As String

      Dim Name As String



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



   'add cable object by coordinates

      ret = SapModel.CableObj.AddByCoord(-300, 0, 0, -100, 0, 124, Name)



   'set cable data

      ret = SapModel.CableObj.SetCableData(Name, 7, 1, 0, 0, 24)



   'assign material overwrite

      ret = SapModel.CableObj.SetMaterialOverwrite(Name, "4000Psi")



   'get material overwrite assignment

      ret = SapModel.CableObj.GetMaterialOverwrite(Name, PropName)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[SetMaterialOverwrite](SetMaterialOverwrite_%7BCable_Object%7D.md)
