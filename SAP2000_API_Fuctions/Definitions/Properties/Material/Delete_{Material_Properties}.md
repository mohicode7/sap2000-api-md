# Delete

## Syntax

SapObject.SapModel.PropMaterial.Delete

## VB6 Procedure

Function Delete(ByVal Name As String) As Long

## Parameters

Name

The name of an existing material property.

## Remarks

The function deletes a specified material property.

The function returns zero if the material property is successfully deleted; otherwise it returns a nonzero value. It returns an error if the specified material property can not be deleted, for example, if it is being used in a section property.

## VBA Example



```
Sub DeleteMaterial()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long



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



   'delete material

      ret = SapModel.PropMaterial.Delete("4000Psi")



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also
