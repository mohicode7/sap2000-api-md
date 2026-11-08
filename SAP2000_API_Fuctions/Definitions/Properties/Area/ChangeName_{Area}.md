# ChangeName

## Syntax

SapObject.SapModel.PropArea.ChangeName

## VB6 Procedure

Function ChangeName(ByVal Name As String, ByVal NewName As String) As Long

## Parameters

Name

The existing name of a defined area property.

NewName

The new name for the area property.

## Remarks

This function changes the name of an existing area property.

The function returns zero if the new name is successfully applied; otherwise it returns a nonzero value.

## VBA Example



```
Sub ChangeAreaPropName()

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

      ret = SapModel.File.NewWall(2, 48, 2, 48)



   'change name of area property

      ret = SapModel.PropArea.ChangeName("ASEC1", "MyArea")



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also
