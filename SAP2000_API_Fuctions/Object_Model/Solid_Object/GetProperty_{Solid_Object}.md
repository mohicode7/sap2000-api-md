# GetProperty

## Syntax

SapObject.SapModel.SolidObj.GetProperty

## VB6 Procedure

Function GetProperty(ByVal Name As String, ByRef PropName As String) As Long

## Parameters

Name

The name of a defined solid object.

PropName

The name of the solid property assigned to the solid object.

## Remarks

This function retrieves the solid property assigned to a solid object.

The function returns zero if the property is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetSolidObjectProp()

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

      ret = SapModel.File.NewSolidBlock(300, 400, 200, , , 2, 2, 2)



   'get solid property

      ret = SapModel.SolidObj.GetProperty("1", PropName)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[SetProperty](SetProperty_%7BSolid_Object%7D.md)
