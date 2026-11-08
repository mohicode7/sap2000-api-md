# ChangeName

## Syntax

SapObject.SapModel.SolidObj.ChangeName

## VB6 Procedure

Function ChangeName(ByVal Name As String, ByVal NewName As String) As Long

## Parameters

Name

The existing name of a defined solid object.

NewName

The new name for the solid object.

## Remarks

This function applies a new name to a solid object.

The function returns zero if the new name is successfully applied; otherwise it returns a nonzero value.

## VBA Example



```
Sub ChangeSolidObjName()

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

      ret = SapModel.File.NewSolidBlock(300, 400, 200, , , 2, 2, 2)



   'change name

      ret = SapModel.SolidObj.ChangeName("1", "MySolid")



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also
