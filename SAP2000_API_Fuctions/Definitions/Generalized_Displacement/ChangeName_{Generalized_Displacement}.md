# ChangeName

## Syntax

SapObject.SapModel.GDispl.ChangeName

## VB6 Procedure

Function ChangeName(ByVal Name As String, ByVal NewName As String) As Long

## Parameters

Name

The existing name of a defined generalized displacement.

NewName

The new name for the generalized displacement.

## Remarks

The function returns zero if the new name is successfully applied, otherwise it returns a nonzero value.

The new generalized displacement name must be different from all other generalized displacement names. If the name is not unique, an error will be returned.

## VBA Example



```
Sub ChangeGDisplName()

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



   'add generalized displacement

      ret = SapModel.GDispl.Add("GD1", 1)



   'change generalized displacement name

      ret = SapModel.GDispl.ChangeName("GD1", "MyGD")



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also
