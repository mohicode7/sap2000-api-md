# Delete

## Syntax

SapObject.SapModel.GDispl.Delete

## VB6 Procedure

Function Delete(ByVal Name As String) As Long

## Parameters

Name

The name of an existing generalized displacement.

## Remarks

This function deletes the specified generalized displacement.

The function returns zero if the generalized displacement is successfully deleted, otherwise it returns a nonzero value.

## VBA Example



```
Sub DeleteGDispl()

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



   'delete generalized displacement

      ret = SapModel.GDispl.Delete("GD1")



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[Add](Add_%7BGeneralized_Displacement%7D.md)
