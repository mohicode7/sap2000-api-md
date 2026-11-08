# Delete

## Syntax

SapObject.SapModel.RespCombo.Delete

## VB6 Procedure

Function Delete(ByVal Name As String) As Long

## Parameters

Name

The name of an existing load combination.

## Remarks

This function deletes the specified load combination.

The function returns zero if the combination is successfully deleted, otherwise it returns a nonzero value.

## VBA Example



```
Sub DeleteCombo()

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



   'add combo

      ret = SapModel.RespCombo.Add("COMB1", 0)



   'delete combo

      ret = SapModel.RespCombo.Delete("COMB1")



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

Changed nomenclature from Load Cases, Analysis Cases and Response Combinations to Load Patterns, Load Cases and Load Combinations, respectively, in version 12.00.

## See Also

[Add](Add_%7BCombinations%7D.md)
