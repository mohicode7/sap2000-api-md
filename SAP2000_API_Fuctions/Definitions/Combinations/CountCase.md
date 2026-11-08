# CountCase

## Syntax

SapObject.SapModel.RespCombo.CountCase

## VB6 Procedure

Function CountCase(ByVal Name As String, ByRef Count As Long) As Long

## Parameters

Name

The name of an existing load combination.

Count

The number of load case and/or combinations included in the specified combination.

## Remarks

This function retrieves the total number of load case and/or combinations included in a specified load combination.

The function returns zero if the count is successfully retrieved, otherwise it returns a nonzero value.

## VBA Example



```
Sub CountComboCases()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Count As Long



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



   'add load case to combo

      ret = SapModel.RespCombo.SetCaseList("COMB1", LoadCase, "DEAD", 1.4)



   'count cases and combos in combo COMB1

      ret = SapModel.RespCombo.CountCase("COMB1", Count)



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

[GetCaseList](GetCaseList.md)

[SetCaseList](SetCaseList.md)
