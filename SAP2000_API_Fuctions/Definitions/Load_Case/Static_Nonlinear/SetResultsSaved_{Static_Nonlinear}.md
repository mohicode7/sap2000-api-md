# SetResultsSaved

## Syntax

SapObject.SapModel.LoadCases.StaticNonlinear.SetResultsSaved

## VB6 Procedure

Function SetResultsSaved(ByVal Name As String, ByVal SaveMultipleSteps As Boolean, Optional ByVal MinSavedStates As Long = 10, Optional ByVal MaxSavedStates As Long = 100, Optional ByVal PositiveOnly As Boolean = True) As Long

## Parameters

Name

The name of an existing static nonlinear load case.

SaveMultipleSteps

This item is True if multiple states are saved for the nonlinear analysis. It is False only if the final state is saved.

MinSavedStates

This item only applies when SaveMultipleSteps = True. It is the minimum number of saved steps.

MaxSavedStates

This item only applies when SaveMultipleSteps = True. It is the maximum number of saved steps.

PositiveOnly

If this item is True, only positive displacement increments are saved. If it is False, all displacement increments are saved.

## Remarks

This function sets the results saved parameters for the specified load case.

The function returns zero if the parameters are successfully set; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetCaseStaticNonlinearResultsSaved()

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

      ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288)



   'add static nonlinear load case

      ret = SapModel.LoadCases.StaticNonlinear.SetCase("LCASE1")



   'set results saved parameters

      ret = SapModel.LoadCases.StaticNonlinear.SetResultsSaved("LCASE1", True, 12, 84, False)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

Changed nomenclature from Load Cases, Analysis Cases and Response Combinations to Load Patterns, Load Cases and Load Combinations, respectively, in version 12.00.

## See Also

[GetResultsSaved](GetResultsSaved_%7BStatic_Nonlinear%7D.md)
