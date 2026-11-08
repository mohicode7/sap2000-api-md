# GetResultsSaved

## Syntax

SapObject.SapModel.LoadCases.StaticNonlinear.GetResultsSaved

## VB6 Procedure

Function GetResultsSaved(ByVal Name As String, ByRef SaveMultipleSteps As Boolean, ByRef MinSavedStates As Long, ByRef MaxSavedStates As Long, ByRef PositiveOnly As Boolean) As Long

## Parameters

Name

The name of an existing static nonlinear load case.

SaveMultipleSteps

This item is True if multiple states are saved for the nonlinear analysis. It is False only if the final state is saved.

MinSavedStates

This item applies only when SaveMultipleSteps = True. It is the minimum number of saved steps.

MaxSavedStates

This item applies only when SaveMultipleSteps = True. It is the maximum number of saved steps.

PositiveOnly

If this item is True, only positive displacement increments are saved. If it is False, all displacement increments are saved.

## Remarks

This function retrieves the results saved parameters for the specified load case.

The function returns zero if the parameters are successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetCaseStaticNonlinearResultsSaved()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim SaveMultipleSteps As Boolean

      Dim MinSavedStates As Long

      Dim MaxSavedStates As Long

      Dim PositiveOnly As Boolean



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



   'get results saved parameters

      ret = SapModel.LoadCases.StaticNonlinear.GetResultsSaved("LCASE1", SaveMultipleSteps, MinSavedStates, MaxSavedStates, PositiveOnly)



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

[SetResultsSaved](SetResultsSaved_%7BStatic_Nonlinear%7D.md)
