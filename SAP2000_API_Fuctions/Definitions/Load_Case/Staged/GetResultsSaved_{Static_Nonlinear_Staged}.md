# GetResultsSaved

## Syntax

SapObject.SapModel.LoadCases.StaticNonlinearStaged.GetResultsSaved

## VB6 Procedure

Function GetResultsSaved(ByVal Name As String, ByRef StagedSaveOption As Long, ByRef StagedMinSteps As Long, ByRef StagedMinStepsTD As Long) As Long

## Parameters

Name

The name of an existing static nonlinear staged analysis case.

StagedSaveOption

This is 0, 1, 2 or 3, indicating the results saved option for the load case.

0 = End of final stage

1 = End of each stage

2 = Start and end of each stage

3 = Two or more times in each stage

StagedMinSteps

The minimum number of steps for application of instantaneous load. This item applies only when StagedSaveOption = 3.

StagedMinStepsTD

The minimum number of steps for analysis of time dependent items. This item applies only when StagedSaveOption = 3.

## Remarks

This function retrieves the results saved parameters for the specified load case.

The function returns zero if the parameters are retrieved successfully; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetCaseStaticNonlinearStagedResultsSaved()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim StagedSaveOption As Long

      Dim StagedMinSteps As Long

      Dim StagedMinStepsTD As Long



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

      ret = SapModel.LoadCases.StaticNonlinearStaged.SetCase("LCASE1")



   'set results saved parameters

      ret = SapModel.LoadCases.StaticNonlinearStaged.SetResultsSaved("LCASE1", 3, 4, 10)



   'get results saved parameters

      ret = SapModel.LoadCases.StaticNonlinearStaged.GetResultsSaved("LCASE1", StagedSaveOption , StagedMinSteps, StagedMinStepsTD)



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

[SetResultsSaved](SetResultsSaved_%7BStatic_Nonlinear_Staged%7D.md)
