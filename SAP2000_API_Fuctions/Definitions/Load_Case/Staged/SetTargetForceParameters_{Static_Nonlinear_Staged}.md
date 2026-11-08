# SetTargetForceParameters

## Syntax

SapObject.SapModel.LoadCases.StaticNonlinearStaged.SetTargetForceParameters

## VB6 Procedure

Function SetTargetForceParameters(ByVal Name As String, ByVal TolConvF As Double, ByVal MaxIter As Long, ByVal AccelFact As Double, ByVal NoStop As Boolean) As Long

## Parameters

Name

The name of an existing static nonlinear staged analysis case.

TolConvF

The relative convergence tolerance for target force iteration.

MaxIter

The maximum iterations per stage for target force iteration.

AccelFact

The acceleration factor.

NoStop

If this item is True, the analysis is continued when there is no convergence in the target force iteration.

## Remarks

This function sets the target force iteration parameters for the specified load case.

The function returns zero if the parameters are successfully set; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetCaseStaticNonlinearStagedTargetForceParameters()

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



   'add static nonlinear staged load case

      ret = SapModel.LoadCases.StaticNonlinearStaged.SetCase("LCASE1")



   'set target force iteration parameters

      ret = SapModel.LoadCases.StaticNonlinearStaged.SetTargetForceParameters("LCASE1", 0.008, 6, 5, True)



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

[GetTargetForceParameters](GetTargetForceParameters_%7BStatic_Nonlinear_Staged%7D.md)
