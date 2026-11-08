# GetTargetForceParameters

## Syntax

SapObject.SapModel.LoadCases.StaticNonlinearStaged.GetTargetForceParameters

## VB6 Procedure

Function GetTargetForceParameters(ByVal Name As String, ByRef TolConvF As Double, ByRef MaxIter As Long, ByRef AccelFact As Double, ByRef NoStop As Boolean) As Long

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

This function retrieves the target force iteration parameters for the specified load case.

The function returns zero if the parameters are successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetCaseStaticNonlinearStagedTargetForceParameters()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim TolConvF As Double

      Dim MaxIter As Long

      Dim AccelFact As Double

      Dim NoStop As Boolean



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



   'get target force iteration parameters

      ret = SapModel.LoadCases.StaticNonlinearStaged.GetTargetForceParameters("LCASE1", TolConvF, MaxIter, AccelFact, NoStop)



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

[SetTargetForceParameters](SetTargetForceParameters_%7BStatic_Nonlinear_Staged%7D.md)
