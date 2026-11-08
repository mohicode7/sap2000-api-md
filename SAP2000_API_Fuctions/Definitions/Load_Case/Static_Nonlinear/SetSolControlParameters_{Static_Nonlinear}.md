# SetSolControlParameters

## Syntax

SapObject.SapModel.LoadCases.StaticNonlinear.SetSolControlParameters

## VB6 Procedure

Function SetSolControlParameters(ByVal Name As String, ByVal MaxTotalSteps As Long, ByVal MaxFailedSubSteps As Long, ByVal MaxIterCS As Long, ByVal MaxIterNR As Long, ByVal TolConvD As Double, ByVal UseEventStepping As Boolean, ByVal TolEventD As Double, ByVal MaxLineSearchPerIter As Long, ByVal TolLineSearch As Double, ByVal LineSearchStepFact As Double) As Long

## Parameters

Name

The name of an existing static nonlinear load case.

MaxTotalSteps

The maximum total steps per stage.

MaxFailedSubSteps

The maximum null (zero) steps per stage.

MaxIterCS

The maximum constant-stiffness iterations per step.

MaxIterNR

The maximum Newton\_Raphson iterations per step.

TolConvD

The relative iteration convergence tolerance.

UseEventStepping

This item is True if event-to-event stepping is used.

TolEventD

The relative event lumping tolerance.

MaxLineSearchPerIter

The maximum number of line searches per iteration.

TolLineSearch

The relative line-search acceptance tolerance.

LineSearchStepFact

The line-search step factor.

## Remarks

This function sets the solution control parameters for the specified load case.

The function returns zero if the parameters are successfully set; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetCaseStaticNonlinearSolutionControlParameters()

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



   'set solution control parameters

      ret = SapModel.LoadCases.StaticNonlinear.SetSolControlParameters("LCASE1", 240, 40, 15, 50, 0.00005, False, 0.02, 10, 0.2, 1.7)



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

[GetSolControlParameters](GetSolControlParameters_%7BStatic_Nonlinear%7D.md)
