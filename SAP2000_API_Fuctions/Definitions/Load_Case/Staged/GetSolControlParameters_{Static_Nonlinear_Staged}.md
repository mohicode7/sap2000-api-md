# GetSolControlParameters

## Syntax

SapObject.SapModel.LoadCases.StaticNonlinearStaged.GetSolControlParameters

## VB6 Procedure

Function GetSolControlParameters(ByVal Name As String, ByRef MaxTotalSteps As Long, ByRef MaxFailedSubSteps As Long, ByRef MaxIterCS As Long, ByRef MaxIterNR As Long, ByRef TolConvD As Double, ByRef UseEventStepping As Boolean, ByRef TolEventD As Double, ByRef MaxLineSearchPerIter As Long, ByRef TolLineSearch As Double, ByRef LineSearchStepFact As Double) As Long

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

This function retrieves the solution control parameters for the specified load case.

The function returns zero if the parameters are successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetCaseStaticNonlinearStagedSolutionControlParameters()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim MaxTotalSteps As Long

      Dim MaxFailedSubSteps As Long

      Dim MaxIterCS As Long

      Dim MaxIterNR As Long

      Dim TolConvD As Double

      Dim UseEventStepping As Boolean

      Dim TolEventD As Double

      Dim MaxLineSearchPerIter As Long

      Dim TolLineSearch As Double

      Dim LineSearchStepFact As Double



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



   'get solution control parameters

      ret = SapModel.LoadCases.StaticNonlinearStaged.GetSolControlParameters("LCASE1", MaxTotalSteps, MaxFailedSubSteps, MaxIterCS, MaxIterNR, TolConvD, UseEventStepping, TolEventD, MaxLineSearchPerIter, TolLineSearch, LineSearchStepFact)



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

[SetSolControlParameters](SetSolControlParameters_%7BStatic_Nonlinear_Staged%7D.md)
