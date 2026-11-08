# GetSolControlParameters

## Syntax

SapObject.SapModel.LoadCases.ModHistNonlinear.GetSolControlParameters

## VB6 Procedure

Function GetSolControlParameters(ByVal Name As String, ByRef tstat As Double, ByRef dtmax As Double, ByRef dtmin As Double, ByRef ftol As Double, ByRef etol As Double, ByRef itmax As Long, ByRef itmin As Long, ByRef Cf As Double) As Long

## Parameters

Name

The name of an existing nonlinear modal time history analysis case.

tstat

The static period.

dtmax

The maximum substep size.

dtmin

The minimum substep size.

ftol

The relative force convergence tolerance.

etol

The relative energy convergence tolerance.

itmax

The maximum iteration limit.

itmin

The minimum iteration limit.

Cf

The convergence factor.

## Remarks

This function retrieves the solution control parameters for the specified load case.

The function returns zero if the parameters are successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetCaseModHistNonlinearSolutionControlParameters()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim tstat As Double

      Dim dtmax As Double

      Dim dtmin As Double

      Dim ftol As Double

      Dim etol As Double

      Dim itmax As Long

      Dim itmin As Long

      Dim Cf As Double



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



   'add nonlinear modal time history load case

      ret = SapModel.LoadCases.ModHistNonlinear.SetCase("LCASE1")



   'get solution control parameters

      ret = SapModel.LoadCases.ModHistNonlinear.GetSolControlParameters("LCASE1", tstat, dtmax, dtmin, ftol, etol, itmax, itmin, Cf)



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

[SetSolControlParameters](SetSolControlParameters_%7BModal_TH_Nonlinear%7D.md)
