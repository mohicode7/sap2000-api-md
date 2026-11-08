# SetSolControlParameters

## Syntax

SapObject.SapModel.LoadCases.ModHistNonlinear.SetSolControlParameters

## VB6 Procedure

Function SetSolControlParameters(ByVal Name As String, ByVal tstat As Double, ByVal dtmax As Double, ByVal dtmin As Double, ByVal ftol As Double, ByVal etol As Double, ByVal itmax As Long, ByVal itmin As Long, ByVal Cf As Double) As Long

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

This function sets the solution control parameters for the specified load case.

The function returns zero if the parameters are successfully set; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetCaseModHistNonlinearSolutionControlParameters()

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



   'add nonlinear modal time history load case

      ret = SapModel.LoadCases.ModHistNonlinear.SetCase("LCASE1")



   'set solution control parameters

      ret = SapModel.LoadCases.ModHistNonlinear.SetSolControlParameters("LCASE1", 0.1, 0.001, 0.000001, 0.00002, 0.00003, 80, 10, 1.5)



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

[GetSolControlParameters](GetSolControlParameters_%7BModal_TH_Nonlinear%7D.md)
