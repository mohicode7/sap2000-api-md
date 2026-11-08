# SetInitialCase

## Syntax

SapObject.SapModel.LoadCases.ModHistNonlinear.SetInitialCase

## VB6 Procedure

Function SetInitialCase(ByVal Name As String, ByVal InitialCase As String) As Long

## Parameters

Name

The name of an existing nonlinear modal history analysis case.

InitialCase

This is blank, None or the name of an existing analysis case. This item specifies if the load case starts from zero initial conditions, that is, an unstressed state, or if it continues from the end of another nonlinear modal time history load case.

If the specified initial case is not a nonlinear modal time history load case, zero initial conditions are assumed.

## Remarks

This function sets the initial condition for the specified load case.

The function returns zero if the initial condition is successfully set; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetCaseModHistNonlinearInitialCondition()

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



   'add nonlinear modal history load cases

      ret = SapModel.LoadCases.ModHistNonlinear.SetCase("LCASE1")

      ret = SapModel.LoadCases.ModHistNonlinear.SetCase("LCASE2")



   'set initial condition

      ret = SapModel.LoadCases.ModHistNonlinear.SetInitialCase("LCASE1", "LCASE2")



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

[GetInitialCase](GetInitialCase_%7BModal_TH_Nonlinear%7D.md)
