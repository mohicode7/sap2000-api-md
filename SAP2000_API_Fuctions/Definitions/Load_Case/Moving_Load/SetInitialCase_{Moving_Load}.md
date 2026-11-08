# SetInitialCase

## Syntax

SapObject.SapModel.LoadCases.Moving.SetInitialCase

## VB6 Procedure

Function SetInitialCase(ByVal Name As String, ByVal InitialCase As String) As Long

## Parameters

Name

The name of an existing moving load load case.

InitialCase

This is blank, None, or the name of an existing analysis case. This item specifies if the load case starts from zero initial conditions, that is, an unstressed state, or if it starts using the stiffness that occurs at the end of a nonlinear static or nonlinear direct integration time history load case.

If the specified initial case is a nonlinear static or nonlinear direct integration time history load case, the stiffness at the end of that case is used. If the initial case is anything else, zero initial conditions are assumed.

## Remarks

This function sets the initial condition for the specified load case.

The function returns zero if the initial condition is successfully set; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetCaseMovingInitialCondition()

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



   'open existing model

      SapModel.File.OpenFile("C:\SapAPI\Example 1-030.SDB")



   'add static nonlinear load case

      ret = SapModel.LoadCases.StaticNonlinear.SetCase("SN1")



   'add moving load load case

      ret = SapModel.LoadCases.Moving.SetCase("LCASE1")



   'set initial condition

      ret = SapModel.LoadCases.Moving.SetInitialCase("LCASE1", "SN1")



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

[GetInitialCase](GetInitialCase_%7BMoving_Load%7D.md)
