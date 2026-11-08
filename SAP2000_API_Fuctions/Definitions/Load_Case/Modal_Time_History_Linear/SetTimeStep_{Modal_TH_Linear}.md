# SetTimeStep

## Syntax

SapObject.SapModel.LoadCases.ModHistLinear.SetTimeStep

## VB6 Procedure

Function SetTimeStep(ByVal Name As String, ByVal nstep As Long, ByVal DT As Double) As Long

## Parameters

Name

The name of an existing linear modal history analysis case.

nstep

The number of output time steps.

DT

The output time step size.

## Remarks

This function sets the time step data for the specified load case.

The function returns zero if the data is successfully set; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetCaseModHistLinearTimeStep()

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



   'add linear modal history load case

      ret = SapModel.LoadCases.ModHistLinear.SetCase("LCASE1")



   'set time step data

      ret = SapModel.LoadCases.ModHistLinear.SetTimeStep("LCASE1", 120, 0.05)



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

[GetTimeStep](GetTimeStep_%7BModal_TH_Linear%7D.md)
