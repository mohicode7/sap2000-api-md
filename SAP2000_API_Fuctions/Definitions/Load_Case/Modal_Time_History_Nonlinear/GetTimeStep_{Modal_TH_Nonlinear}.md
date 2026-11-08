# GetTimeStep

## Syntax

SapObject.SapModel.LoadCases.ModHistNonlinear.GetTimeStep

## VB6 Procedure

Function GetTimeStep(ByVal Name As String, ByRef nstep As Long, ByRef DT As Double) As Long

## Parameters

Name

The name of an existing nonlinear modal history analysis case.

nstep

The number of output time steps.

DT

The output time step size.

## Remarks

This function retrieves the time step data for the specified load case.

The function returns zero if the data is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetCaseModHistNonlinearTimeStep()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim nstep As Long

      Dim DT As Double



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



   'add nonlinear modal history load case

      ret = SapModel.LoadCases.ModHistNonlinear.SetCase("LCASE1")



   'get time step data

      ret = SapModel.LoadCases.ModHistNonlinear.GetTimeStep ("LCASE1", nstep, DT )



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also

[SetTimeStep](SetTimeStep_%7BModal_TH_Nonlinear%7D.md)
