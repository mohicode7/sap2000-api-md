# SetOptionSteadyState

## Syntax

SapObject.SapModel.Results.Setup.SetOptionSteadyState

## VB6 Procedure

Function SetOptionSteadyState(ByVal Value As Long, ByVal SteadyStateOption As Long) As Long

## Parameters

Value

This item is either 1 or 2

1 = Envelopes

2 = At Frequencies

SteadyStateOption

This item is 1, 2 or 3

1 = In and Out of Phase

2 = Magnitude

3 = All

## Remarks

This function sets the output option for steady state results.

The function returns 0 if the output option is successfully set, otherwise it returns nonzero.

## VBA Example



```
Sub SetSteadyStateOutputOption()

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

      ret = SapModel.File.New2DFrame(PortalFrame, 3, 124, 3, 200)



   'set output option

      ret = SapModel.Results.Setup.SetOptionSteadyState(1, 1)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[GetOptionSteadyState](GetOptionSteadyState.md)
