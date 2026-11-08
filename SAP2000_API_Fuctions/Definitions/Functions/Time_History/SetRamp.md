# SetRamp

## Syntax

SapObject.SapModel.Func.FuncTH.SetRamp

## VB6 Procedure

Function SetRamp(ByVal Name As String, ByVal RampTime As Double, ByVal RampAmp As Double, ByVal RampMaxTime As Double) As Long

## Parameters

Name

The name of an existing or new function. If this is an existing function, that function is modified; otherwise, a new function is added.

RampTime

The time it takes for the ramp function to initially reach its maximum value. [s]

RampAmp

The maximum amplitude of the ramp function.

RampMaxTime

The time at the end of the ramp function. This time must be greater than the RampTime. [s]

## Remarks

This function defines a ramp-type time history function.

The function returns zero if the function is successfully defined; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetTHFuncRamp()

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



   'add ramp TH function

      ret = SapModel.Func.FuncTH.SetRamp("TH-1", 1, 1.25, 8)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also

[GetRamp](GetRamp.md)
