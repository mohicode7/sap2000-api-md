# GetRamp

## Syntax

SapObject.SapModel.Func.FuncTH.GetRamp

## VB6 Procedure

Function GetRamp(ByVal Name As String, ByRef RampTime As Double, ByRef RampAmp As Double, ByRef RampMaxTime As Double) As Long

## Parameters

Name

The name of a ramp-type time history function.

RampTime

The time it takes for the ramp function to initially reach its maximum value. [s]

RampAmp

The maximum amplitude of the ramp function.

RampMaxTime

The time at the end of the ramp function. This time must be greater than the RampTime. [s]

## Remarks

This function retrieves the definition of a ramp-type time history function.

The function returns zero if the function definition is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetTHFuncRamp()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim RampTime As Double

      Dim RampAmp As Double

      Dim RampMaxTime As Double



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



   'get ramp TH function

      ret = SapModel.Func.FuncTH.GetRamp("TH-1", RampTime, RampAmp, RampMaxTime)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also

[SetRamp](SetRamp.md)
