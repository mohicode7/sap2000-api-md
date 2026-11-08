# GetSine

## Syntax

SapObject.SapModel.Func.FuncTH.GetSine

## VB6 Procedure

Function GetSine(ByVal Name As String, ByRef SineP As Double, ByRef SineSteps As Long, ByRef SineCycles As Long, ByRef SineAmp As Double) As Long

## Parameters

Name

The name of a sine-type time history function.

SineP

The period of the sine function. [s]

SineSteps

The number of steps in the sine function. This item can not be less than 8.

SineCycles

The number of cycles in the sine function.

SineAmp

The amplitude of the sine function.

## Remarks

This function retrieves the definition of a sine-type time history function.

The function returns zero if the function definition is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetTHFuncSine()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim SineP As Double

      Dim SineSteps As Long

      Dim SineCycles As Long

      Dim SineAmp As Double



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



   'add sine TH function

      ret = SapModel.Func.FuncTH.SetSine("TH-1", 1, 16, 4, 1.25)



   'get sine TH function

      ret = SapModel.Func.FuncTH.GetSine("TH-1", SineP, SineSteps, SineCycles, SineAmp)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also

[SetSine](SetSine.md)
