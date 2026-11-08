# SetSine

## Syntax

SapObject.SapModel.Func.FuncTH.SetSine

## VB6 Procedure

Function SetSine(ByVal Name As String, ByVal SineP As Double, ByVal SineSteps As Long, ByVal SineCycles As Long, ByVal SineAmp As Double) As Long

## Parameters

Name

The name of an existing or new function. If this is an existing function, that function is modified; otherwise, a new function is added.

SineP

The period of the sine function. [s]

SineSteps

The number of steps in the sine function. This item can not be less than 8.

SineCycles

The number of cycles in the sine function.

SineAmp

The amplitude of the sine function.

## Remarks

This function defines a sine-type time history function.

The function returns zero if the function is successfully defined; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetTHFuncSine()

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



   'add sine TH function

      ret = SapModel.Func.FuncTH.SetSine("TH-1", 1, 16, 4, 1.25)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also

[GetSine](GetSine.md)
