# SetSawtooth

## Syntax

SapObject.SapModel.Func.FuncTH.SetSawtooth

## VB6 Procedure

Function SetSawtooth(ByVal Name As String, ByVal SawP As Double, ByVal SawTime As Double, ByVal SawCycles As Long, ByVal SawAmp As Double) As Long

## Parameters

Name

The name of an existing or new function. If this is an existing function, that function is modified; otherwise, a new function is added.

SawP

The period of the sawtooth function. [s]

SawTime

The time it takes for the sawtooth function to ramp up from a function value of zero to its maximum amplitude. [s]

SawCycles

The number of cycles in the function.

SawAmp

The maximum amplitude of the sawtooth function.

## Remarks

This function defines a sawtooth-type time history function.

The function returns zero if the function is successfully defined; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetTHFuncSawtooth()

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



   'add sawtooth TH function

      ret = SapModel.Func.FuncTH.SetSawtooth("TH-1", 1, 0.25, 4, 1.75)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also

[GetSawtooth](GetSawtooth.md)
