# GetSawtooth

## Syntax

SapObject.SapModel.Func.FuncTH.GetSawtooth

## VB6 Procedure

Function GetSawtooth(ByVal Name As String, ByRef SawP As Double, ByRef SawTime As Double, ByRef SawCycles As Long, ByRef SawAmp As Double) As Long

## Parameters

Name

The name of a sawtooth-type time history function.

SawP

The period of the sawtooth function. [s]

SawTime

The time it takes for the sawtooth function to ramp up from a function value of zero to its maximum amplitude. [s]

SawCycles

The number of cycles in the function.

SawAmp

The maximum amplitude of the sawtooth function.

## Remarks

This function retrieves the definition of a sawtooth-type time history function.

The function returns zero if the function definition is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetTHFuncSawtooth()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim SawP As Double

      Dim SawTime As Double

      Dim SawCycles As Long

      Dim SawAmp As Double



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



   'get sawtooth TH function

      ret = SapModel.Func.FuncTH.GetSawtooth("TH-1", SawP, SawTime, SawCycles, SawAmp)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also

[SetSawtooth](SetSawtooth.md)
