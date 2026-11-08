# GetValues

## Syntax

SapObject.SapModel.Func.GetValues

## VB6 Procedure

Function GetValues(ByVal Name As String, ByRef NumberItems As Long, ByRef MyTime() As Double, ByRef Value() As Double) As Long

## Parameters

Name

The name of an existing function.

NumberItems

The number of time and function value pairs retrieved.

MyTime

This is an array that includes the time value for each data point. [s] for response spectrum and time history functions, [cyc/s] for power spectral density and steady state functions

Value

This is an array that includes the function value for each data point.

## Remarks

This function retrieves the time and function values for any defined function.

The function returns zero if the function definition is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetFuncValues()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim NumberItems As Long

      Dim MyTime() As Double

      Dim Value() As Double



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



   'get function values

      ret = SapModel.Func.GetValues("TH-1", NumberItems, MyTime, Value)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also
