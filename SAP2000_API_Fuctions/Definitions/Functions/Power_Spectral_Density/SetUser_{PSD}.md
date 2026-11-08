# SetUser

## Syntax

SapObject.SapModel.Func.FuncPSD.SetUser

## VB6 Procedure

Function SetUser(ByVal Name As String, ByVal NumberItems As Long, ByRef Frequency() As Double, ByRef Value() As Double) As Long

## Parameters

Name

The name of an existing or new function. If this is an existing function, that function is modified; otherwise, a new function is added.

NumberItems

The number of frequency and value pairs defined.

Frequency

This is an array that includes the frequency in Hz for each data point. [cyc/s]

Value

This is an array that includes the function value for each data point.

## Remarks

This function defines a user power spectral density function.

The function returns zero if the function is successfully defined; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetPSDFuncUser()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim NumberItems As Long

      Dim Frequency() As Double

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



   'add user PSD function

      NumberItems = 4

      ReDim Frequency(NumberItems - 1)

      ReDim Value(NumberItems - 1)

      Frequency(0) = 0:  Value(0) = 1

      Frequency(1) = 1:  Value(1) = 1

      Frequency(2) = 2:  Value(2) = 2

      Frequency(3) = 3:  Value(3) = 2

      ret = SapModel.Func.FuncPSD.SetUser("PSD-1", NumberItems, Frequency, Value)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also

[GetUser](GetUser_%7BPSD%7D.md)
