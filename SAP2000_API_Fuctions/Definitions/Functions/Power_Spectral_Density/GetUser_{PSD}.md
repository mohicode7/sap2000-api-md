# GetUser

## Syntax

SapObject.SapModel.Func.FuncPSD.GetUser

## VB6 Procedure

Function GetUser(ByVal Name As String, ByRef NumberItems As Long, ByRef Frequency() As Double, ByRef Value() As Double) As Long

## Parameters

Name

The name of a user defined power spectral density function.

NumberItems

The number of frequency and value pairs defined.

Frequency

This is an array that includes the frequency in Hz for each data point. [cyc/s]

Value

This is an array that includes the function value for each data point.

## Remarks

This function retrieves the definition of a user defined power spectral density function.

The function returns zero if the function definition is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetPSDFuncUser()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Num As Long

      Dim Freq() As Double

      Dim Val() As Double

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

      Num = 4

      ReDim Freq(Num - 1)

      ReDim Val(Num - 1)

      Freq(0) = 0:  Val(0) = 1

      Freq(1) = 1:  Val(1) = 1

      Freq(2) = 2:  Val(2) = 2

      Freq(3) = 3:  Val(3) = 2

      ret = SapModel.Func.FuncPSD.SetUser("PSD-1", Num, Freq, Val)



   'get user PSD function

      ret = SapModel.Func.FuncPSD.GetUser("PSD-1", NumberItems, Frequency, Value)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also

[SetUser](SetUser_%7BPSD%7D.md)
