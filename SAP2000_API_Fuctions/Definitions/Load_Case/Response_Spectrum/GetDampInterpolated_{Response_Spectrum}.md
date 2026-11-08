# GetDampInterpolated

## Syntax

SapObject.SapModel.LoadCases.ResponseSpectrum.GetDampInterpolated

## VB6 Procedure

Function GetDampInterpolated(ByVal Name As String, ByRef DampType As Long, ByRef NumberItems As Long, ByRef Time() As Double, ByRef Damp() As Double) As Long

## Parameters

Name

The name of an existing response spectrum load case that has interpolated damping.

DampType

This is either 5 or 6, indicating the interpolated modal damping type.

5 = Interpolated damping by period

6 = Interpolated damping by frequency

NumberItems

The number of Time and Damp pairs.

Time

This is an array that includes the period or the frequency depending on the value of the DampType item. [s] for DampType = 5 and [cyc/s] for DampType = 6

Damp

This is an array that includes the damping for the specified period of frequency (0 <= Damp < 1).

## Remarks

This function retrieves the interpolated modal damping data assigned to the specified load case.

The function returns zero if the option is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetCaseResponseSpectrumDampInterpolated()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim MyTime() As Double

      Dim MyDamp() As Double

      Dim DampType As Long

      Dim NumberItems As Long

      Dim Time() As Double

      Dim Damp() As Double



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



   'add response spectrum load case

      ret = SapModel.LoadCases.ResponseSpectrum.SetCase("LCASE1")



   'set interpolated damping

      ReDim MyTime(2)

      ReDim MyDamp(2)

      MyTime(0) = 0.001

      MyDamp(0) = 0.1

      MyTime(1) = 0.3

      MyDamp(1) = 0.03

      MyTime(2) = 1

      MyDamp(2) = 0.05

      ret = SapModel.LoadCases.ResponseSpectrum.SetDampInterpolated("LCASE1", 5, 3, MyTime, MyDamp)



   'get interpolated damping

      ret = SapModel.LoadCases.ResponseSpectrum.GetDampType("LCASE1", DampType)

      If DampType = 5 or DampType = 6 Then

         ret = SapModel.LoadCases.ResponseSpectrum.GetDampInterpolated("LCASE1", DampType, NumberItems, Time, Damp)

      End If



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

Changed nomenclature from Load Cases, Analysis Cases and Response Combinations to Load Patterns, Load Cases and Load Combinations, respectively, in version 12.00.

## See Also

[SetDampInterpolated](SetDampInterpolated_%7BResponse_Spectrum%7D.md)
