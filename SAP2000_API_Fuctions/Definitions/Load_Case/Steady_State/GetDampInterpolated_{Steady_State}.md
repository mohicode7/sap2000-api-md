# GetDampInterpolated

## Syntax

SapObject.SapModel.LoadCases.SteadyState.GetDampInterpolated

## VB6 Procedure

Function GetDampInterpolated(ByVal Name As String, ByRef HysIntFreqUnits As Long, ByRef HysIntNumFreqs As Long, ByRef HysIntFreq() As Double, ByRef HysIntMassCoeff() As Double, ByRef HysIntStiffCoeff() As Double) As Long

## Parameters

Name

The name of an existing steady state load case that has interpolated damping.

HysIntFreqUnits

This is 1 or 2, indicating the units for the frequency.

1 = Hz [cyc/s]

2 = RPM

HysIntNumFreqs

The number of sets of frequency, mass coefficient and stiffness coefficient data.

HysIntFreq

This is an array of frequencies. The frequency is in Hz or RPM, depending on the value of HysIntFreqUnits.

HysIntMassCoeff

This is an array that includes the mass proportional damping coefficient.

HysIntStiffCoeff

This is an array that includes the stiffness proportional damping coefficient.

## Remarks

This function retrieves the interpolated hysteretic damping by frequency assigned to the specified load case.

The function returns zero if the damping is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetCaseSteadyStateDampInterpolated()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim MyHysIntFreq() As Double

      Dim MyHysIntMassCoeff() As Double

      Dim MyHysIntStiffCoeff() As Double

      Dim DampType As Long

      Dim HysIntFreqUnits As Long

      Dim HysIntNumFreqs As Long

      Dim HysIntFreq() As Double

      Dim HysIntMassCoeff() As Double

      Dim HysIntStiffCoeff() As Double



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



   'add steady state load case

      ret = SapModel.LoadCases.SteadyState.SetCase("LCASE1")



   'set interpolated damping

      ReDim MyHysIntFreq(2)

      ReDim MyHysIntMassCoeff(2)

      ReDim MyHysIntStiffCoeff(2)

      MyHysIntFreq(0) = 1

      MyHysIntMassCoeff(0) = 0.6

      MyHysIntStiffCoeff(0) = 0.04

      MyHysIntFreq(1) = 10

      MyHysIntMassCoeff(1) = 0.7

      MyHysIntStiffCoeff(1) = 0.05

      MyHysIntFreq(2) = 100

      MyHysIntMassCoeff(2) = 0.8

      MyHysIntStiffCoeff(2) = 0.08

      ret = SapModel.LoadCases.SteadyState.SetDampInterpolated("LCASE1", 2, 3, MyHysIntFreq, MyHysIntMassCoeff, MyHysIntStiffCoeff)



   'get interpolated damping

      ret = SapModel.LoadCases.SteadyState.GetDampType("LCASE1", DampType)

      If DampType = 2 Then

         ret = SapModel.LoadCases.SteadyState.GetDampInterpolated("LCASE1", HysIntFreqUnits, HysIntNumFreqs, HysIntFreq, HysIntMassCoeff, HysIntStiffCoeff)

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

[SetDampInterpolated](SetDampInterpolated_%7BSteady_State%7D.md)
