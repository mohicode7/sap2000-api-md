# SetDampInterpolated

## Syntax

SapObject.SapModel.LoadCases.PSD.SetDampInterpolated

## VB6 Procedure

Function SetDampInterpolated(ByVal Name As String, ByVal HysIntFreqUnits As Long, ByVal HysIntNumFreqs As Long, ByRef HysIntFreq() As Double, ByRef HysIntMassCoeff() As Double, ByRef HysIntStiffCoeff() As Double) As Long

## Parameters

Name

The name of an existing power spectral density analysis case.

HysIntFreqUnits

This is either 1 or 2, indicating the units for the frequency.

1 = Hz [cyc/s]

2 = RPM

HysIntNumFreqs

The number of sets of frequency, mass coefficient and stiffness coefficient data.

HysIntFreq

This is an array of frequencies. The frequency is either in Hz or RPM depending on the value of HysIntFreqUnits.

HysIntMassCoeff

This is an array that includes the mass proportional damping coefficient.

HysIntStiffCoeff

This is an array that includes the stiffness proportional damping coefficient.

## Remarks

This function sets interpolated hysteretic damping by frequency for the specified load case.

The function returns zero if the damping is successfully set; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetCasePSDDampInterpolated()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim MyHysIntFreq() As Double

      Dim MyHysIntMassCoeff() As Double

      Dim MyHysIntStiffCoeff() As Double



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



   'add power spectral density load case

      ret = SapModel.LoadCases.PSD.SetCase("LCASE1")



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

      ret = SapModel.LoadCases.PSD.SetDampInterpolated("LCASE1", 2, 3, MyHysIntFreq, MyHysIntMassCoeff, MyHysIntStiffCoeff)



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

[GetDampInterpolated](GetDampInterpolated_%7BPSD%7D.md)
