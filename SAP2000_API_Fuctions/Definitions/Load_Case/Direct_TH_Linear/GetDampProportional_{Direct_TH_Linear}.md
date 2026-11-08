# GetDampProportional

## Syntax

SapObject.SapModel.LoadCases.DirHistLinear.GetDampProportional

## VB6 Procedure

Function GetDampProportional(ByVal Name As String, ByRef DampType As Long, ByRef Dampa As Double, ByRef Dampb As Double, ByRef Dampf1 As Double, ByRef Dampf2 As Double, ByRef Dampd1 As Double, ByRef Dampd2 As Double) As Long

## Parameters

Name

The name of an existing linear direct integration time history load case that has proportional damping.

DampType

This is 1, 2 or 3, indicating the proportional modal damping type.

1 = Mass and stiffness proportional damping by direct specification

2 = Mass and stiffness proportional damping by period

3 = Mass and stiffness proportional damping by frequency

Dampa

The mass proportional damping coefficient.

Dampb

The stiffness proportional damping coefficient.

Dampf1

This is the period or the frequency (depending on the value of the DampType item) for point 1. [s] for DampType = 2 and [cyc/s] for DampType = 3

This item applies only when DampType = 2 or 3.

Dampf2

This is the period or the frequency (depending on the value of the DampType item) for point 2. [s] for DampType = 2 and [cyc/s] for DampType = 3

This item applies only when DampType = 2 or 3.

Dampd1

This is the damping at point 1 (0 <= Dampd1 < 1).

This item applies only when DampType = 2 or 3.

Dampd2

This is the damping at point 2 (0 <= Dampd2 < 1).

This item applies only when DampType = 2 or 3.

## Remarks

This function retrieves the proportional modal damping data assigned to the specified load case.

The function returns zero if the option is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetCaseDirHistLinearDampProportional()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim DampType As Long

      Dim Dampa As Double

      Dim Dampb As Double

      Dim Dampf1 As Double

      Dim Dampf2 As Double

      Dim Dampd1 As Double

      Dim Dampd2 As Double



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



   'add linear direct history load case

      ret = SapModel.LoadCases.DirHistLinear.SetCase("LCASE1")



   'set proportional damping

      ret = SapModel.LoadCases.DirHistLinear.SetDampProportional("LCASE1", 2, 0, 0, 0.1, 1, 0.05, 0.06)



   'get proportional damping

      ret = SapModel.LoadCases.DirHistLinear.GetDampProportional("LCASE1", DampType, Dampa, Dampb, Dampf1, Dampf2, Dampd1, Dampd2)



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

[SetDampProportional](SetDampProportional_%7BDirect_TH_Linear%7D.md)
