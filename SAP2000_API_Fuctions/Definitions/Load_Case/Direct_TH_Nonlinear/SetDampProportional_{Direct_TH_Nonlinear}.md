# SetDampProportional

## Syntax

SapObject.SapModel.LoadCases.DirHistNonlinear.SetDampProportional

## VB6 Procedure

Function SetDampProportional(ByVal Name As String, ByVal DampType As Long, ByVal Dampa As Double, ByVal Dampb As Double, ByVal Dampf1 As Double, ByVal Dampf2 As Double, ByVal Dampd1 As Double, ByVal Dampd2 As Double) As Long

## Parameters

Name

The name of an existing nonlinear direct integration time history load case.

DampType

This is 1, 2 or 3, indicating the proportional modal damping type.

1 = Mass and stiffness proportional damping by direct specification

2 = Mass and stiffness proportional damping by period

3 = Mass and stiffness proportional damping by frequency

Dampa

The mass proportional damping coefficient. This item applies only when DampType = 1.

Dampb

The stiffness proportional damping coefficient. This item applies only when DampType = 1.

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

This function sets proportional modal damping data for the specified load case.

The function returns zero if the damping is successfully set; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetCaseDirHistNonlinearDampProportional()

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



   'add nonlinear direct history load case

      ret = SapModel.LoadCases.DirHistNonlinear.SetCase("LCASE1")



   'set proportional damping

      ret = SapModel.LoadCases.DirHistNonlinear.SetDampProportional("LCASE1", 2, 0, 0, 0.1, 1, 0.05, 0.06)



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

[GetDampProportional](GetDampProportional_%7BDirect_TH_Nonlinear%7D.md)
