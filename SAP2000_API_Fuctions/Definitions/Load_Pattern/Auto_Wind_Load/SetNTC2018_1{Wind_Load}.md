# SetNTC2018\_1

## Syntax

SapObject.SapModel.LoadPatterns.AutoWind.SetNTC2018\_1

## VB6 Procedure

Function SetNTC2018\_1(ByVal Name As String, ByRef ExposureFrom As Long, ByRef DirAngle As Double, ByRef Cpw As Double, ByRef Cpl As Double, ByRef UserZ As Boolean, ByRef TopZ As Double, ByRef BottomZ As Double, ByRef vr As Double, ByRef ExposureCategory As Long, ByRef ct As Double, ByRef cd As Double, ByRef UserExposure As Boolean = False) As Long

## Parameters

Name

The name of an existing Wind-type load pattern.

ExposureFrom

This is 1 or 2, indicating the source of the wind exposure.

1 = From extents of rigid diaphragms

2 = From area objects

DirAngle

The direction angle for the wind load. This item applies only when ExposureFrom = 1.

Cpw

The windward coefficient, Cp. This item applies only when ExposureFrom = 1.

Cpl

The leeward coefficient, Cp. This item applies only when ExposureFrom = 1.

UserZ

This item is True if the top and bottom elevations of the wind load are user specified. It is False if the elevations are determined by the program.

TopZ

This item applies only when the UserZ item is True. It is the global Z-coordinate at the highest level where auto wind loads are applied. [L]

BottomZ

This item applies only when the UserZ item is True. It is the global Z-coordinate at the lowest level where auto wind loads are applied. [L]

vr

The wind velocity in m/s.

ExposureCategory

This is 1, 2, 3, 4, or 5, indicating the exposure category.

1 = I

2 = II

3 = III

4 = IV

5 = V

ct

The topography factor, ct.

cd

The dynamic coefficient, cd.

UserExposure

If this item is True, the wind exposure widths are provided by the user. If it is False, the wind exposure widths are calculated by the program from the extents of the diaphragms.

## Remarks

This function retrieves auto wind loading parameters for NTC 2018.

The function returns zero if the parameters are successfully assigned; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetWindNTC2018()

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

    ret = SapModel.File.New3DFrame(BeamSlab, 2, 144, 3, 336, 2, 432)



    'define diaphragm constraints

    ret = SapModel.ConstraintDef.SetDiaphragm("Diaph1", Z)

    ret = SapModel.ConstraintDef.SetDiaphragm("Diaph2", Z)



    'assign points to diaphragm

    ret = SapModel.SelectObj.ClearSelection

    ret = SapModel.SelectObj.PlaneXY("2")

    ret = SapModel.PointObj.SetConstraint("", "Diaph1", SelectedObjects)

    ret = SapModel.SelectObj.ClearSelection

    ret = SapModel.SelectObj.PlaneXY("3")

    ret = SapModel.PointObj.SetConstraint("", "Diaph2", SelectedObjects)

    ret = SapModel.SelectObj.ClearSelection



    'add new load pattern

    ret = SapModel.LoadPatterns.Add("WIND", eLoadPatternType_WIND)



    'assign NTC2018 parameters

    ret = SapModel.LoadPatterns.AutoWind.SetNTC2018_1("WIND", 1, 0, 0.8, 0.5, False, 0, 0, 35, 3, 1, 1, False)



    'close Sap2000

    SapObject.ApplicationExit False

    Set SapModel = Nothing

    Set SapObject = Nothing

    End Sub
  
```



## Release Notes

Initial release in version 25.2.0.

This function supersedes [SetNTC2018](../../../Obsolete_Functions/SetNTC2018_%7BWind_Load%7D_obsolete.md).

## See Also

[GetNTC2018\_1](GetNTC2018_1%7BWind_Load%7D.md)
