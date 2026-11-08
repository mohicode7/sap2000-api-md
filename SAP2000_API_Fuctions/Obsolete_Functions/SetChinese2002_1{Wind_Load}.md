# SetChinese2002\_1

## Syntax

SapObject.SapModel.LoadPatterns.AutoWind.SetChinese2002\_1

## VB6 Procedure

Function SetChinese2002\_1(ByVal Name As String, ByVal ExposureFrom As Long, ByVal DirAngle As Double, ByVal BuildingWidth As Double, ByVal Us As Double, ByVal UniformTaper As Boolean, ByVal BHoverB0 As Double, ByVal UserZ As Boolean, ByVal TopZ As Double, ByVal BottomZ As Double, ByVal wzero As Double, ByVal Rt As Long, ByVal PhiZOpt As Long, ByVal T1Opt As Long, ByVal UserT As Double, ByVal DampRatio As Double, Optional ByVal UserExposure As Boolean = False) As Long

## Parameters

Name

The name of an existing Wind-type load pattern.

ExposureFrom

This is 1 or 2, indicating the source of the wind exposure.

1 = From extents of rigid diaphragms

2 = From area objects

DirAngle

The direction angle for the wind load. This item only applies when ExposureFrom = 1.

BuildingWidth

The building width. [L]

Us

The shape coefficient. This item applies only when ExposureFrom = 1.

UniformTaper

This item is True if a correction is to be applied to the wind load for a uniform taper.

BHoverB0

The taper ratio, Bh/B0. This item applies only when UniformTaper = True.

UserZ

This item is True if the top and bottom elevations of the wind load are user specified. It is False if the elevations are determined by the program.

TopZ

This item applies only when the UserZ item is True. It is the global Z-coordinate at the highest level where auto wind loads are applied. [L]

BottomZ

This item applies only when the UserZ item is True. It is the global Z-coordinate at the lowest level where auto wind loads are applied. [L]

wzero

The basic wind pressure in kN/m2.

Rt

This is 1, 2, 3 or 4, indicating the ground roughness.

1 = A

2 = B

3 = C

4 = D

PhiZOpt

This is 0 or 1, indicating the Phi Z source.

0 = Modal analysis

1 = Z/H ratio

T1Opt

This is 0 or 1, indicating the T1 source.

0 = Modal analysis

1 = User defined

UserT

This item applies only when the T1 source is user defined (T1Opt = 1). It is the user defined T1 period. [s]

DampRatio

The damping ratio.

UserExposure

If this item is True, the wind exposure widths are provided by the user. If it is False, the wind exposure widths are calculated by the program from the extents of the diaphragms.

## Remarks

This function assigns auto wind loading parameters for Chinese 2002.

The function returns zero if the parameters are successfully assigned; otherwise it returns a nonzero value.

## VBA Example



```
Sub AssignWindChinese2002_1()

   'dimension variables

      Dim SapObject As Sap2000v15.SapObject

      Dim SapModel As cSapModel

      Dim ret As Long



   'create Sap2000 object

      Set SapObject = New Sap2000v15.SapObject



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

      ret = SapModel.LoadPatterns.Add("WIND", LTYPE_WIND)



   'assign Chinese2002_1 parameters

      ret = SapModel.LoadPatterns.AutoWind.SetChinese2002_1("WIND", 1, 0, 1200, 0.5, False, 1, False, 0, 0, 0.48, 3, 1, 1, 0.6, 0.04)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 14.1.0.

This function supersedes [SetChinese2002 {Wind Load}](SetChinese2002_%7BWind_Load%7D.md)

## See Also

[GetChinese2002\_1](GetChinese2002_1%7BWind_Load%7D.md)
