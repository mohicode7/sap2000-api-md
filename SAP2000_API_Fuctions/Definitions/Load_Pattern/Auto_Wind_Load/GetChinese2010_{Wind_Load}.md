# GetChinese2010

## Syntax

SapObject.SapModel.LoadPatterns.AutoWind.GetChinese2010

## VB6 Procedure

Function GetChinese2010\_1(ByVal Name As String, ByRef ExposureFrom As Long, ByRef DirAngle As Double, ByRef BuildingWidth As Double, ByRef Us As Double, ByRef UniformTaper As Boolean, ByRef BHoverB0 As Double, ByRef UserZ As Boolean, ByRef TopZ As Double, ByRef BottomZ As Double, ByRef wzero As Double, ByRef Rt As Long, ByRef PhiZOpt As Long, ByRef T1Opt As Long, ByRef UserT As Double, ByRef DampRatio As Double, ByRef UserExposure As Boolean) As Long

## Parameters

Name

The name of an existing Wind-type load pattern with an Chinese 2010 auto wind assignment.

ExposureFrom

This is 1 or 2 indicating the source of the wind exposure.

1 = From extents of rigid diaphragms

2 = From area objects

DirAngle

The direction angle for the wind load. This item applies only when ExposureFrom = 1.

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

This item only applies when the T1 source is user defined (T1Opt = 1). It is the user defined T1 period. [s]

DampRatio

The damping ratio.

UserExposure

If this item is True, the wind exposure widths are provided by the user. If it is False, the wind exposure widths are calculated by the program from the extents of the diaphragms.

## Remarks

This function retrieves auto wind loading parameters for Chinese 2010.

The function returns zero if the parameters are successfully assigned; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetWindChinese2010()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim ExposureFrom As Long

      Dim DirAngle As Double

      Dim BuildingWidth As Double

      Dim Us As Double

      Dim UniformTaper As Boolean

      Dim BHoverB0 As Double

      Dim UserZ As Boolean

      Dim TopZ As Double

      Dim BottomZ As Double

      Dim wzero As Double

      Dim Rt As Long

      Dim PhiZOpt As Long

      Dim T1Opt As Long

      Dim UserT As Double

      Dim DampRatio As Double

      Dim UserExposure As Boolean



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

      ret = SapModel.LoadPatterns.Add("WIND", LTYPE_WIND)



   'assign Chinese2010 parameters

      ret = SapModel.LoadPatterns.AutoWind.SetChinese2010("WIND", 1, 0, 1200, 0.5, False, 1, False, 0, 0, 0.48, 3, 1, 1, 0.6, 0.04)



   'get Chinese2010 parameters

      ret = SapModel.LoadPatterns.AutoWind.GetChinese2010("WIND", ExposureFrom, DirAngle, BuildingWidth, Us, UniformTaper, BHoverB0, UserZ, TopZ, BottomZ, wzero, Rt, PhiZOpt, T1Opt, UserT, DampRatio, UserExposure)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 15.0.2.

## See Also

[SetChinese2010](SetChinese2010_%7BWind_Load%7D.md)
