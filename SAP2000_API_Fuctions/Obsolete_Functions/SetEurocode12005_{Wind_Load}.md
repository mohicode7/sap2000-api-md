# SetEurocode12005 (Note:  Newer function available)

## Syntax

SapObject.SapModel.LoadPatterns.AutoWind.SetEurocode12005

## VB6 Procedure

Function SetEurocode12005(ByVal Name As String, ByVal ExposureFrom As Long, ByVal DirAngle As Double, ByVal Cpw As Double, ByVal Cpl As Double, ByVal UserZ As Boolean, ByVal TopZ As Double, ByVal BottomZ As Double, ByVal WindSpeed As Double, ByVal Terrain As Long, ByVal Orography As Double, ByVal k1 As Double, ByVal CsCd As Double, Optional ByVal UserExposure As Boolean = False) As Long

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

WindSpeed

The basic wind speed, vb, in meters per second.

Terrain

This is 0, 1, 2, 3 or 4, indicating the terrain category.

0 = 0

1 = I

2 = II

3 = III

4 = IV

Orography

The orography factor, Co.

k1

The turbulence factor, k1.

CsCd

The structural factor, CsCd.

UserExposure

If this item is True, the wind exposure widths are provided by the user. If it is False, the wind exposure widths are calculated by the program from the extents of the diaphragms.

## Remarks

This function assigns auto wind loading parameters for Eurocode 1 2005.

The function returns zero if the parameters are successfully assigned; otherwise it returns a nonzero value.

## VBA Example



```
Sub AssignWindEurocode12005()

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



   'assign Eurocode 1 2005 parameters

      ret = SapModel.LoadPatterns.AutoWind.SetEurocode12005("WIND", 1, 0, 0.8, 0.5, False, 0, 0, 35, 2, 1, 1, 1)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 14.00.

The function is obsolete and has been superseded by [SetEurocode12005\_1](../Definitions/Load_Pattern/Auto_Wind_Load/SetEurocode12005_1%7BWind_Load%7D.md) as of version 14.1.0. This function is maintained for backward compatibility. New function added.

## See Also

[GetEurocode12005](GetEurocode12005_%7BWind_Load%7D.md)
