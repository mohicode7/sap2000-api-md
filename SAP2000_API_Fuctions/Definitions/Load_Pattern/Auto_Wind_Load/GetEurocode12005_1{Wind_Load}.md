# GetEurocode12005\_1

## Syntax

SapObject.SapModel.LoadPatterns.AutoWind.GetEurocode12005\_1

## VB6 Procedure

Function GetEurocode12005\_1(ByVal Name As String, ByRef ExposureFrom As Long, ByRef DirAngle As Double, ByRef Cpw As Double, ByRef Cpl As Double, ByRef UserZ As Boolean, ByRef TopZ As Double, ByRef BottomZ As Double, ByRef WindSpeed As Double, ByRef Terrain As Long, ByRef Orography As Double, ByRef k1 As Double, ByRef CsCd As Double, ByRef Rho As Double, ByRef UserExposure As Boolean) As Long

## Parameters

Name

The name of an existing Wind-type load pattern with a Eurocode 1 2005 auto wind assignment.

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

This item is the global Z-coordinate at the highest level where auto wind loads are applied. [L]

BottomZ

This item is the global Z-coordinate at the lowest level where auto wind loads are applied. [L]

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

Rho

The air density in kg/m3, Rho.

UserExposure

If this item is True, the wind exposure widths are provided by the user. If it is False, the wind exposure widths are calculated by the program from the extents of the diaphragms.

## Remarks

This function retrieves auto wind loading parameters for Eurocode 1 2005.

The function returns zero if the parameters are successfully assigned; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetWindEurocode12005()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim ExposureFrom As Long

      Dim DirAngle As Double

      Dim Cpw As Double

      Dim Cpl As Double

      Dim UserZ As Boolean

      Dim TopZ As Double

      Dim BottomZ As Double

      Dim WindSpeed As Double

      Dim Terrain As Long

      Dim Orography As Double

      Dim k1 As Double

      Dim CsCd As Double

      Dim Rho As Double

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



   'assign Eurocode 1 2005 parameters

      ret = SapModel.LoadPatterns.AutoWind.SetEurocode12005_1("WIND", 1, 0, 0.8, 0.5, False, 0, 0, 35, 2, 1, 1, 1)



   'get Eurocode 1 2005 parameters

      ret = SapModel.LoadPatterns.AutoWind.GetEurocode12005_1("WIND", ExposureFrom, DirAngle, Cpw, Cpl, UserZ, TopZ, BottomZ, WindSpeed, Terrain, Orography, k1, CsCd, UserExposure)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 14.1.0.

This function supersedes [GetEurocode12005](../../../Obsolete_Functions/GetEurocode12005_%7BWind_Load%7D.md).

## See Also

[SetEurocode12005\_1](SetEurocode12005_1%7BWind_Load%7D.md)
