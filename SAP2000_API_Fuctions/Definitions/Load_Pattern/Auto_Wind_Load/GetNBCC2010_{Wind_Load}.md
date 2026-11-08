# GetNBCC2010\_1 {Wind Load}

## Syntax

SapObject.SapModel.LoadPatterns.AutoWind.GetNBCC2010\_1

## VB6 Procedure

Function GetNBCC2010\_1(ByVal Name As String, ByRef ExposureFrom As Long, ByRef DirAngle As Double, ByRef Cpw As Double, ByRef Cpl As Double, ByRef UserZ As Boolean, ByRef TopZ As Double, ByRef BottomZ As Double, ByRef q As Double, ByRef GustFactor As Double, ByRef ImportanceFactor As Double, ByRef TerrainType As Long, ByRef CeWindward As Double, ByRef CeLeeward As Double, ByRef UserExposure As Boolean) As Long

## Parameters

Name

The name of an existing Wind-type load pattern with an NBCC 2010 auto wind assignment.

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

This item applies only when the UserZ item is True. It  is the global Z-coordinate at the lowest level where auto wind loads are applied. [L]

q

The velocity pressure in kPa.

GustFactor

The gust effect factor.

ImportanceFactor

The importance factor.

TerrainType

This is 1, 2, or 3, indicating the terrain type.

1 = Open

2 = Rough

3 = User

CeWindward

The windward exposure factor, Ce. This item applies only when TerrainType = 3.

CeLeeward

The windward exposure factor, Ce. This item applies only when TerrainType = 3.

UserExposure

If this item is True, the wind exposure widths are provided by the user. If it is False, the wind exposure widths are calculated by the program from the extents of the diaphragms.

## Remarks

This function retrieves auto wind loading parameters for NBCC 2010.

The function returns zero if the parameters are successfully assigned; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetWindNBCC2010()

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

      Dim q As Double

      Dim GustFactor As Double

      Dim ImportanceFactor As Double
      Dim TerrainType As Long
      Dim CeWindward As Double
      Dim CeLeeward As Double

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



   'assign NBCC2010 parameters

      ret = SapModel.LoadPatterns.AutoWind.SetNBCC2010_1("WIND", 1, 0, 0.8, 0.5, False, 0, 0, 0.75, 2.1, 1.1, 1, 0, 0)



   'get NBCC2010 parameters

      ret = SapModel.LoadPatterns.AutoWind.GetNBCC2010_1("WIND", ExposureFrom, DirAngle, Cpw, Cpl, UserZ, TopZ, BottomZ, q, GustFactor, ImportanceFactor, TerrainType, CeWindward, CeLeeward, UserExposure)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 19.0.0.

This function supersedes GetNBCCC2010.

## See Also

[SetNBCC2010\_1](SetNBCC2010_1.md)
