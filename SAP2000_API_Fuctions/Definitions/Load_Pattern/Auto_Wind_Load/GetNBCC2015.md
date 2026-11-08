# GetNBCC2015{Wind Load}

## Syntax

SapObject.SapModel.LoadPatterns.AutoWind.GetNBCC2015

## VB6 Procedure

Function GetNBCC2015(ByVal Name As String, ByVal ExposureFrom As Long, ByRef DirAngle As Double, ByRef Cpw As Double, ByRef Cpl As Double, ByVal NBCCCase As Long, ByVal e1 As Double, ByVal e2 As Double, ByVal UserZ As Boolean, ByVal TopZ As Double, ByVal BottomZ As Double, ByRef q As Double, ByRef GustFactor As Double, ByVal TopographicFactor As Double, ByRef ImportanceFactor As Double, ByRef TerrainType As Long, ByRef CeWindward As Double, ByRef CeLeeward As Double, ByRef UserExposure As Boolean) As Long

## Parameters

Name

The name of an existing Wind-type load pattern with an NBCC 2015 auto wind assignment..

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

NBCCCase

This is 1, 2, 3, 4, or 5, indicating the desired case from NBCC 2105 Figure A-4.1.7.9(1). 1,2,3, and 4 refer to cases 1 through 4 in the figure, while 5 means all cases. This item applies only when ExposureFrom = 1.

e1

This is the value e1 in the NBCC 2015 Figure A-4.1.7.9(1). This item applies only when ExposureFrom = 1.

e2

This is the value e2 in the NBCC 2015 Figure A-4.1.7.9(1). This item applies only when ExposureFrom = 1.

UserZ

This item is True if the top and bottom elevations of the wind load are user specified. It is False if the elevations are determined by the program.

TopZ

This item applies only when the UserZ item is True. It is the global Z-coordinate at the highest level where auto wind loads are applied. [L]

BottomZ

This item applies only when the UserZ item is True. It  is the global Z-coordinate at the lowest level where auto wind loads are applied. [L]

q

The velocity pressure in kPa.

GustFactor

The gust effect factor, Cg.

TopographicFactor

The importance factor, Ct.

ImportanceFactor

The importance Factor, Iw.

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

This function retrieves auto wind loading parameters for NBCC 2015.

The function returns zero if the parameters are successfully assigned; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetWindNBCC2015()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long
      Dim ExposureFrom As Long
      Dim DirAngle As Double

      Dim Cpw As Double

      Dim Cpl As Double

      Dim NBCCCase As Long
      Dim e1 As Double
      Dim e2 As Double
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
   'assign NBCC2015 parameters

      ret = SapModel.LoadPatterns.AutoWind.SetNBCC2015("WIND", 1, 0, 0.8, 0.5, False, 0, 0, 0.75, 2.1, 1.1, 1, 0, 0)
   'get NBCC2015 parameters

      ret = SapModel.LoadPatterns.AutoWind.GetNBCC2015("WIND", ExposureFrom, DirAngle, Cpw, Cpl, Case, e1, e2, UserZ, TopZ, BottomZ, q, GustFactor, TopographicFactor, ImportanceFactor, TerrainType, CeWindward, CeLeeward, UserExposure)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 19.0.0.

## See Also

[Set NBCC2015](SetNBCC2015_%7BWind_Load%7D.md)
