# SetASCE716

## Syntax

SapObject.SapModel.LoadPatterns.AutoWind.SetASCE716

## VB6 Procedure

Function SetASCE716(ByVal Name As String, ByVal ExposureFrom
As Long, ByVal DirAngle As Double, ByVal Cpw As Double, ByVal Cpl As Double,
ByVal ASCECase As Long, ByVal ASCEe1 As Double, ByVal ASCEe2 As Double,
ByVal UserZ As Boolean, ByVal TopZ As Double, ByVal BottomZ As Double,
ByVal WindSpeed As Double, ByVal ExposureType As Long, ByVal Kzt As Double, ByVal GustFactor
As Double, ByVal Kd As Double, Optional ByVal SolidGrossRatio As Double
= 0.2, Optional ByVal UserExposure As Boolean = False) As Long

## Parameters

Name

The name of an existing Wind-type load pattern.

ExposureFrom

This is 1, 2, 3 or 4, indicating the source of the wind
exposure.

1 = From extents of rigid
diaphragms

2 = From area objects

3 = From frame objects
(open structure)

4 = From area objects and
frame objects (open structure)

DirAngle

The direction angle for the wind load. This item applies
only when ExposureFrom = 1, 3, or 4.

Cpw

The windward coefficient, Cp. This item applies only
when ExposureFrom = 1.

Cpl

The leeward coefficient, Cp. This item applies only
when ExposureFrom = 1.

ASCECase

This is 1, 2, 3, 4 or 5, indicating the desired case
from ASCE7-16 Figure 27.3-8 . 1, 2, 3 and 4 refer to cases 1 through 4
in the figure. 5 means to create all cases. This item applies only when
ExposureFrom = 1.

ASCEe1

This is the value e1 in ASCE7-16 Figure 27.3-8. This
item applies only when ExposureFrom = 1.

ASCEe2

This is the value e2 in ASCE7-16 Figure 27.3-8. This
item applies only when ExposureFrom = 1.

UserZ

This item is True if the top and bottom elevations of
the wind load are user specified. It is False if the elevations are determined
by the program.

TopZ

This item applies only when the UserZ item is True.
It is the global Z-coordinate at the highest level where auto wind loads
are applied. [L]

BottomZ

This item applies only when the UserZ item is True.
It is the global Z-coordinate at the lowest level where auto wind loads
are applied. [L]

WindSpeed

The wind speed in miles per hour.

ExposureType

This is 1, 2 or 3, indicating the exposure category.

1 = B

2 = C

3 = D

Kzt

The topographical factor.

GustFactor

The gust factor.

Kd

The directionality factor.

SolidGrossRatio

The solid area divided by gross area ratio for open
frame structure loading. This item applies only when the loading is from
open structure wind loading (ExposureFrom = 3 or ExposureFrom = 4).

UserExposure

If this item is True, the wind exposure widths are provided
by the user. If it is False, the wind exposure widths are calculated by
the program from the extents of the diaphragms.

## Remarks

This function assigns auto wind loading parameters for
ASCE 7-16.

The function returns zero if the parameters are successfully
assigned; otherwise it returns a nonzero value.

## VBA Example



```
Sub AssignWindASCE716()

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

      ret = SapModel.File.New3DFrame(BeamSlab, 
 2, 144, 3, 336, 2, 432)



   'define diaphragm constraints

      ret = SapModel.ConstraintDef.SetDiaphragm("Diaph1", 
 Z)

      ret = SapModel.ConstraintDef.SetDiaphragm("Diaph2", 
 Z)



   'assign points to diaphragm

      ret = SapModel.SelectObj.ClearSelection

      ret = SapModel.SelectObj.PlaneXY("2")

      ret = SapModel.PointObj.SetConstraint("", 
 "Diaph1", SelectedObjects)

      ret = SapModel.SelectObj.ClearSelection

      ret = SapModel.SelectObj.PlaneXY("3")

      ret = SapModel.PointObj.SetConstraint("", 
 "Diaph2", SelectedObjects)

      ret = SapModel.SelectObj.ClearSelection



   'add new load pattern

      ret = SapModel.LoadPatterns.Add("WIND", 
 LTYPE_WIND)



   'assign ASCE716 parameters

      ret = SapModel.LoadPatterns.AutoWind.SetASCE716("WIND", 
 1, 0, 0.8, 0.5, 2, 0.15, 0.18, False, 0, 0, 80, 3, 1.1, 0.85, 0.88)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 23.2.0

## See Also

[GetASCE716](GetASCE716.md)
