# GetASCE705

## Syntax

SapObject.SapModel.LoadPatterns.AutoWind.GetASCE705

## VB6 Procedure

Function GetASCE705(ByVal Name As String, ByRef ExposureFrom
As Long, ByRef DirAngle As Double, ByRef Cpw As Double, ByRef Cpl As Double,
ByRef ASCECase As Long, ByRef ASCEe1 As Double, ByRef ASCEe2 As Double,
ByRef UserZ As Boolean, ByRef TopZ As Double, ByRef BottomZ As Double,
ByRef WindSpeed As Double, ByRef ExposureType As Long, ByRef ImportanceFactor
As Double, ByRef Kzt As Double, ByRef GustFactor As Double, ByRef Kd As
Double, ByRef SolidGrossRatio As Double, ByRef UserExposure As Boolean)
As Long

## Parameters

Name

The name of an existing Wind-type load pattern with
an ASCE7-05 auto wind assignment.

ExposureFrom

This is 1, 3 or 4, indicating the source of the wind
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
from ASCE7-05 Figure 6-9. 1, 2, 3 and 4 refer to cases 1 through 4 in
the figure. 5 means to create all cases. This item applies only when ExposureFrom
= 1.

ASCEe1

This is the value e1 in ASCE7-05 Figure 6-9. This item
applies only when ExposureFrom = 1.

ASCEe2

This is the value e2 in ASCE7-05 Figure 6-9. This item
applies only when ExposureFrom = 1.

UserZ

This item is True if the top and bottom elevations of
the wind load are user specified. It is False if the elevations are determined
by the program.

TopZ

This item is the global Z-coordinate at the highest
level where auto wind loads are applied. [L]

BottomZ

This item is the global Z-coordinate at the lowest level
where auto wind loads are applied. [L]

WindSpeed

The wind speed in miles per hour.

ExposureType

This is 1, 2 or 3, indicating the exposure category.

1 = B

2 = C

3 = D

ImportanceFactor

The importance factor.

Kzt

The topographical factor.

GustFactor

The gust factor.

Kd

The directionality factor.

SolidGrossRatio

The solid area divided by the gross area ratio for open
frame structure loading. This item applies only when the loading is from
open structure wind loading (ExposureFrom = 3 or ExposureFrom = 4).

UserExposure

If this item is True, the wind exposure widths are provided
by the user. If it is False, the wind exposure widths are calculated by
the program from the extents of the diaphragms.

## Remarks

This function retrieves auto wind loading parameters
for ASCE 7-05.

The function returns zero if the parameters are successfully
assigned; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetWindASCE705()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim ExposureFrom As Long

      Dim DirAngle As Double

      Dim Cpw As Double

      Dim Cpl As Double

      Dim ASCECase As Long

      Dim ASCEe1 As Double

      Dim ASCEe2 As Double

      Dim UserZ As Boolean

      Dim TopZ As Double

      Dim BottomZ As Double

      Dim WindSpeed As Double

      Dim ExposureType As Long

      Dim ImportanceFactor As Double

      Dim Kzt As Double

      Dim GustFactor As Double

      Dim Kd As Double

      Dim SolidGrossRatio As Double

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



   'assign ASCE705 parameters

      ret = SapModel.LoadPatterns.AutoWind.SetASCE705("WIND", 
 1, 0, 0.8, 0.5, 2, 0.15, 0.18, False, 0, 0, 80, 3, 1, 1.1, 0.85, 0.88)



   'get ASCE705 parameters

      ret = SapModel.LoadPatterns.AutoWind.GetASCE705("WIND", 
 ExposureFrom, DirAngle, Cpw, Cpl, ASCECase, ASCEe1, ASCEe2, UserZ, TopZ, 
 BottomZ, WindSpeed, ExposureType, ImportanceFactor, Kzt, GustFactor, Kd, 
 SolidGrossRatio, UserExposure)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.01.

Changed nomenclature from Load Cases, Analysis Cases
and Response Combinations to Load Patterns, Load Cases and Load Combinations,
respectively, in version 12.00.

## See Also

[SetASCE705](GetASCE705.md)
