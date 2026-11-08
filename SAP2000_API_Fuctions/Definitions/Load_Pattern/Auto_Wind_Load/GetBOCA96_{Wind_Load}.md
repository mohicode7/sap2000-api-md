# GetBOCA96

## Syntax

SapObject.SapModel.LoadPatterns.AutoWind.GetBOCA96

## VB6 Procedure

Function GetBOCA96(ByVal Name As String, ByRef ExposureFrom As Long, ByRef DirAngle As Double, ByRef Cpw As Double, ByRef Cpl As Double, ByRef UserZ As Boolean, ByRef TopZ As Double, ByRef BottomZ As Double, ByRef WindSpeed As Double, ByRef ExposureType As Long, ByRef ImportanceFactor As Double, ByRef UserGust As Boolean, ByRef GustFactor As Double, ByRef UserExposure As Boolean) As Long

## Parameters

Name

The name of an existing Wind-type load pattern with an BOCA 96 auto wind assignment.

ExposureFrom

This is either 1 or 2, indicating the source of the wind exposure.

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

The wind speed in miles per hour.

ExposureType

This is 1, 2, 3 or 4, indicating the exposure category.

1 = A

2 = B

3 = C

4 = D

ImportanceFactor

The importance factor.

UserGust

If this item is True, the gust factor is user defined. If it is False, the gust factor is determined from the code specified values.

GustFactor

The user defined gust factor. This item applies only when UserGust is True.

UserExposure

If this item is True, the wind exposure widths are provided by the user. If it is False, the wind exposure widths are calculated by the program from the extents of the diaphragms.

## Remarks

This function retrieves auto wind loading parameters for BOCA 96.

The function returns zero if the parameters are successfully assigned; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetWindBOCA96()

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

      Dim ExposureType As Long

      Dim ImportanceFactor As Double

      Dim UserGust As Boolean

      Dim GustFactor As Double

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



   'assign BOCA96 parameters

      ret = SapModel.LoadPatterns.AutoWind.SetBOCA96("WIND", 1, 0, 0.8, 0.5, False, 0, 0, 80, 3, 1, 0.85, True, False)



   'get BOCA96 parameters

      ret = SapModel.LoadPatterns.AutoWind.GetBOCA96("WIND", ExposureFrom, DirAngle, Cpw, Cpl, UserZ, TopZ, BottomZ, WindSpeed, ExposureType, ImportanceFactor, UserGust, GustFactor, UserExposure)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.01.

Changed nomenclature from Load Cases, Analysis Cases and Response Combinations to Load Patterns, Load Cases and Load Combinations, respectively, in version 12.00.

## See Also

[SetBOCA96](SetBOCA96_%7BWind_Load%7D.md)
