# GetIBC2009

## Syntax

SapObject.SapModel.LoadPatterns.AutoSeismic.GetIBC2009

## VB6 Procedure

Function GetIBC2009(ByVal Name As String, ByRef DirFlag As Long, ByRef Eccen As Double, ByRef PeriodFlag As Long, ByRef CtType As Long, ByRef UserT As Double, ByRef UserZ As Boolean, ByRef TopZ As Double, ByRef BottomZ As Double, ByRef R As Double, ByRef Omega As Double, ByRef Cd As Double, ByRef I As Double, ByRef IBC2009Option As Long, ByRef Latitude As Double, ByRef Longitude As Double, ByRef ZipCode As String, ByRef SS As Double, ByRef S1 As Double, ByRef TL As Double, ByRef SiteClass As Long, ByRef Fa As Double, ByRef Fv As Double) As Long

## Parameters

Name

The name of an existing Quake-type load pattern with a IBC2009 auto seismic load assignment.

DirFlag

This is either 1 or 2, indicating the seismic load direction.

1 = Global X

2 = Global Y

Eccen

The eccentricity ratio that applies to all diaphragms.

PeriodFlag

This is 1, 2 or 3, indicating the time period option.

1 = Approximate

2 = Program calculated

3 = User defined

CtType

This is 0, 1, 2 or 3, indicating the values of Ct and x. This item is meaningful when the PeriodFlag item is 1 or 2.

0 = Ct = 0.028 (ft),     x = 0.8

1 = Ct = 0.016 (ft),     x = 0.9

2 = Ct = 0.03 (ft),      x = 0.75

3 = Ct = 0.02 (ft),      x = 0.75

UserT

The user specified time period. This item is meaningful when the PeriodFlag item is 3. [s]

UserZ

This item is True if the top and bottom elevations of the seismic load are user specified. It is False if the elevations are determined by the program.

TopZ

This item applies only when the UserZ item is True. It is the global Z-coordinate at the highest level where auto seismic loads are applied. [L]

BottomZ

This item applies only when the UserZ item is True. It is the global Z-coordinate at the lowest level where auto seismic loads are applied. [L]

R

The response modification factor.

Omega

The system overstrength factor.

Cd

The deflection amplification factor.

I

The occupancy importance factor.

IBC2009Option

This is 0, 1 or 2, indicating the seismic coefficient option.

0 = Ss and S1 from USGS by latitude and longitude

1 = Ss and S1 from USGS by zip code

2 = Ss and S1 are user defined

Latitude, Longitude

The latitude and longitude for which the seismic coefficients are obtained. These items are meaningful only when IBC2009Option = 0 or 1.

ZipCode

The zip code for which the seismic coefficients are obtained. This item is meaningful only when IBC2009Option = 1.

SS, S1

The seismic coefficients Ss and S1.

TL

The long-period transition period. [s]

SiteClass

This is 1, 2, 3, 4, 5 or 6, indicating the site class.

1 = A

2 = B

3 = C

4 = D

5 = E

6 = F

Fa, Fv

The site coefficients Fa and Fv.

## Remarks

This function retrieves auto seismic loading parameters for the 2009 IBC code.

The function returns zero if the parameters are successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetSeismicParametersIBC2009()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim DirFlag As Long

      Dim Eccen As Double

      Dim PeriodFlag As Long

      Dim CtType As Long

      Dim UserT As Double

      Dim UserZ As Boolean

      Dim TopZ As Double

      Dim BottomZ As Double

      Dim R As Double

      Dim Omega As Double

      Dim Cd As Double

      Dim I As Double

      Dim IBC2009Option As Long

      Dim Latitude As Double

      Dim Longitude As Double

      Dim ZipCode As String

      Dim SS As Double

      Dim S1 As Double

      Dim TL As Double

      Dim SiteClass As Long

      Dim Fa As Double

      Dim Fv As Double



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



   'add new load pattern

      ret = SapModel.LoadPatterns.Add("EQX", LTYPE_QUAKE)



   'assign IBC2009 parameters

      ret = SapModel.LoadPatterns.AutoSeismic.SetIBC2009("EQX", 1, 0.05, 1, 1, 0, False, 0, 0, 8, 3, 5.5, 1, 1, 0, 0, "94704", 0, 0, 8, 3, 0, 0)



   'get IBC2009 parameters

      ret = SapModel.LoadPatterns.AutoSeismic.GetIBC2009("EQX", DirFlag, Eccen, PeriodFlag, CtType, UserT, UserZ, TopZ, BottomZ, R, Omega, Cd, I, IBC2009Option, Latitude, Longitude, ZipCode, SS, S1, TL, SiteClass, Fa, Fv)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 16.0.0.

## See Also

[SetIBC2009](SetIBC2009.md)
