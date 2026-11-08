# SetIBC2012

## Syntax

SapObject.SapModel.LoadPatterns.AutoSeismic.SetIBC2012

## VB6 Procedure

Function SetIBC2012(ByVal Name As String, ByVal DirFlag As Long, ByVal Eccen As Double, ByVal PeriodFlag As Long, ByVal CtType As Long, ByVal UserT As Double, ByVal UserZ As Boolean, ByVal TopZ As Double, ByVal BottomZ As Double, ByVal R As Double, ByVal Omega As Double, ByVal Cd As Double, ByVal I As Double, ByVal IBC2012Option As Long, ByVal Latitude As Double, ByVal Longitude As Double, ByVal ZipCode As String, ByVal SS As Double, ByVal S1 As Double, ByVal TL As Double, ByVal SiteClass As Long, ByVal Fa As Double, ByVal Fv As Double) As Long

## Parameters

Name

The name of an existing Quake-type load pattern.

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

This is 0, 1, 2 or 3, indicating the values of Ct and x. This item applies when the PeriodFlag item is 1 or 2.

0 = Ct = 0.028 (ft),     x = 0.8

1 = Ct = 0.016 (ft),     x = 0.9

2 = Ct = 0.03 (ft),      x = 0.75

3 = Ct = 0.02 (ft),      x = 0.75

UserT

The user specified time period. This item applies when the PeriodFlag item is 3. [s]

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

IBC2012Option

This is 0, 1 or 2, indicating the seismic coefficient option.

0 = Ss and S1 from USGS by latitude and longitude

1 = Ss and S1 from USGS by zip code

2 = Ss and S1 are user defined

Latitude, Longitude

The latitude and longitude for which the seismic coefficients are obtained. These items are used only when IBC2012Option = 0.

ZipCode

The zip code for which the seismic coefficients are obtained. This item is used only when IBC2012Option = 1.

SS, S1

The seismic coefficients Ss and S1. This item is used only when IBC2012Option = 2.

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

The site coefficients Fa and Fv. These items are used only when IBC2012SiteClass= 6.

## Remarks

This function assigns auto seismic loading parameters for the 2012 IBC code.

The function returns zero if the parameters are successfully assigned; otherwise it returns a nonzero value.

## VBA Example



```
Sub AssignSeismicIBC2012()

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

      ret = SapModel.File.New3DFrame(BeamSlab, 2, 144, 3, 336, 2, 432)



   'add new load pattern

      ret = SapModel.LoadPatterns.Add("EQX", LTYPE_QUAKE)



   'assign IBC2012 parameters

      ret = SapModel.LoadPatterns.AutoSeismic.SetIBC2012("EQX", 1, 0.05, 1, 1, 0, False, 0, 0, 8, 3, 5.5, 1, 1, 0, 0, "94704", 0, 0, 8, 3, 0, 0)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 16.0.0.

## See Also

[GetIBC2012](GetIBC2012.md)
