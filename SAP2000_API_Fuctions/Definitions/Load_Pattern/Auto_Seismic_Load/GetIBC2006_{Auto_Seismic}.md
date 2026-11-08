# GetIBC2006

## Syntax

SapObject.SapModel.LoadPatterns.AutoSeismic.GetIBC2006

## VB6 Procedure

Function GetIBC2006(ByVal Name As String, ByRef DirFlag As Long, ByRef Eccen As Double, ByRef PeriodFlag As Long, ByRef IBC2006CtType As Long, ByRef UserT As Double, ByRef UserZ As Boolean, ByRef TopZ As Double, ByRef BottomZ As Double, ByRef IBC2006R As Double, ByRef IBC2006Omega As Double, ByRef IBC2006Cd As Double, ByRef IBC2006I As Double, ByRef IBC2006Option As Long, ByRef IBC2006Latitude As Double, ByRef IBC2006Longitude As Double, ByRef IBC2006ZipCode As String, ByRef IBC2006SS As Double, ByRef IBC2006S1 As Double, ByRef IBC2006TL As Double, ByRef IBC2006SiteClass As Long, ByRef IBC2006Fa As Double, ByRef IBC2006Fv As Double) As Long

## Parameters

Name

The name of an existing Quake-type load pattern with a IBC2006 auto seismic load assignment.

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

IBC2006CtType

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

IBC2006R

The response modification factor.

IBC2006Omega

The system overstrength factor.

IBC2006Cd

The deflection amplification factor.

IBC2006I

The occupancy importance factor.

IBC2006Option

This is 0, 1 or 2, indicating the seismic coefficient option.

0 = Ss and S1 from USGS by latitude and longitude

1 = Ss and S1 from USGS by zip code

2 = Ss and S1 are user defined

IBC2006Latitude, IBC2006Longitude

The latitude and longitude for which the seismic coefficients are obtained. These items are meaningful only when IBC2006Option = 0 or 1.

IBC2006ZipCode

The zip code for which the seismic coefficients are obtained. This item is meaningful only when IBC2006Option = 1.

IBC2006SS, IBC2006S1

The seismic coefficients Ss and S1.

IBC2006TL

The long-period transition period. [s]

IBC2006SiteClass

This is 1, 2, 3, 4, 5 or 6, indicating the site class.

1 = A

2 = B

3 = C

4 = D

5 = E

6 = F

IBC2006Fa, IBC2006Fv

The site coefficients Fa and Fv.

## Remarks

This function retrieves auto seismic loading parameters for the 2006 IBC code.

The function returns zero if the parameters are successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetSeismicParametersIBC2006()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim DirFlag As Long

      Dim Eccen As Double

      Dim PeriodFlag As Long

      Dim IBC2006CtType As Long

      Dim UserT As Double

      Dim UserZ As Boolean

      Dim TopZ As Double

      Dim BottomZ As Double

      Dim IBC2006R As Double

      Dim IBC2006Omega As Double

      Dim IBC2006Cd As Double

      Dim IBC2006I As Double

      Dim IBC2006Option As Long

      Dim IBC2006Latitude As Double

      Dim IBC2006Longitude As Double

      Dim IBC2006ZipCode As String

      Dim IBC2006SS As Double

      Dim IBC2006S1 As Double

      Dim IBC2006TL As Double

      Dim IBC2006SiteClass As Long

      Dim IBC2006Fa As Double

      Dim IBC2006Fv As Double



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



   'assign IBC2006 parameters

      ret = SapModel.LoadPatterns.AutoSeismic.SetIBC2006("EQX", 1, 0.05, 1, 1, 0, False, 0, 0, 8, 3, 5.5, 1, 1, 0, 0, "94704", 0, 0, 8, 3, 0, 0)



   'get IBC2006 parameters

      ret = SapModel.LoadPatterns.AutoSeismic.GetIBC2006("EQX", DirFlag, Eccen, PeriodFlag, IBC2006CtType, UserT, UserZ, TopZ, BottomZ, IBC2006R, IBC2006Omega, IBC2006Cd, IBC2006I, IBC2006Option, IBC2006Latitude, IBC2006Longitude, IBC2006ZipCode, IBC2006SS, IBC2006S1, IBC2006TL, IBC2006SiteClass, IBC2006Fa, IBC2006Fv)



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

[SetIBC2006](SetIBC2006_%7BAuto_Seismic%7D.md)
