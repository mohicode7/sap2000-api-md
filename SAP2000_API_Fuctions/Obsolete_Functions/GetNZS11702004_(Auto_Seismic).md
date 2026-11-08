# GetNZS11702004

## Syntax

SapObject.SapModel.LoadPatterns.AutoSeismic.GetNZS11702004

## VB6 Procedure

Function GetNZS11702004(ByVal Name As String, ByRef DirFlag As Long, ByRef Eccen As Double, ByRef PeriodFlag As Long, ByRef UserT As Double, ByRef UserZ As Boolean, ByRef TopZ As Double, ByRef BottomZ As Double, ByRef NZS2004SiteClass As Long, ByRef NZS2004Z As Double, ByRef NZS2004R As Double, ByRef NZS2004DIST As Double, ByRef NZS2004Sp As Double, ByRef NZS2004Mu As Double) As Long

## Parameters

Name

The name of an existing Quake-type load pattern with a NZS 1170 2004 auto seismic load assignment.

DirFlag

This is 1 or 2, indicating the seismic load direction.

1 = Global X

2 = Global Y

Eccen

The eccentricity ratio that applies to all diaphragms.

PeriodFlag

This is 1, 2 or 3, indicating the time period option.

1 = Approximate

2 = Program calculated

3 = User defined

UserT

The user specified time period. This item applies when the PeriodFlag item is 3. [s]

UserZ

This item is True if the top and bottom elevations of the seismic load are user specified. It is False if the elevations are determined by the program.

TopZ

This item applies only when the UserZ item is True. It is the global Z-coordinate at the highest level where auto seismic loads are applied. [L]

BottomZ

This item applies only when the UserZ item is True. It is the global Z-coordinate at the lowest level where auto seismic loads are applied. [L]

NZS2004SiteClass

This is 1, 2, 3, 4 or 5, indicating the site class.

1 = A

2 = B

3 = C

4 = D

5 = E

NZS2004Z

The hazard factor, Z.

NZS2004R

The return period factor, R.

NZS2004DIST

Distance to the fault in km, used to calculate the near fault factor..

NZS2004Sp

The structural performance factor, Sp.

NZS2004Mu

The structural ductility factor, u.

## Remarks

This function retrieves auto seismic loading parameters for the NZS 1170 2004 code.

The function returns zero if the parameters are successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetSeismicParametersNZS11702004()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim DirFlag As Long

      Dim Eccen As Double

      Dim PeriodFlag As Long

      Dim UserT As Double

      Dim UserZ As Boolean

      Dim TopZ As Double

      Dim BottomZ As Double

      Dim NZS2004SiteClass As Long

      Dim NZS2004Z As Double

      Dim NZS2004R As Double

      Dim NZS2004DIST As Double

      Dim NZS2004Sp As Double

      Dim NZS2004Mu As Double



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



   'assign NZS 1170 2004 parameters

      ret = SapModel.LoadPatterns.AutoSeismic.SetNZS11702004("EQX", 2, 0.1, 2, 0, False, 0, 0, 3, 0.4, 1.3, 20, 0.7, 3)



   'get NZS 1170 2004 parameters

      ret = SapModel.LoadPatterns.AutoSeismic.GetNZS11702004("EQX", DirFlag, Eccen, PeriodFlag, UserT, UserZ, TopZ, BottomZ, NZS2004SiteClass, NZS2004Z, NZS2004R, NZS2004DIST, NZS2004Sp, NZS2004Mu)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 14.00.

Modified NZS2004N TO NZS2004DIST in version 14.1.0.

## See Also

[SetNZS11702004](SetNZS11702004_%28Auto_Seismic%29.md)
