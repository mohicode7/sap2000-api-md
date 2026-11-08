# SetNBCC2010

## Syntax

SapObject.SapModel.LoadPatterns.AutoSeismic.SetNBCC2010

## VB6 Procedure

Function SetNBCC2010(ByVal Name As String, ByVal DirFlag As Long, ByVal Eccen As Double, ByVal CtType As Long, ByVal PeriodFlag As Long, ByVal UserT As Double, ByVal UserZ As Boolean, ByVal TopZ As Double, ByVal BottomZ As Double, ByVal PGA As Double, ByVal S02 As Double, ByVal S05 As Double, ByVal S1 As Double, ByVal S2 As Double, ByVal S4 As Double, ByVal SiteClass As Long, ByVal Fa As Double, ByVal Fv As Double, ByVal I As Double, ByVal Mv As Double, ByVal Rd As Double, ByVal Ro As Double) As Long

## Parameters

Name

The name of an existing Quake-type load pattern with a NBCC2010 auto seismic load assignment.

DirFlag

This is 1 or 2, indicating the seismic load direction.

1 = Global X

2 = Global Y

Eccen

The eccentricity ratio that applies to all diaphragms.

CtType

This is 0, 1, 2, 3 or 4, indicating the structure type.

0 = Steel moment frame

1 = Concrete moment frame

2 = Other moment frame

3 = Braced frame

4 = Shear wall

PeriodFlag

This is 1, 2 or 3, indicating the time period option.

1 = Code

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

PGA

The peak ground acceleration.

S02

The spectral acceleration at a 0.2 second period.

S05

The spectral acceleration at a 0.5 second period.

S1

The spectral acceleration at a 1 second period.

S2

The spectral acceleration at a 2 second period.

S4

The spectral acceleration at a 4 second period.

SiteClass

This is 1, 2, 3, 4, 5 or 6, indicating the site class.

1 = A

2 = B

3 = C

4 = D

5 = E

6 = F

Fa

The site coefficient, Fa. This item is read when the site class is F only.

Fv

The site coefficient, Fv. This item is read when the site class is F only.

I

The importance factor.

Mv

The higher mode factor.

Rd

The ductility modifier.

Ro

The overstrength modifier.

## Remarks

This function retrieves auto seismic loading parameters for the 2010 NBCC code.

The function returns zero if the parameters are successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub AssignSeismicNBCC2010()

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



   'assign NBCC2010 parameters

      ret = SapModel.LoadPatterns.AutoSeismic.SetNBCC2010("EQX", 2, 0.1, 2, 1, 0, False, 0, 0, 0.6, 1.1, 0.7, 0.35, 0.2, 0.1, 6, 1.8, 2, 1.5, 1.2, 6, 1.6)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 19.0.0.

## See Also

[GetNBCC2010](GetNBCC2010_%7BAuto_Seismic%7D.md)
