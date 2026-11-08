# SetNBCC2005

## Syntax

SapObject.SapModel.LoadPatterns.AutoSeismic.SetNBCC2005

## VB6 Procedure

Function SetNBCC2005(ByVal Name As String, ByVal DirFlag As Long, ByVal eccen As Double, ByVal NBCC2005CtType As Long, ByVal PeriodFlag As Long, ByVal UserT As Double, ByVal UserZ As Boolean, ByVal TopZ As Double, ByVal BottomZ As Double, ByVal NBCC2005PGA As Double, ByVal NBCC2005S02 As Double, ByVal NBCC2005S05 As Double, ByVal NBCC2005S1 As Double, ByVal NBCC2005S2 As Double, ByVal NBCC2005SiteClass As Long, ByVal NBCC2005Fa As Double, ByVal NBCC2005Fv As Double, ByVal NBCC2005I As Double, ByVal NBCC2005Mv As Double, ByVal NBCC2005Rd As Double, ByVal NBCC2005Ro As Double) As Long

## Parameters

Name

The name of an existing Quake-type load pattern.

DirFlag

This is 1 or 2, indicating the seismic load direction.

1 = Global X

2 = Global Y

Eccen

The eccentricity ratio that applies to all diaphragms.

NBCC2005CtType

This is 0, 1, 2, 3 or 4, indicating the structure type.

0 = Steel moment frame

1 = Concrete moment frame

2 = Other moment frame

3 = Braced frame

4 = Shear wall

NBCC95DS

This item applies only when the NBCCPFlag = 2. It is the dimension of the lateral load resisting system in the direction of the applied forces. [L]

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

NBCC2005PGA

The peak ground acceleration.

NBCC2005S02

The spectral acceleration at a 0.2 second period.

NBCC2005S05

The spectral acceleration at a 0.52 second period.

NBCC2005S1

The spectral acceleration at a 1 second period.

NBCC2005S2

The spectral acceleration at a 2 second period.

NBCC2005SiteClass

This is 1, 2, 3, 4, 5 or 6, indicating the site class.

1 = A

2 = B

3 = C

4 = D

5 = E

6 = F

NBCC2005Fa

The site coefficient, Fa. This item is read when the site class is F only.

NBCC2005Fv

The site coefficient, Fv. This item is read when the site class is F only.

NBCC2005I

The importance factor.

NBCC2005Mv

The higher mode factor.

NBCC2005Rd

The ductility modifier.

NBCC2005Ro

The overstrength modifier.

## Remarks

This function assigns auto seismic loading parameters for the 2005 NBCC code.

The function returns zero if the parameters are successfully assigned; otherwise it returns a nonzero value.

## VBA Example



```
Sub AssignSeismicNBCC2005()

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



   'assign NBCC2005 parameters

      ret = SapModel.LoadPatterns.AutoSeismic.SetNBCC2005("EQX", 2, 0.1, 2, 1, 0, False, 0, 0, 0.6, 1.1, 0.7, 0.35, 0.2, 6, 1.8, 2, 1.5, 1.2, 6, 1.6)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 12.00.

## See Also

[GetNBCC2005](GetNBCC2005_%7BAuto_Seismic%7D.md)
