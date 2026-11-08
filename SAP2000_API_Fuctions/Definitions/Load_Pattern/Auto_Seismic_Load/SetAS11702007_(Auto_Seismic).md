# SetAS11702007

## Syntax

SapObject.SapModel.LoadPatterns.AutoSeismic.SetAS11702007

## VB6 Procedure

Function SetAS11702007(ByVal Name As String, ByVal DirFlag As Long, ByVal Eccen As Double, ByVal PeriodFlag As Long, ByVal CT As Double, ByVal UserT As Double, ByVal UserZ As Boolean, ByVal TopZ As Double, ByVal BottomZ As Double, ByVal AS2007SiteClass As Long, ByVal AS2007kp As Double, ByVal AS2007Z As Double, ByVal AS2007Sp As Double, ByVal AS2007Mu As Double) As Long

## Parameters

Name

The name of an existing Quake-type load pattern.

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

CT

The code-specified kt factor. This item applies when the PeriodFlag item is 1.

UserT

The user specified time period. This item applies when the PeriodFlag item is 3. [s]

UserZ

This item is True if the top and bottom elevations of the seismic load are user specified. It is False if the elevations are determined by the program.

TopZ

This item applies only when the UserZ item is True. It is the global Z-coordinate at the highest level where auto seismic loads are applied. [L]

BottomZ

This item applies only when the UserZ item is True. It is the global Z-coordinate at the lowest level where auto seismic loads are applied. [L]

AS2007SiteClass

This is 1, 2, 3, 4 or 5, indicating the site class.

1 = A

2 = B

3 = C

4 = D

5 = E

AS2007kp

The probability factor, kp.

AS2007Z

The hazard factor, Z.

AS2007Sp

The structural performance factor, Sp.

AS2007Mu

The structural ductility factor, u.

## Remarks

This function assigns auto seismic loading parameters for the AS 1170.4 2007 code.

The function returns zero if the parameters are successfully assigned; otherwise it returns a nonzero value.

## VBA Example



```
Sub AssignSeismicAS11702007()

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



   'assign AS 1170 2007 parameters

      ret = SapModel.LoadPatterns.AutoSeismic.SetAS11702007("EQX", 2, 0.1, 2, 0.075, 0, False, 0, 0, 3, 1.3, 0.09, 0.77, 2)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 14.00.

## See Also

[GetAS11702007](GetAS11702007_%28Auto_Seismic%29.md)
