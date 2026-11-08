# GetAS11702007

## Syntax

SapObject.SapModel.LoadPatterns.AutoSeismic.GetAS11702007

## VB6 Procedure

Function GetAS11702007(ByVal Name As String, ByRef DirFlag As Long, ByRef Eccen As Double, ByRef PeriodFlag As Long, ByRef CT As Double, ByRef UserT As Double, ByRef UserZ As Boolean, ByRef TopZ As Double, ByRef BottomZ As Double, ByRef AS2007SiteClass As Long, ByRef AS2007kp As Double, ByRef AS2007Z As Double, ByRef AS2007Sp As Double, ByRef AS2007Mu As Double) As Long

## Parameters

Name

The name of an existing Quake-type load pattern with a AS 1170 2007 auto seismic load assignment.

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

This function retrieves auto seismic loading parameters for the AS 1170 2007 code.

The function returns zero if the parameters are successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetSeismicParametersAS11702007()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim DirFlag As Long

      Dim Eccen As Double

      Dim PeriodFlag As Long

      Dim CT As Double

      Dim UserT As Double

      Dim UserZ As Boolean

      Dim TopZ As Double

      Dim BottomZ As Double

      Dim AS2007SiteClass As Long

      Dim AS2007kp As Double

      Dim AS2007Z As Double

      Dim AS2007Sp As Double

      Dim AS2007Mu As Double



   'create Sap2000 object

      Set SapObject = New Sap2000v16.SapObject



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



   'get AS 1170 2007 parameters

      ret = SapModel.LoadPatterns.AutoSeismic.GetAS11702007("EQX", DirFlag, Eccen, PeriodFlag, CT, UserT, UserZ, TopZ, BottomZ, AS2007SiteClass, AS2007kp, AS2007Z, AS2007Sp, AS2007Mu)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 14.00.

## See Also

[SetAS11702007](SetAS11702007_%28Auto_Seismic%29.md)
