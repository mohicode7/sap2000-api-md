# GetNEHRP97

## Syntax

SapObject.SapModel.LoadPatterns.AutoSeismic.GetNEHRP97

## VB6 Procedure

Function GetNEHRP97(ByVal Name As String, ByRef DirFlag As Long, ByRef Eccen As Double, ByRef PeriodFlag As Long, ByRef CT As Double, ByRef UserT As Double, ByRef UserZ As Boolean, ByRef TopZ As Double, ByRef BottomZ As Double, ByRef NEHRP97SG As Long, ByRef NEHRP97SeismicCoeffFlag As Long, ByRef NEHRP97Site As Long, ByRef NEHRP97SS As Double, ByRef NEHRP97S1 As Double, ByRef NEHRP97Fa As Double, ByRef NEHRP97Fv As Double, ByRef NEHRP97R As Double) As Long

## Parameters

Name

The name of an existing Quake-type load pattern with a NEHRP97 auto seismic load assignment.

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

CT

The code-specified CT factor. This item applies when the PeriodFlag item is 1 or 2.

UserT

The user specified time period. This item applies when the PeriodFlag item is 3. [s]

UserZ

This item is True if the top and bottom elevations of the seismic load are user specified. It is False if the elevations are determined by the program.

TopZ

This item applies only when the UserZ item is True. It is the global Z-coordinate at the highest level where auto seismic loads are applied. [L]

BottomZ

This item applies only when the UserZ item is True. It is the global Z-coordinate at the lowest level where auto seismic loads are applied. [L]

NEHRP97SG

This is 1, 2 or 3, indicating the seismic group.

1 = I

2 = II

3 = III

NEHRP97SeismicCoeffFlag

This is 1 or 2, indicating the seismic coefficient option.

1 = Coefficients are per code

2 = Coefficients are user defined

NEHRP97Site

This is 1, 2, 3, 4 or 5, indicating the site class. This item is only filled when the NEHRP97SeismicCoeffFlag = 1.

1 = A

2 = B

3 = C

4 = D

5 = E

NEHRP97SS

The response acceleration for short periods, (g).

NEHRP97S1

The response acceleration for a one second period, (g).

NEHRP97Fa

The site coefficient Fa.

NEHRP97Fv

The site coefficient Fv.

NEHRP97R

The response modification factor.

## Remarks

This function retrieves auto seismic loading parameters for the 1997 NEHRP code.

The function returns zero if the parameters are successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetSeismicParametersNEHRP97()

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

      Dim NEHRP97SG As Long

      Dim NEHRP97SeismicCoeffFlag As Long

      Dim NEHRP97Site As Long

      Dim NEHRP97SS As Double

      Dim NEHRP97S1 As Double

      Dim NEHRP97Fa As Double

      Dim NEHRP97Fv As Double

      Dim NEHRP97R As Double



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



   'assign NEHRP97 parameters

      ret = SapModel.LoadPatterns.AutoSeismic.SetNEHRP97("EQX", 1, 0.05, 1, 0.035, 0, False, 0, 0, 1, 1, 3, 1, 0.4, 0, 0, 8)



   'get NEHRP97parameters

      ret = SapModel.LoadPatterns.AutoSeismic.GetNEHRP97("EQX", DirFlag, Eccen, PeriodFlag, CT, UserT, UserZ, TopZ, BottomZ, NEHRP97SG, NEHRP97SeismicCoeffFlag, NEHRP97Site, NEHRP97SS, NEHRP97S1, NEHRP97Fa, NEHRP97Fv, NEHRP97R)



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

[SetNEHRP97](SetNEHRP97_%7BAuto_Seismic%7D.md)
