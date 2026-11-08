# GetBOCA96

## Syntax

SapObject.SapModel.LoadPatterns.AutoSeismic.GetBOCA96

## VB6 Procedure

Function GetBOCA96(ByVal Name As String, ByRef DirFlag As Long, ByRef Eccen As Double, ByRef PeriodFlag As Long, ByRef CT As Double, ByRef UserT As Double, ByRef UserZ As Boolean, ByRef TopZ As Double, ByRef BottomZ As Double, ByRef BOCA96Aa As Double, ByRef BOCA96Av As Double, ByRef BOCA96S As Double, ByRef BOCA96R As Double) As Long

## Parameters

Name

The name of an existing Quake-type load pattern with a BOCA96 auto seismic load assignment.

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

BOCA96Aa

The effective peak acceleration coefficient.

BOCA96Av

The effective peak velocity-related coefficient.

BOCA96S

This is 1, 1.2, 1.5 or 2, indicating the site coefficient.

BOCA96R

The response modification factor.

## Remarks

This function retrieves auto seismic loading parameters for the 1996 BOCA code.

The function returns zero if the parameters are successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetSeismicParametersBOCA96()

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

      Dim BOCA96Aa As Double

      Dim BOCA96Av As Double

      Dim BOCA96S As Double

      Dim BOCA96R As Double



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



   'assign BOCA96 parameters

      ret = SapModel.LoadPatterns.AutoSeismic.SetBOCA96("EQX", 1, 0.05, 1, 0.035, 0, False, 0, 0, 0.4, 0.4, 1.5, 8)



   'get BOCA96 parameters

      ret = SapModel.LoadPatterns.AutoSeismic.GetBOCA96("EQX", DirFlag, Eccen, PeriodFlag, CT, UserT, UserZ, TopZ, BottomZ, BOCA96Aa, BOCA96Av, BOCA96S, BOCA96R)



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

[SetBOCA96](SetBOCA96_%7BAuto_Seismic%7D.md)
