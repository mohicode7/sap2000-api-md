# GetUBC97

## Syntax

SapObject.SapModel.LoadPatterns.AutoSeismic.GetUBC97

## VB6 Procedure

Function GetUBC97(ByVal Name As String, ByRef DirFlag As Long, ByRef Eccen As Double, ByRef PeriodFlag As Long, ByRef CT As Double, ByRef UserT As Double, ByRef UserZ As Boolean, ByRef TopZ As Double, ByRef BottomZ As Double, ByRef UBC97SeismicCoeffFlag As Long, ByRef UBC97SoilProfileType As Long, ByRef UBC97Z As Double, ByRef UBC97Ca As Double, ByRef UBC97Cv As Double, ByRef UBC97NearSourceFlag As Long, ByRef UBC97SourceType As Long, ByRef UBC97Dist As Double, ByRef UBC97Na As Double, ByRef UBC97Nv As Double, ByRef UBC97I As Double, ByRef UBC97R As Double) As Long

## Parameters

Name

The name of an existing Quake-type load pattern with a UBC97 auto seismic load assignment.

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

UBC97SeismicCoeffFlag

This is either 1 or 2, indicating if the seismic coefficients Ca and Cv are per code or user defined.

1 = Per code

2 = User defined

UBC97SoilProfileType

This is 1, 2, 3, 4 or 5, indicating the soil profile type.

1 = SA

2 = SB

3 = SC

4 = SD

5 = SE

This item is applicable only when the seismic coefficients Ca and Cv are calculated per code (UBC97SeismicCoeffFlag = 1).

UBC97Z

This is 0.075, 0.15, 0.2, 0.3 or 0.4, indicating the seismic zone factor.

This item is applicable only when the seismic coefficients Ca and Cv are calculated in accordance with code (UBC97SeismicCoeffFlag = 1).

UBC97Ca

The seismic coefficient, Ca.

UBC97Cv

The seismic coefficient, Cv.

UBC97NearSourceFlag

This is 1 or 2, indicating if the near source factor coefficients Na and Nv are per code or user defined.

1 = Per code

2 = User defined

This item is applicable only when the seismic coefficients Ca and Cv are calculated per code (UBC97SeismicCoeffFlag = 1) and UBC97Z = 0.4.

UBC97SourceType

This is 1, 2 or 3, indicating the seismic source type.

1 = A

2 = B

3 = C

This item is applicable only when the seismic coefficients Ca and Cv are calculated per code (UBC97SeismicCoeffFlag = 1), UBC97Z = 0.4, and the near source factor coefficients Na and Nv are calculated per code (UBC97NearSourceFlag = 1).

UBC97Dist

This is the distance to the seismic source in kilometers.

This item is only applicable when the seismic coefficients Ca and Cv are calculated per code (UBC97SeismicCoeffFlag = 1), UBC97Z = 0.4, and the near source factor coefficients Na and Nv are calculated per code (UBC97NearSourceFlag = 1).

UBC97Na

The near source factor coefficient, Na.

This item is applicable only when the seismic coefficients Ca and Cv are user defined (UBC97SeismicCoeffFlag = 2).

UBC97Nv

The near source factor coefficient, Nv.

This item is applicable only when the seismic coefficients Ca and Cv are user defined (UBC97SeismicCoeffFlag = 2).

UBC97I

The importance factor.

UBC97R

The overstrength factor.

## Remarks

This function retrieves auto seismic loading parameters for the 1997 UBC code.

The function returns zero if the parameters are successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetSeismicParametersUBC97()

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

      Dim UBC97SeismicCoeffFlag As Long

      Dim UBC97SoilProfileType As Long

      Dim UBC97Z As Double

      Dim UBC97Ca As Double

      Dim UBC97Cv As Double

      Dim UBC97NearSourceFlag As Long

      Dim UBC97SourceType As Long

      Dim UBC97Dist As Double

      Dim UBC97Na As Double

      Dim UBC97Nv As Double

      Dim UBC97I As Double

      Dim UBC97R As Double



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



   'assign UBC97 parameters

      ret = SapModel.LoadPatterns.AutoSeismic.SetUBC97("EQX", 1, 0.05, 1, 0.035, 0, False, 0, 0, 1, 3, 0.4, 0, 0, 1, 3, 5, 0, 0, 1.15, 6)



   'get UBC97 parameters

      ret = SapModel.LoadPatterns.AutoSeismic.GetUBC97("EQX", DirFlag, Eccen, PeriodFlag, CT, UserT, UserZ, TopZ, BottomZ, UBC97SeismicCoeffFlag, UBC97SoilProfileType, UBC97Z, UBC97Ca, UBC97Cv, UBC97NearSourceFlag, UBC97SourceType, UBC97Dist, UBC97Na, UBC97Nv, UBC97I, UBC97R)



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

[SetUBC97](SetUBC97_%7BAuto_Seismic%7D.md)
