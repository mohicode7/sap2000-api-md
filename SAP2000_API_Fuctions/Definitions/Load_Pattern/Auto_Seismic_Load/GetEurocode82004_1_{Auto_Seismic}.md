# GetEurocode82004\_1

## Syntax

SapObject.SapModel.LoadPatterns.AutoSeismic.GetEurocode82004\_1

## VB6 Procedure

Function GetEurocode82004\_1(ByVal Name As String, ByRef DirFlag As Long, ByRef Eccen As Double, ByRef PeriodFlag As Long, ByRef CT As Double, ByRef UserT As Double, ByRef UserZ As Boolean, ByRef TopZ As Double, ByRef BottomZ As Double, ByRef EURO2004Country As Long, ByRef EURO2004SpectrumType As Long, ByRef EURO2004GroundType As Long, ByRef EURO2004ag As Double, ByRef EURO2004S As Double, ByRef EURO2004Tb As Double, ByRef EURO2004Tc As Double, ByRef EURO2004Td As Double, ByRef EURO2004Beta As Double, ByRef EURO2004q As Double, ByRef EURO2004Lambda As Double) As Long

## Parameters

Name

The name of an existing Quake-type load pattern with a Eurocode 8 2004 auto seismic load assignment.

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

The code-specified Ct factor. This item applies when the PeriodFlag item is 1.

UserT

The user specified time period. This item applies when the PeriodFlag item is 3. [s]

UserZ

This item is True if the top and bottom elevations of the seismic load are user specified. It is False if the elevations are determined by the program.

TopZ

This item applies only when the UserZ item is True. It is the global Z-coordinate at the highest level where auto seismic loads are applied. [L]

BottomZ

This item applies only when the UserZ item is True. It is the global Z-coordinate at the lowest level where auto seismic loads are applied. [L]

EURO2004Country

This is 0, 1, 5, 6 or 10 indicating the country for which the Nationally Determined Parameters (NDPs) are specified.

 0 = Other (NDPs are user specified)

 1 = CEN Default

 5 = Norway

 6 = Singapore

 10 = Portugal

EURO2004SpectrumType

This is 1 or 2, indicating the spectrum type.

1 = Type 1

2 = Type 2 (Does not apply when EURO2004Country = 5 or 6)

EURO2004GroundType

This is 1, 2, 3, 4 or 5, or 6, indicating the ground type.

1 = A (Does not apply when EURO2004Country = 6)

2 = B (Does not apply when EURO2004Country = 6)

3 = C

4 = D

5 = E (Does not apply when EURO2004Country = 6)

6 = S1 (Only applies when EURO2004Country = 6)

EURO2004ag

The design ground acceleration in g, ag.

EURO2004S

The soil factor, S.

EURO2004Tb

The lower limit of period of the constant spectral acceleration branch, Tb.

EURO2004Tc

The upper limit of period of the constant spectral acceleration branch, Tc.

EURO2004Td

The period defining the start of the constant displacement range, Td.

EURO2004Beta

The lower bound factor, Beta.

EURO2004q

The behavior factor, q.

EURO2004Lambda

The correction factor, Lambda.

## Remarks

This function retrieves auto seismic loading parameters for the Eurocode 8 2004 code.

The function returns zero if the parameters are successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetSeismicParametersEurocode82004()

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

      Dim EURO2004Country As Long

      Dim EURO2004SpectrumType As Long

      Dim EURO2004GroundType As Long

      Dim EURO2004ag As Double

      Dim EURO2004S As Double

      Dim EURO2004Tb As Double

      Dim EURO2004Tc As Double

      Dim EURO2004Td As Double

      Dim EURO2004Beta As Double

      Dim EURO2004q As Double

      Dim EURO2004Lambda As Double



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



   'assign Eurocode 8 2004 parameters

      ret = SapModel.LoadPatterns.AutoSeismic.SetEurocode82004_1("EQX", 2, 0.1, 2, 0.075, 0, False, 0, 0, 1, 1, 2, 0.4, 1, 1, 1, 1, 0.2, 2, 1)



   'get Eurocode 8 2004 parameters

      ret = SapModel.LoadPatterns.AutoSeismic.GetEurocode82004_1("EQX", DirFlag, Eccen, PeriodFlag, CT, UserT, UserZ, TopZ, BottomZ, EURO2004Country, EURO2004SpectrumType, EURO2004GroundType, EURO2004ag, EURO2004S, EURO2004Tb, EURO2004Tc, EURO2004Td, EURO2004Beta, EURO2004q, EURO2004Lambda)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 14.1.0.

This function supersedes [GetEurocode82004](../../../Obsolete_Functions/GetEurocode82004_%28Auto_Seismic%29.md).

Added Portugal as a Country parameter in SAP2000 Version 15.0.0 and CSiBridge Version 15.1.0.

Added Singapore as a Country parameter in v20.1.0.

## See Also

[SetEurocode82004\_1](SetEurocode82004_1_%7BAuto_Seismic%7D.md)
