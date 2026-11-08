# SetEurocode82004  (Note:  Newer function available)

## Syntax

SapObject.SapModel.LoadPatterns.AutoSeismic.SetEurocode82004

## VB6 Procedure

Function SetEurocode82004(ByVal Name As String, ByVal DirFlag As Long, ByVal Eccen As Double, ByVal PeriodFlag As Long, ByVal CT As Double, ByVal UserT As Double, ByVal UserZ As Boolean, ByVal TopZ As Double, ByVal BottomZ As Double, ByVal EURO2004GroundType As Long, ByVal EURO2004SpectrumType As Long, ByVal EURO2004ag As Double, ByVal EURO2004Beta As Double, ByVal EURO2004q As Double, ByVal EURO2004Lambda As Double) As Long

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

The code-specified Ct factor. This item applies when the PeriodFlag item is 1.

UserT

The user specified time period. This item applies when the PeriodFlag item is 3. [s]

UserZ

This item is True if the top and bottom elevations of the seismic load are user specified. It is False if the elevations are determined by the program.

TopZ

This item applies only when the UserZ item is True. It is the global Z-coordinate at the highest level where auto seismic loads are applied. [L]

BottomZ

This item applies only when the UserZ item is True. It is the global Z-coordinate at the lowest level where auto seismic loads are applied. [L]

EURO2004GroundType

This is 1, 2, 3, 4 or 5, indicating the ground type.

1 = A

2 = B

3 = C

4 = D

5 = E

EURO2004SpectrumType

This is 1 or 2, indicating the spectrum type.

1 = Type 1

2 = Type 2

EURO2004ag

The design ground acceleration in g, ag.

EURO2004Beta

The lower bound factor, Beta.

EURO2004q

The behavior factor, q.

EURO2004Lambda

The correction factor, Lambda.

## Remarks

This function assigns auto seismic loading parameters for the Eurocode 8 2004 code.

The function returns zero if the parameters are successfully assigned; otherwise it returns a nonzero value.

## VBA Example



```
Sub AssignSeismicEurocode82004()

   'dimension variables

      Dim SapObject As Sap2000v15.SapObject

      Dim SapModel As cSapModel

      Dim ret As Long



   'create Sap2000 object

      Set SapObject = New Sap2000v15.SapObject



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

      ret = SapModel.LoadPatterns.AutoSeismic.SetEurocode82004("EQX", 2, 0.1, 2, 0.075, 0, False, 0, 0, 2, 1, 0.4, 0.2, 2, 1)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 14.00.

The function is obsolete and has been superseded by [SetEurocode82004\_1](../Definitions/Load_Pattern/Auto_Seismic_Load/SetEurocode82004_1_%7BAuto_Seismic%7D.md) as of version 14.1.0. This function is maintained for backward compatibility. New function added.

## See Also

[GetEurocode82004](GetEurocode82004_%28Auto_Seismic%29.md)
