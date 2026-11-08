# SetChinese2002

## Syntax

SapObject.SapModel.LoadPatterns.AutoSeismic.SetChinese2002

## VB6 Procedure

Function SetChinese2002(ByVal Name As String, ByVal DirFlag As Long, ByVal Eccen As Double, ByVal PeriodFlag As Long, ByVal UserT As Double, ByVal UserZ As Boolean, ByVal TopZ As Double, ByVal BottomZ As Double, ByVal JGJ32002AlphaMax As Double, ByVal JGJ32002SI As Long, ByVal JGJ32002DampRatio As Double, ByVal JGJ32002Tg As Double, ByVal JGJ32002PTDF As Double, ByVal EnhancementFactor As Double) As Long

## Parameters

Name

The name of an existing Quake-type load pattern.

DirFlag

This is 1, 2 or 3, indicating the seismic load direction.

1 = Global X

2 = Global Y

3 = Global Z

Eccen

The eccentricity ratio that applies to all diaphragms.

PeriodFlag

This is either 2 or 3, indicating the time period option.

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

JGJ32002AlphaMax

The maximum influence factor.

JGJ32002SI

This is 1, 2, 3, 4, 5 or 6, indicating the seismic intensity.

1 = 6  (0.05g)

2 = 7  (0.10g)

3 = 7  (0.15g)

4 = 8  (0.20g)

5 = 8  (0.30g)

6 = 9  (0.40g)

JGJ32002DampRatio

The damping ratio.

JGJ32002Tg

The characteristic ground period. [s]

JGJ32002PTDF

The period time discount factor.

EnhancementFactor

The enhancement factor.

## Remarks

This function assigns auto seismic loading parameters for the Chinese 2002 code.

The function returns zero if the parameters are successfully assigned; otherwise it returns a nonzero value.

## VBA Example



```
Sub AssignSeismicChinese2002()

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



   'assign Chinese 2002 parameters

      ret = SapModel.LoadPatterns.AutoSeismic.SetChinese2002("EQX", 1, 0.05, 2, 0, False, 0, 0, 0.16, 4, 0.06, 0.4, 1, 1)



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

[GetChinese2002](GetChinese2002_%7BAuto_Seismic%7D.md)
