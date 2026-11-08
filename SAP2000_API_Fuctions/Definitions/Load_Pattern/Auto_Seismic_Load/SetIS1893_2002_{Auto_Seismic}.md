# SetIS1893\_2002

## Syntax

SapObject.SapModel.LoadPatterns.AutoSeismic.SetIS1893\_2002

## VB6 Procedure

Function SetIS1893\_2002(ByVal Name As String, ByVal DirFlag As Long, ByVal Eccen As Double, ByVal PeriodFlag As Long, ByVal CT As Double, ByVal UserT As Double, ByVal UserZ As Boolean, ByVal TopZ As Double, ByVal BottomZ As Double, ByVal INZFlag As Long, ByVal INZ As Double, ByVal INS As Long, ByVal INI As Double, ByVal INR As Double) As Long

## Parameters

Name

The name of an existing Quake-type load pattern.

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

The code-specified CT factor (m). This item applies when the PeriodFlag item is 1.

UserT

The user specified time period. This item applies when the PeriodFlag item is 3. [s]

UserZ

This item is True if the top and bottom elevations of the seismic load are user specified. It is False if the elevations are determined by the program.

TopZ

This item applies only when the UserZ item is True. It is the global Z-coordinate at the highest level where auto seismic loads are applied. [L]

BottomZ

This item applies only when the UserZ item is True. It is the global Z-coordinate at the lowest level where auto seismic loads are applied. [L]

INZFlag

This is either 1 or 2, indicating if the seismic zone factor is per code or user defined.

1 = Per code

2 = User defined

INZ

The seismic zone factor, Z.

If the seismic zone factor is per code (INZFlag = 1), this item should be one of the following: 0.10, 0.16, 0.24, 0.36.

INS

This is 1, 2 or 3, indicating the soil type.

1 = I

2 = II

3 = III

BOCA96R

The response modification factor.

INR

The importance factor.

INR

The response modification factor.

## Remarks

This function assigns auto seismic loading parameters for the 2002 IS1893 code.

The function returns zero if the parameters are successfully assigned; otherwise it returns a nonzero value.

## VBA Example



```
Sub AssignSeismicIS18932002()

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



   'assign IS1893_2002 parameters

      ret = SapModel.LoadPatterns.AutoSeismic.SetIS1893_2002("EQX", 1, 0.05, 1, 0.075, 0, False, 0, 0, 1, 0.36, 3, 1, 5)



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

[GetIS1893\_2002](GetIS1893_2002_%7BAuto_Seismic%7D.md)
