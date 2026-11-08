# SetIBC2003

## Syntax

SapObject.SapModel.LoadPatterns.AutoSeismic.SetIBC2003

## VB6 Procedure

Function SetIBC2003(ByVal Name As String, ByVal DirFlag As Long, ByVal Eccen As Double, ByVal PeriodFlag As Long, ByVal CT As Double, ByVal UserT As Double, ByVal UserZ As Boolean, ByVal TopZ As Double, ByVal BottomZ As Double, ByVal IBC2003SG As Long, ByVal IBC2003SeismicCoeffFlag As Long, ByVal IBC2003Site As Long, ByVal IBC2003SS As Double, ByVal IBC2003S1 As Double, ByVal IBC2003Fa As Double, ByVal IBC2003Fv As Double, ByVal IBC2003R As Double, ByVal IBC2003Omega As Double, ByVal IBC2003Cd As Double) As Long

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

The code-specified CT factor (ft). This item applies when the PeriodFlag item is 1 or 2.

UserT

The user specified time period. This item applies when the PeriodFlag item is 3. [s]

UserZ

This item is True if the top and bottom elevations of the seismic load are user specified. It is False if the elevations are determined by the program.

TopZ

This item applies only when the UserZ item is True. It is the global Z-coordinate at the highest level where auto seismic loads are applied. [L]

BottomZ

This item applies only when the UserZ item is True. It is the global Z-coordinate at the lowest level where auto seismic loads are applied. [L]

IBC2003SG

This is 1, 2 or 3, indicating the seismic group.

1 = I

2 = II

3 = III

IBC2003SeismicCoeffFlag

This is either 1 or 2, indicating the seismic coefficient option.

1 = Coefficients are per code

2 = Coefficients are user defined

IBC2003Site

This is 1, 2, 3, 4 or 5, indicating the site class. This item is only used when the IBC2003SeismicCoeffFlag = 1.

1 = A

2 = B

3 = C

4 = D

5 = E

IBC2003SS

The response acceleration for short periods, (g).

IBC2003S1

The response acceleration for a one second period, (g).

IBC2003Fa

The site coefficient Fa. This item is used only when the IBC2003SeismicCoeffFlag = 2.

IBC2003Fv

The site coefficient Fv. This item is used only when the IBC2003SeismicCoeffFlag = 2

IBC2003R

The response modification factor.

IBC2003Omega

The system overstrength factor.

IBC2003Cd

The deflection amplification factor.

## Remarks

This function assigns auto seismic loading parameters for the 2003 IBC code.

The function returns zero if the parameters are successfully assigned; otherwise it returns a nonzero value.

## VBA Example



```
Sub AssignSeismicIBC2003()

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



   'assign IBC2003 parameters

      ret = SapModel.LoadPatterns.AutoSeismic.SetIBC2003("EQX", 1, 0.05, 1, 0.035, 0, False, 0, 0, 1, 1, 3, 1, 0.4, 0, 0, 8, 3, 5.5)



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

[GetIBC2003](GetIBC2003_%7BAuto_Seismic%7D.md)
