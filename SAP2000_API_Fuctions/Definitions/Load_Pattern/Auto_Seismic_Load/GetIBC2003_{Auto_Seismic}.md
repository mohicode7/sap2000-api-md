# GetIBC2003

## Syntax

SapObject.SapModel.LoadPatterns.AutoSeismic.GetIBC2003

## VB6 Procedure

Function GetIBC2003(ByVal Name As String, ByRef DirFlag As Long, ByRef Eccen As Double, ByRef PeriodFlag As Long, ByRef CT As Double, ByRef UserT As Double, ByRef UserZ As Boolean, ByRef TopZ As Double, ByRef BottomZ As Double, ByRef IBC2003SG As Long, ByRef IBC2003SeismicCoeffFlag As Long, ByRef IBC2003Site As Long, ByRef IBC2003SS As Double, ByRef IBC2003S1 As Double, ByRef IBC2003Fa As Double, ByRef IBC2003Fv As Double, ByRef IBC2003R As Double, ByRef IBC2003Omega As Double, ByRef IBC2003Cd As Double) As Long

## Parameters

Name

The name of an existing Quake-type load pattern with a IBC2003 auto seismic load assignment.

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

This is either 1, 2, 3, 4 or 5, indicating the site class. This item is filled only when the IBC2003SeismicCoeffFlag = 1.

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

The site coefficient Fa.

IBC2003Fv

The site coefficient Fv.

IBC2003R

The response modification factor.

IBC2003Omega

The system overstrength factor.

IBC2003Cd

The deflection amplification factor.

## Remarks

This function retrieves auto seismic loading parameters for the 2003 IBC code.

The function returns zero if the parameters are successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetSeismicParametersIBC2003()

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

      Dim IBC2003SG As Long

      Dim IBC2003SeismicCoeffFlag As Long

      Dim IBC2003Site As Long

      Dim IBC2003SS As Double

      Dim IBC2003S1 As Double

      Dim IBC2003Fa As Double

      Dim IBC2003Fv As Double

      Dim IBC2003R As Double

      Dim IBC2003Omega As Double

      Dim IBC2003Cd As Double



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



   'get IBC2003parameters

      ret = SapModel.LoadPatterns.AutoSeismic.GetIBC2003("EQX", DirFlag, Eccen, PeriodFlag, CT, UserT, UserZ, TopZ, BottomZ, IBC2003SG, IBC2003SeismicCoeffFlag, IBC2003Site, IBC2003SS, IBC2003S1, IBC2003Fa, IBC2003Fv, IBC2003R, IBC2003Omega, IBC2003Cd)



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

[SetIBC2003](SetIBC2003_%7BAuto_Sesimic%7D.md)
