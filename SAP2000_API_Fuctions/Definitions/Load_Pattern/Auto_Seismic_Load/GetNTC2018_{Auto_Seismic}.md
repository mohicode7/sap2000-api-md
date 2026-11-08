# GetNTC2018

## Syntax

SapObject.SapModel.LoadPatterns.AutoSeismic.GetNTC2018

## VB6 Procedure

Function GetNTC2018(ByVal Name As String, ByRef DirFlag As Long, ByRef Eccen As Double, ByRef PeriodFlag As Long, ByRef UserT As Double, ByRef UserZ As Boolean, ByRef TopZ As Double, ByRef BottomZ As Double, ByRef ParamsOption As Long, ByRef Latitude As Double, ByRef Longitude As Double, ByRef Island As Long, ByRef LimitState As Long, ByRef UsageClass As Long, ByRef NomLife As Double, ByRef PeakAccel As Double, ByRef F0 As Double, ByRef Tcs As Double, ByRef SpecType As Long, ByRef SoilType As Long, ByRef Topography As Long, ByRef hRatio As Double, ByRef Damping As Double, ByRef q As Double, ByRef lambda As Double) As Long

## Parameters

Name

The name of an existing Quake-type load pattern with a NTC 2018 auto seismic load assignment.

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

UserT

The user specified time period. This item is meaningful when the PeriodFlag item is 3. [s]

UserZ

This item is True if the top and bottom elevations of the seismic load are user specified. It is False if the elevations are determined by the program.

TopZ

This item applies only when the UserZ item is True. It is the global Z-coordinate at the highest level where auto seismic loads are applied. [L]

BottomZ

This item applies only when the UserZ item is True. It is the global Z-coordinate at the lowest level where auto seismic loads are applied. [L]

ParamsOption

This is 1, 2, or 3, indicating the option for defining the parameters.

1 = by latitude and longitude

2 = by island

3 = user specified

Latitude, Longitude

The latitude and longitude for which the seismic coefficients are obtained. These items are meaningful only when ParamsOption = 1.

Island

This is one of the following values. This item is used only when ParamsOption = 2.

 1 = Alicudi  
 2 = Arcipelago Toscano  
 3 = Filcudi  
 4 = Isole Egadi  
 5 = Lampedusa  
 6 = Linosa  
 7 = Lipari  
 8 = Palmarola  
 9 = Panarea  
 10 = Pantelleria  
 11 = Ponza  
12 = Salina  
13 = Santo Stefano  
14 = Sardegna  
15 = Stromboli  
16 = Tremiti  
17 = Ustica  
18 = Ventotene  
19 = Vulcano  
20 = Zannone

LimitState

This is 1, 2, 3, or 4, indicating the limit state.

1 = SLO

2 = SLD

3 = SLV

4 = SLC

UsageClass

This is 1, 2, 3, or 4, indicating the usage class.

1 = I

2 = II

3 = III

4 = IV

NomLife

The nominal life to be considered.

PeakAccel

The peak ground acceleration, ag/g.

F0

The magnitude factor, F0.

Tcs

The reference period, Tc\* [s].

SpecType

This is 1, 2, 3, or 4, indicating the type of spectrum to consider.

1 = Elastic horizontal

2 = Elastic vertical

3 = Design horizontal

4 = Design vertical

SoilType

This is 1, 2, 3, 4, or 5, indicating the subsoil type.

1 = A

2 = B

3 = C

4 = D

5 = E

Topography

This is 1, 2, 3, or 4, indicating the topography type.

1 = T1

2 = T2

3 = T3

4 = T4

hRatio

The ratio for the site altitude at the base of the hill to the height of the hill.

Damping

The damping, in percent. This is only applicable for SpecType 1 and 2.

q

The behavior correction factor. This is only applicable for SpecType 3 and 4.

lambda

A correction factor.

## Remarks

This function retrieves auto seismic loading parameters for the NTC 2018 code.

The function returns zero if the parameters are successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetSeismicParametersNTC2018()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim DirFlag As Long

      Dim Eccen As Double

      Dim PeriodFlag As Long

      Dim UserT As Double

      Dim UserZ As Boolean

      Dim TopZ As Double

      Dim BottomZ As Double

      Dim ParamsOption As Long

      Dim Latitude As Double

      Dim Longitude As Double
      Dim Island As Long
      Dim LimitState As Long
      Dim UsageClass As Long

      Dim NomLife As Double

      Dim PeakAccel As Double

      Dim F0 As Double

      Dim Tcs As Double
      Dim SpecType As Long
      Dim SoilType As Long
      Dim Topography As Long

      Dim hRatio As Double

      Dim Damping As Double

      Dim q As Double

      Dim lambda As Double


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

      ret = SapModel.LoadPatterns.Add("EQX", eLoadPatternType_QUAKE)



   'assign NTC2018 parameters

      ret = SapModel.LoadPatterns.AutoSeismic.SetNTC2018("EQX", 1, 0.05, 1, 0, False, 0, 0, 3, 0, 0, 1, 1, 1, 50, 0.2, 2.4, 0.3, 3, 2, 1, 1, 5, 1, 0.85 )



   'get NTC2018 parameters

      ret = SapModel.LoadPatterns.AutoSeismic.GetNTC2018("EQX", DirFlag, Eccen, PeriodFlag, UserT, UserZ, TopZ, BottomZ, ParamsOption, Latitude, Longitude, Island, LimitState, UsageClass, NomLife, PeakAccel, F0, Tcs, SpecType, SoilType, Topography, hRatio, Damping, q, lambda)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in v20.1.0.

## See Also

[SetNTC2018](SetNTC2018_%7BAuto_Seismic%7D.md)
