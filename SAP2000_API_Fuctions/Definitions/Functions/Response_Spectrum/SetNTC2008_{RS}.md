# SetNTC2008

## Syntax

SapObject.SapModel.Func.FuncRS.SetNTC2008

## VB6 Procedure

Function SetNTC2008(ByVal Name As String, ByVal ParamsOption As Long, ByVal Latitude As Double, ByVal Longitude As Double, ByVal Island As Long, ByVal LimitState As Long, ByVal UsageClass As Long, ByVal NomLife As Double, ByVal PeakAccel As Double, ByVal F0 As Double, ByVal Tcs As Double, ByVal SpecType As Long, ByVal SoilType As Long, ByVal Topography As Long, ByVal hRatio As Double, ByVal Damping As Double, ByVal q As Double) As Long

## Parameters

Name

The name of an existing or new function. If this is an existing function, that function is modified; otherwise, a new function is added.

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

## Remarks

This function defines a NTC2008 response spectrum function.

The function returns zero if the function is successfully defined; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetRSFuncNTC2008()

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

      ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288)



   'add NTC2008 RS function

      ret = SapModel.Func.FuncRS.SetNTC2008("RS-1", 1, 45.9, 12.6, 1, 3, 2, 50, 0.2, 2.4, 0.3, 3, 2, 1, 1, 5, 1)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 18.1.0.

## See Also

[GetNTC2008](GetNTC2008_%7BRS%7D.md)
