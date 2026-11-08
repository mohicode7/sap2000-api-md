# SetUBC97Iso

## Syntax

SapObject.SapModel.LoadPatterns.AutoSeismic.SetUBC97Iso

## VB6 Procedure

Function SetUBC97Iso(ByVal Name As String, ByVal DirFlag As Long, ByVal Eccen As Double, ByVal UserZ As Boolean, ByVal TopZ As Double, ByVal BottomZ As Double, ByVal UBC97IsoSeismicCoeffFlag As Long, ByVal UBC97IsoSoilProfileType As Long, ByVal UBC97IsoZ As Double, ByVal UBC97IsoCv As Double, ByVal UBC97IsoNearSourceFlag As Long, ByVal UBC97IsoSourceType As Long, ByVal UBC97IsoDist As Double, ByVal UBC97IsoNv As Double, ByVal UBC97IsoRI As Double, ByVal UBC97IsoBD As Double, ByVal UBC97IsoKDmax As Double, ByVal UBC97IsoKDmin As Double) As Long

## Parameters

Name

The name of an existing Quake-type load pattern.

DirFlag

This is either 1 or 2, indicating the seismic load direction.

1 = Global X

2 = Global Y

Eccen

The eccentricity ratio that applies to all diaphragms.

UserZ

This item is True if the top and bottom elevations of the seismic load are user specified. It is False if the elevations are determined by the program.

TopZ

This item applies only when the UserZ item is True. It is the global Z-coordinate at the highest level where auto seismic loads are applied. [L]

BottomZ

This item applies only when the UserZ item is True. It is the global Z-coordinate at the lowest level where auto seismic loads are applied. [L]

UBC97IsoSeismicCoeffFlag

This is either 1 or 2, indicating if the seismic coefficient Cv is per code or user defined.

1 = Per code

2 = User defined

UBC97IsoSoilProfileType

This is 1, 2, 3, 4 or 5, indicating the soil profile type.

1 = SA

2 = SB

3 = SC

4 = SD

5 = SE

This item is applicable only when the seismic coefficients Ca and Cv are calculated in accordance with code (UBC97SeismicCoeffFlag = 1).

UBC97IsoZ

This is 0.075, 0.15, 0.2, 0.3 or 0.4, indicating the seismic zone factor.

This item is applicable only when the seismic coefficient Cv is calculated in accordance with code (UBC97IsoSeismicCoeffFlag = 1).

UBC97IsoCv

The seismic coefficient, Cv.

This item is applicable only when the seismic coefficient Cv is user defined (UBC97IsoSeismicCoeffFlag = 2).

UBC97IsoNearSourceFlag

This is either 1 or 2, indicating if the near source factor coefficient Nv is per code or user defined.

1 = Per code

2 = User defined

This item is applicable only when the seismic coefficient Cv is calculated per code (UBC97IsoSeismicCoeffFlag = 1) and UBC97IsoZ = 0.4.

UBC97IsoSourceType

This is 1, 2 or 3, indicating the seismic source type.

1 = A

2 = B

3 = C

This item is applicable only when the seismic coefficient Cv is calculated in accordance with code (UBC97IsoSeismicCoeffFlag = 1), UBC97IsoZ = 0.4, and the near source factor coefficient Nv is calculated in accordance with code (UBC97IsoNearSourceFlag = 1).

UBC97IsoDist

This is the distance to the seismic source in kilometers.

This item is only applicable when the seismic coefficient Cv is calculated in accordance with code (UBC97IsoSeismicCoeffFlag = 1), UBC97IsoZ = 0.4, and the near source factor coefficient Nv is calculated in accordance with code (UBC97IsoNearSourceFlag = 1).

UBC97IsoNv

The near source factor coefficient, Nv.

This item is applicable only when the seismic coefficient Cv is user defined (UBC97IsoSeismicCoeffFlag = 2) and the near source factor coefficient Nv is user defined (UBC97IsoNearSourceFlag = 2).

UBC97IsoRI

The overstrength factor, Ri.

UBC97IsoBD

The coefficient for damping.

UBC97IsoKDmax

The maximum effective stiffness of the isolation system. [F/L]

UBC97IsoKDmin

The minimum effective stiffness of the isolation system. [F/L]

## Remarks

This function assigns auto seismic loading parameters for seismically isolated buildings using the 1997 UBC code.

The function returns zero if the parameters are successfully assigned; otherwise it returns a nonzero value.

## VBA Example



```
Sub AssignSeismicUBC97Iso()

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



   'assign UBC97Iso parameters

      ret = SapModel.LoadPatterns.AutoSeismic.SetUBC97Iso("EQX", 1, 0.05, False, 0, 0, 1, 3, 0.4, 0, 1, 3, 5, 0, 2, 1.25, 200, 150)



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

[GetUBC97Iso](GetUBC97Iso.md)
