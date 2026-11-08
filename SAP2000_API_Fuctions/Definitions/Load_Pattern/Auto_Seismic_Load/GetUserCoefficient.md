# GetUserCoefficient

## Syntax

SapObject.SapModel.LoadPatterns.AutoSeismic.GetBOCA96

## VB6 Procedure

Function GetBOCA96(ByVal Name As String, ByRef DirFlag As Long, ByRef Eccen As Double, ByRef UserZ As Boolean, ByRef TopZ As Double, ByRef BottomZ As Double, ByRef c As Double, ByRef k As Double) As Long

## Parameters

Name

The name of an existing Quake-type load pattern with a User Coefficient auto seismic load assignment.

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

c

The base shear coefficient.

k

The building height exponent.

## Remarks

This function retrieves auto seismic loading parameters for User Coefficient type loading.

The function returns zero if the parameters are successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetSeismicParametersUserCoefficient()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim DirFlag As Long

      Dim Eccen As Double

      Dim UserT As Double

      Dim UserZ As Boolean

      Dim TopZ As Double

      Dim BottomZ As Double

      Dim c As Double

      Dim k As Double



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



   'assign user coefficient parameters

      ret = SapModel.LoadPatterns.AutoSeismic.SetUserCoefficient("EQX", 1, 0.06, False, 0, 0, 0.15, 1.3)



   'get user coefficient parameters

      ret = SapModel.LoadPatterns.AutoSeismic.GetUserCoefficient("EQX", DirFlag, Eccen, UserZ, TopZ, BottomZ, c, k)



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

[SetUserCoefficient](SetUserCoefficient.md)
