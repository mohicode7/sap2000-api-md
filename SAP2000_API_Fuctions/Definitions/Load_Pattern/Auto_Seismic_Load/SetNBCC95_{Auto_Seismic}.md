# SetNBCC95

## Syntax

SapObject.SapModel.LoadPatterns.AutoSeismic.SetNBCC95

## VB6 Procedure

Function SetNBCC95(ByVal Name As String, ByVal DirFlag As Long, ByVal eccen As Double, ByVal NBCCPFlag As Long, ByVal NBCC95DS As Double, ByVal PeriodFlag As Long, ByVal UserT As Double, ByVal UserZ As Boolean, ByVal TopZ As Double, ByVal BottomZ As Double, ByVal NBCC95ZA As Long, ByVal NBCC95ZV As Long, ByVal NBCC95ZVFlag As Long, ByVal NBCC95ZVR As Double, ByVal NBCC95I As Double, ByVal NBCC95F As Double, ByVal NBCC95R As Double) As Long

## Parameters

Name

The name of an existing Quake-type load pattern.

DirFlag

This is either 1 or 2, indicating the seismic load direction.

1 = Global X

2 = Global Y

Eccen

The eccentricity ratio that applies to all diaphragms.

NBCCPFlag

This is either 1 or 2, indicating the structure type.

1 = Moment frame

2 = Other

NBCC95DS

This item applies only when the NBCCPFlag = 2. It is the dimension of the lateral load resisting system in the direction of the applied forces. [L]

PeriodFlag

This is 1, 2 or 3, indicating the time period option.

1 = Code

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

NBCC95ZA

This is 0, 1, 2, 3, 4, 5 or 6, indicating the acceleration related zone, Za.

NBCC95ZV

This is 0, 1, 2, 3, 4, 5 or 6, indicating the velocity related zone, Zv.

NBCC95ZVFlag

This is either 1 or 2, indicating how the zonal velocity ratio, V, is specified.

1 = From code based on Zv

2 = User specified

NBCC95ZVR

The zonal velocity ratio, V. This item is used only when NBCC95ZVFlag = 2.

NBCC95I

The importance factor.

NBCC95F

The foundation factor.

NBCC95R

The force modification factor.

## Remarks

This function assigns auto seismic loading parameters for the 1995 NBCC code.

The function returns zero if the parameters are successfully assigned; otherwise it returns a nonzero value.

## VBA Example



```
Sub AssignSeismicNBCC95()

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



   'assign NBCC95 parameters

      ret = SapModel.LoadPatterns.AutoSeismic.SetNBCC95("EQX", 1, 0.05, 1, 0, 1, 0, False, 0, 0, 4, 5, 1, 0, 1, 1.3, 4)



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

[GetNBCC95](GetNBCC95_%7BAuto_Seismic%7D.md)
