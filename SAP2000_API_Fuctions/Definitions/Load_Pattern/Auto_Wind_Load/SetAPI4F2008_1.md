# SetAPI4F2008\_1

## Syntax

SapObject.SapModel.LoadPatterns.AutoWind.SetAPI4F2008\_1

## VB6 Procedure

Function SetAPI4F2008\_1(ByVal Name As String, ByVal
ExposureFrom As Long, ByVal DirAngle As Double, ByVal UserZ As Boolean,
ByVal TopZ As Double, ByVal BottomZ As Double, ByVal WindSpeed As Double,
ByVal SSLFactor As Double, ByVal ShieldingFactor As Double) As Long

## Parameters

Name

The name of an existing Wind-type load case.

ExposureFrom

This is 2, 3, or 4, indicating the source of the wind
exposure.

2 = From area objects

3 = From frame objects
(open structure)

4
= From area objects and frame objects (open structure)

DirAngle

The direction angle for the wind load. This item applies
only when ExposureFrom = 3 or 4.

UserZ

This item is True if the top and bottom elevations of
the wind load are user specified. It is False if the elevations are determined
by the program.

TopZ

This item is the global Z-coordinate at the highest
level where auto wind loads are applied. [L]

BottomZ

This item is the global Z-coordinate at the lowest level
where auto wind loads are applied. [L]

WindSpeed

The design reference wind velocity, Vref, in knots.

SSLFactor

The structural safety level multiplier.

ShieldingFactor

The shielding factor.

## Remarks

This function assigns auto wind loading parameters for
API 4F 2008.

The function returns zero if the parameters are successfully
assigned; otherwise, it returns a nonzero value.

## VBA Example



```
Sub AssignWindAPI4F2008()

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

      ret = SapModel.File.New3DFrame(BeamSlab, 
 2, 144, 3, 336, 2, 432)



   'add new load case

      ret = SapModel.LoadPatterns.Add("WIND", 
 LTYPE_WIND)



   'assign API 4F 2008 parameters

      ret = SapModel.LoadPatterns.AutoWind.SetAPI4F2008_1("WIND", 
 3, 0, False, 0, 0, 93, 1.1, 0.85)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 14.1.0.

This function supersedes [SetAPI4F2008](../../../Obsolete_Functions/SetAPI4F2008.md).

## See Also

[GetAPI4F2008\_1](GetAPI4F2008_1.md)
