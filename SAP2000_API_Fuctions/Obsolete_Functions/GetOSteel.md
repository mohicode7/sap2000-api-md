# GetOSteel (Note:  Newer function available)

## Syntax

SapObject.SapModel.PropMaterial.GetOSteel

## VB6 Procedure

Function GetOSteel(ByVal Name As String, ByRef Fy As Double, ByRef Fu As Double, ByRef eFy As Double, ByRef eFu As Double, ByRef SSType As Long, ByRef SSHysType As Long, ByRef StrainAtHardening As Double, ByRef StrainAtMaxStress As Double, ByRef StrainAtRupture As Double, Optional ByVal Temp As Double = 0) As Long

## Parameters

Name

The name of an existing steel material property.

Fy

The minimum yield stress. [F/L2]

Fu

The minimum tensile stress. [F/L2]

eFy

The expected yield stress. [F/L2]

eFu

The expected tensile stress. [F/L2]

SSType

This is 0 or 1. indicating the stress-strain curve type.

0 = User defined

1 = Parametric - Simple

SSHysType

This is 0, 1 or 2, indicating the stress-strain hysteresis type.

0 = Elastic

1 = Kinematic

2 = Takeda

StrainAtHardening

This item applies only to parametric stress-strain curves. It is the strain at the onset of strain hardening.

StrainAtMaxStress

This item applies only to parametric stress-strain curves. It is the strain at maximum stress.

StrainAtRupture

This item applies only to parametric stress-strain curves. It is the strain at rupture.

Temp

This item applies only if the specified material has properties that are temperature dependent. That is, it applies only if properties are specified for the material at more than one temperature.

This item is the temperature at which the specified data is to be retrieved. The temperature must have been defined previously for the material.

## Remarks

This function retrieves the other material property data for steel materials.

The function returns zero if the data is successfully retrieved; otherwise it returns a nonzero value. The function returns an error if the specified material is not steel.

## VBA Example



```
Sub GetMatPropSteelData()

   'dimension variables

      Dim SapObject As Sap2000v15.SapObject

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Fy As Double

      Dim Fu As Double

      Dim eFy As Double

      Dim eFu As Double

      Dim SSType As Long

      Dim SSHysType As Long

      Dim StrainAtHardening As Double

      Dim StrainAtMaxStress As Double

      Dim StrainAtRupture As Double



   'create Sap2000 object

      Set SapObject = New Sap2000v15.SapObject



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'create model from template

      ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288)



   'initialize new material property

      ret = SapModel.PropMaterial.SetMaterial("Steel", MATERIAL_STEEL)



   'assign other properties

      ret = SapModel.PropMaterial.SetOSteel("Steel", 55, 68, 60, 70, 1, 2, 0.02, 0.1, 0.2)



   'get other properties

      ret = SapModel.PropMaterial.GetOSteel("Steel", Fy, Fu, eFy, eFu, SSType, SSHysType, StrainAtHardening, StrainAtMaxStress, StrainAtRupture)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

This function is obsolete and has been superseded by [GetOSteel\_1](../Definitions/Properties/Material/GetOSteel_1.md) as of version 12.00. This function is maintained for backwards compatibility.

## See Also

[SetOSteel](SetOSteel.md)
