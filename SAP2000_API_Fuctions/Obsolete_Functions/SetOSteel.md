# SetOSteel (Note:  Newer function available)

## Syntax

SapObject.SapModel.PropMaterial.SetOSteel

## VB6 Procedure

Function SetOSteel(ByVal Name As String, ByVal Fy As Double, ByVal Fu As Double, ByVal eFy As Double, ByVal eFu As Double, ByVal SSType As Long, ByVal SSHysType As Long, ByVal StrainAtHardening As Double, ByVal StrainAtMaxStress As Double, ByVal StrainAtRupture As Double, Optional ByVal Temp As Double = 0) As Long

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

This is 0 or 1, indicating the stress-strain curve type.

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

This item applies only to parametric stress-strain curves. It is the strain at maximum stress. This item must be larger than the StrainAtHardening item.

StrainAtRupture

This item applies only to parametric stress-strain curves. It is the strain at rupture. This item must be larger than the StrainAtMaxStress item.

Temp

This item applies only if the specified material has properties that are temperature dependent. That is, it applies only if properties are specified for the material at more than one temperature.

This item is the temperature at which the specified data applies. The temperature must have been defined previously for the material.

## Remarks

This function sets the other material property data for steel materials.

The function returns zero if the data is successfully assigned; otherwise it returns a nonzero value.

## VBA Example



```
Sub AssignMatPropSteelData()

   'dimension variables

      Dim SapObject As Sap2000v15.SapObject

      Dim SapModel As cSapModel

      Dim ret As Long



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



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

This function is obsolete and has been superseded by [SetOSteel\_1](../Definitions/Properties/Material/SetOSteel_1.md) as of version 12.00. This function is maintained for backwards compatibility.

## See Also

[GetOSteel](GetOSteel.md)
