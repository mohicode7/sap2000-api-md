# SetORebar\_1

## Syntax

SapObject.SapModel.PropMaterial.SetORebar\_1

## VB6 Procedure

Function SetORebar\_1(ByVal Name As String, ByVal Fy As Double, ByVal Fu As Double, ByVal eFy As Double, ByVal eFu As Double, ByVal SSType As Long, ByVal SSHysType As Long, ByVal StrainAtHardening As Double, ByVal StrainUltimate As Double, ByVal FinalSlope As Double, ByVal UseCaltransSSDefaults As Boolean, Optional ByVal Temp As Double = 0) As Long

## Parameters

Name

The name of an existing rebar material property.

Fy

The minimum yield stress. [F/L2]

Fu

The minimum tensile stress. [F/L2]

eFy

The expected yield stress. [F/L2]

eFu

The expected tensile stress. [F/L2]

SSType

This is 0, 1 or 2, indicating the stress-strain curve type.

0 = User defined

1 = Parametric - Simple

2 = Parametric - Park

SSHysType

This is 0, 1 or 2, indicating the stress-strain hysteresis type.

0 = Elastic

1 = Kinematic

2 = Takeda

StrainAtHardening

This item applies only when parametric stress-strain curves are used and when UseCaltransSSDefaults is False. It is the strain at the onset of strain hardening.

StrainUltimate

This item applies only when parametric stress-strain curves are used and when UseCaltransSSDefaults is False. It is the ultimate strain capacity. This item must be larger than the StrainAtHardening item.

FinalSlope

This item applies only to parametric stress-strain curves. It is a multiplier on the material modulus of elasticity, E. This value multiplied times E gives the final slope of the curve.

UseCaltransSSDefaults

If this item is True, the program uses Caltrans default controlling strain values, which are bar size dependent.

Temp

This item applies only if the specified material has properties that are temperature dependent. That is, it applies only if properties are specified for the material at more than one temperature.

This item is the temperature at which the specified data applies. The temperature must have been defined previously for the material.

## Remarks

This function sets the other material property data for rebar materials.

The function returns zero if the data is successfully assigned; otherwise it returns a nonzero value.

## VBA Example



```
Sub AssignMatPropRebarData_1()

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



   'initialize new material property

      ret = SapModel.PropMaterial.SetMaterial("Rebar", MATERIAL_REBAR)



   'assign other properties

      ret = SapModel.PropMaterial.SetORebar_1("Rebar", 62, 93, 70, 102, 2, 2, 0.02, 0.1, -0.1, False)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 12.00.

This function supersedes [SetORebar](../../../Obsolete_Functions/SetORebar.md).

## See Also

[GetORebar\_1](GetORebar_1.md)
