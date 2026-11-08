# SetOConcrete\_1

## Syntax

SapObject.SapModel.PropMaterial.SetOConcrete\_1

## VB6 Procedure

Function SetOConcrete\_1(ByVal Name As String, ByVal fc As Double, ByVal IsLightweight As Boolean, ByVal fcsfactor As Double, ByVal sstype As Long, ByVal SSHysType As Long, ByVal StrainAtfc As Double, ByVal StrainUltimate As Double, ByVal FinalSlope As Double, Optional ByVal FrictionAngle As Double = 0, Optional ByVal DilatationalAngle As Double = 0, Optional ByVal Temp As Double = 0) As Long

## Parameters

Name

The name of an existing concrete material property.

fc

The concrete compressive strength. [F/L2]

IsLightweight

If this item is True, the concrete is assumed to be lightweight concrete.

fcsfactor

The shear strength reduction factor for lightweight concrete.

SSType

This is 0, 1 or 2, indicating the stress-strain curve type.

0 = User defined

1 = Parametric - Simple

2 = Parametric - Mander

SSHysType

This is 0, 1 or 2, indicating the stress-strain hysteresis type.

0 = Elastic

1 = Kinematic

2 = Takeda

StrainAtfc

This item applies only to parametric stress-strain curves. It is the strain at the unconfined compressive strength.

StrainUltimate

This item applies only to parametric stress-strain curves. It is the ultimate unconfined strain capacity. This item must be larger than the StrainAtfc item.

FinalSlope

This item applies only to parametric stress-strain curves. It is a multiplier on the material modulus of elasticity, E. This value multiplied times E gives the final slope on the compression side of the curve.

FrictionAngle

The Drucker-Prager friction angle, 0 <= FrictionAngle < 90. [deg]

DilatationalAngle

The Drucker-Prager dilatational angle, 0 <= DilatationalAngle < 90. [deg]

Temp

This item applies only if the specified material has properties that are temperature dependent. That is, it applies only if properties are specified for the material at more than one temperature.

This item is the temperature at which the specified data applies. The temperature must have been defined previously for the material.

## Remarks

This function sets the other material property data for concrete materials.

The function returns zero if the data is successfully assigned; otherwise it returns a nonzero value.

## VBA Example



```
Sub AssignMatPropConcreteData_1()

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

      ret = SapModel.PropMaterial.SetMaterial("Concrete", MATERIAL_CONCRETE)



   'assign other properties

      ret = SapModel.PropMaterial.SetOConcrete_1("Concrete", 5, False, 0, 1, 2, 0.0022, 0.0052, -0.1)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 12.00.

This function supersedes [SetOConcrete](SetOConcrete.md).

## See Also

[GetOConcrete\_1](GetOConcrete_1.md)
