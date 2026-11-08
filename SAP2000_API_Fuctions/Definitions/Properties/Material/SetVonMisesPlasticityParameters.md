# SetVonMisesPlasticityParameters

## Syntax

SapObject.SapModel.PropMaterial.SetVonMisesPlasticityParameters

## VB6 Procedure

Function SetVonMisesPlasticityParameters(ByVal
Name As String, ByVal UseIsoLinHardening As Boolean, ByVal IsoLinHardeningModulus
As Double, ByVal UseIsoNLHardening As Boolean, ByVal IsoNLUltimateStressRatio
As Double, ByVal IsoNLHardeningRate As Double, ByVal UseKinLinHardening
As Boolean, ByVal KinLinHardeningModulus As Double, Optional ByVal YieldStress
As Double = 0.0, Optional ByVal Temp As Double = 0.0) As Long

## Parameters

Name

The name of an existing material property.

UseIsoLinHardening

If this item is True, the linear isotropic hardening
is enabled.

IsoLinHardeningModulus

The linear isotropic
hardening modulus, used if UseIsoLinHardening = True.

UseIsoNLHardening

If this item is True,
the nonlinear isotropic saturation hardening is enabled.

IsoNLUltimateStressRatio

The ultimate stress
as a ratio of the yield stress, used if UseIsoNLHardening = True.

IsoNLHardeningRate

The hardening rate parameter, used if UseIsoNLHardening
= True.

UseKinLinHardening

If this item is True, the linear kinematic hardening
is enabled.

KinLinHardeningModulus

The linear kinematic hardening modulus, used if UseKinLinHardening
= True.

YieldStress

The yield stress.

This item applies only if the specified material has
a material type of  eMatType.NoDesign.

Temp

This item applies only
if the specified material has properties that are temperature dependent.
That is, it applies only if properties are specified for the material
at more than one temperature.

This item is the temperature
at which the specified data is to be retrieved. The temperature must have
been defined previously for the material.

## Remarks

This function sets the
Von Mises Plasticity coupled modeling parameters for a specified material

The
function returns zero if the assignment data is successfully assigned;
otherwise it returns a nonzero value.

## VBA Example



```
Sub SetMatCoupledTypeData()
   'dimension 
 variables
      Dim 
 SapObject as cOAPI
      Dim 
 SapModel As cSapModel
      Dim 
 ret As Long
 
   'create 
 Sap2000 object
      Set 
 SapObject = CreateObject("CSI.SAP2000.API.SapObject")
 
   'start 
 Sap2000 application
      SapObject.ApplicationStart
 
   'create 
 SapModel object
      Set 
 SapModel = SapObject.SapModel
 
   'initialize 
 model
      ret 
 = SapModel.InitializeNewModel
 
   'create 
 model from template
      ret 
 = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288)
 
   'initialize 
 new material property
      ret 
 = SapModel.PropMaterial.SetMaterial("Steel", eMatType.eMatType_Steel, 
 -1, "API example test", "Default")
 
   'set 
 coupled material type data
      ret 
 = SapModel.PropMaterial.SetCoupledModelType("Steel", eMatCoupledType.eMatCoupledType_VonMisesPlasticity)
 
   'set 
 Von Mises Plasticity data
      ret 
 = SapModel.PropMaterial.SetVonMisesPlasticityParameters ("Steel", 
 True, 1000.0, True, 1.2, 50, True, 1000.0)
 
   'close 
 Sap2000
      SapObject.ApplicationExit 
 False
      Set 
 SapModel = Nothing
      Set 
 SapObject = Nothing
End Sub
```



## Release Notes

Initial release in version 22.1.0

## See Also

[GetCoupledModelType](GetCoupledModelType.md)

[SetCoupledModelType](SetCoupledModelType.md)

[GetVonMisesPlasticityParameters](GetVonMisesPlasticityParameters.md)
