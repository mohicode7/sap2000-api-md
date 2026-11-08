# GetVonMisesPlasticityParameters

## Syntax

SapObject.SapModel.PropMaterial.GetVonMisesPlasticityParameters

## VB6 Procedure

Function GetVonMisesPlasticityParameters(ByVal
Name As String, ByRef UseIsoLinHardening As Boolean, ByRef IsoLinHardeningModulus
As Double, ByRef UseIsoNLHardening As Boolean, ByRef IsoNLUltimateStressRatio
As Double, ByRef IsoNLHardeningRate As Double, ByRef UseKinLinHardening
As Boolean, ByRef KinLinHardeningModulus As Double, ByRef YieldStress
As Double, Optional ByVal Temp As Double = 0.0)
As Long

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

This function gets the
Von Mises Plasticity coupled modeling parameters for a specified material

The function returns
zero if the assignment data is successfully obtained; otherwise it returns
a nonzero value.

## VBA Example



```
Sub GetMatCoupledTypeData()
   'dimension 
 variables
      Dim 
 SapObject as cOAPI
      Dim 
 SapModel As cSapModel
      Dim 
 UseIsoLinHardening As Boolean
      Dim 
 IsoLinHardeningModulus As Double
      Dim 
 UseIsoNLHardening As Boolean
      Dim 
 IsoNLUltimateStressRatio As Double
      Dim 
 IsoNLHardeningRate As Double
      Dim 
 UseKinLinHardening As Boolean
      Dim 
 KinLinHardeningModulus As Double
      Dim 
 YieldStress As Double
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
 
   'set coupled material type 
 data
      ret 
 = SapModel.PropMaterial.SetCoupledModelType("Steel", eMatCoupledType.eMatCoupledType_VonMisesPlasticity)
 
   'get 
 Von Mises Plasticity data
      ret 
 = SapModel.PropMaterial.GetVonMisesPlasticityParameters ("Steel", 
 UseIsoLinHardening, IsoLinHardeningModulus, UseIsoNLHardening, IsoNLUltimateStressRatio, 
 IsoNLHardeningRate, UseKinLinHardening, KinLinHardeningModulus, YieldStress)
 
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

[SetVonMisesPlasticityParameters](SetVonMisesPlasticityParameters.md)
