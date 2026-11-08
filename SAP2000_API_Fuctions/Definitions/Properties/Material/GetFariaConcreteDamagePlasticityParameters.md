# GetFariaConcreteDamagePlasticityParameters

## Syntax

SapObject.SapModel.PropMaterial.GetFariaConcreteDamagePlasticityParameters

## VB6 Procedure

Function GetFariaConcreteDamagePlasticityParameters(ByVal Name As String, ByRef beta As Double, ByRef fb0ratio As Double, ByRef UseAutoCompSS As Boolean, ByRef fc0 As Double, ByRef An As Double, ByRef Bn As Double, ByRef UseAutoTensSS As Boolean, ByRef ft As Double, ByRef Ap As Double, ByRef UseZeroTens As Boolean, ByRef UseTensFractureEnergyAdjustment As Boolean, ByRef Gf As Double, Optional ByVal Temp As Double = 0.0) As Long

## Parameters

Name

The name of an existing concrete material property.

beta

The plasticity rate parameter.

fb0ratio

The biaxial compressive stress ratio.

UseAutoCompSS

If this item is True, the compressive stress-strain curve parameters are auto-generated based on the material's uniaxial stress-strain curve.

fc0

The compressive yield stress, used if UseAutoCompSS = False.

An

The compressive curve parameter An, used if UseAutoCompSS = False.

Bn

The compressive curve parameter Bn, used if UseAutoCompSS = False.

UseAutoTensSS

If this item is True, the tensile stress-strain curve parameters are auto-generated based on the material's uniaxial stress-strain curve.

ft

The tensile yield stress, used if UseAutoTensSS = False.

Ap

The tensile curve parameter Ap, used if UseAutoTensSS = False.

UseZeroTens

If this item is True, the tensile strength of the material is set to zero.

UseTensFractureEnergyAdjustment

If this item is True, the tensile curve is adjusted based on fracture energy.

Gf

The tensile fracture energy, used if UseTensFractureEnergyAdjustment = True.

Temp

This item applies only if the specified material has properties that are temperature dependent. That is, it applies only if properties are specified for the material at more than one temperature.

This item is the temperature at which the specified data is to be retrieved. The temperature must have been defined previously for the material.

## Remarks

This function gets the Faria Concrete Damage Plasticity parameters for a specified concrete material.  
The function returns zero if the assignment data is successfully obtained; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetFariaConcreteDamagePlasticityData()

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

          ret = SapModel.PropMaterial.SetMaterial("Concrete", eMatType.eMatType_Concrete, -1, "API example test", "Default")
 
   'set coupled material type data

          ret = SapModel.PropMaterial.SetCoupledModelType("Concrete", eMatCoupledType.eMatCoupledType_FariaConcreteDamagePlasticity)
 
   'set Faria Concrete Damage Plasticity data

          Dim beta As Double

          Dim fb0ratio As Double

          Dim UseAutoCompSS As Boolean

          Dim fc0 As Double

          Dim An As Double

          Dim Bn As Double

          Dim UseAutoTensSS As Boolean

          Dim ft As Double

          Dim Ap As Double

          Dim UseZeroTens As Boolean

          Dim UseTensFractureEnergyAdjustment As Boolean

          Dim Gf As Double

          ret = SapModel.PropMaterial. GetFariaConcreteDamagePlasticityParameters("Concrete", beta, fb0ratio, UseAutoCompSS, fc0, An, Bn, UseAutoTensSS, ft, Ap, UseZeroTens, UseTensFractureEnergyAdjustment, Gf)
 
   'close Sap2000

          SapObject.ApplicationExit False

          Set SapModel = Nothing

          Set SapObject = Nothing

    End Sub
```



## Release Notes

Initial release in version 27.1.0

## See Also

[GetCoupledModelType3D](GetCoupledModelType3D.md)

[SetCoupledModelType3D](SetCoupledModelType3D.md)

[SetFaria Concrete Damage Plasticity Parameters](SetFariaConcreteDamagePlasticityParameters.md)
