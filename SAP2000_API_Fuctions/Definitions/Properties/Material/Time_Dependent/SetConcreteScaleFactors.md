# SetConcreteScaleFactors

## Syntax

SapObject.SapModel.PropMaterial.TimeDep.SetConcreteScaleFactors

## VB6 Procedure

Function SetConcreteScaleFactors(ByVal Name As String, ByVal ScaleFactorAge As Double, ByVal ScaleFactorCreep As Double, ByVal ScaleFactorShrinkage As Double, Optional ByVal Temp As Double = 0) As Long

## Parameters

Name

The name of an existing concrete material property.

ScaleFactorAge

This value multiplies the stiffness (modulus of elasticity) computed with age for the material during a time-dependent analysis. It has no effect for load cases that do not consider time-dependent effects, or for materials that do not consider time-dependent age effects. The default value is unity, and the specified value must be positive.

ScaleFactorCreep

This value multiplies the creep coefficient, and hence the creep strain, computed for the material during a time-dependent analysis. It has no effect for load cases that do not consider time-dependent effects, or for materials that do not consider creep effects. The default value is unity, and the specified value must be positive.

ScaleFactorShrinkage

This value multiplies the shrinkage strain computed for the material during a time-dependent analysis. It has no effect for load cases that do not consider time-dependent effects, or for materials that do not consider shrinkage effects. The default value is unity, and the specified value must be positive.

Temp

This item applies only if the specified material has properties that are temperature dependent. That is, it applies only if properties are specified for the material at more than one temperature.

This item is the temperature at which the specified data is to be retrieved. The temperature must have been previously defined for the material.

## Remarks

This function sets scale factors for the time-dependent material property data for concrete materials. If this function is not called, default values of unity are assumed for all scale factors.

The function returns zero if the data is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub AssignMatPropConcreteTimeDepScaleFactors()

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

      ret = SapModel.PropMaterial.SetMaterial("Concrete", eMatType,Concrete)



   'assign CEB FIP-90 time dependent data

      ret = SapModel.PropMaterial.TimeDep.SetConcreteCEBFIP90("Concrete", True, True, True, 0.2, 40, 4, 4.9, 2, 1, 12)
 
   'assign time dependent scale factors

      ret = SapModel.PropMaterial.TimeDep.SetConcreteScaleFactors("Concrete", 1.0, 1.2, 1.1)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 17.2.0.

## See Also

[GetConcreteScaleFactors](GetConcreteScaleFactors.md)
